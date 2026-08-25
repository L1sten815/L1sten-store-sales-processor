# -*- coding: utf-8 -*-
"""页签② rules 模式：按 (店铺ASIN,天数) 展开+比对 + legacy 兼容 回归测试。"""
import os, sys, json, tempfile
import openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import server

TMP = tempfile.mkdtemp(prefix='fb rules_')

def write_xlsx(path, header, rows):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(path)

def build():
    # 销量（MSKU 级）
    sales_p = os.path.join(TMP, 'sales.xlsx')
    sh = ['店铺', 'ASIN', 'MSKU', '状态', '三天日均销量', '七天日均销量', '十四天日均销量', '三十天日均销量', '60天日均销量', '90天日均销量']
    srows = [
        ['店铺A', 'ASIN1', 'M1', '正常销售', 10, 10, 10, 10, 5, 5],
        ['店铺A', 'ASIN1', 'M2', '正常销售', 20, 20, 20, 20, 10, 10],
        ['店铺B', 'ASIN2', 'M3', '正常销售', 5, 5, 5, 0, 0, 0],
        ['店铺C', 'ASIN3', 'M4', '正常销售', 1, 1, 1, 1, 0, 0],   # 60天=0 → 固定60规则 M=0
        ['店铺D', 'ASIN4', 'M5', '停止销售', 9, 9, 9, 9, 9, 9],   # 非正常销售
    ]
    write_xlsx(sales_p, sh, srows)

    # 映射表
    map_p = os.path.join(TMP, 'map.xlsx')
    write_xlsx(map_p, ['店铺', 'ASIN', '运营级别', '补货方式'], [
        ['店铺A', 'ASIN1', 'L1', '常规'],
        ['店铺B', 'ASIN2', 'L2', '常规'],
        ['店铺C', 'ASIN3', 'L3', '常规'],
        ['店铺D', 'ASIN4', 'L1', '常规'],
    ])

    # 运营级别
    ops_p = os.path.join(TMP, 'ops.xlsx')
    oh = ['店铺', 'ASIN', '销售状态', '配送类型', '运营级别', '店铺ASIN']
    orows = [
        ['店铺A', 'ASIN1', '正常销售', 'FBA', 'L1', '店铺AASIN1'],
        ['店铺B', 'ASIN2', '正常销售', 'FBA', 'L2', '店铺BASIN2'],
        ['店铺C', 'ASIN3', '正常销售', 'FBA', 'L3', '店铺CASIN3'],
        ['店铺D', 'ASIN4', '正常销售', 'FBA', 'L1', '店铺DASIN4'],
    ]
    write_xlsx(ops_p, oh, orows)

    # FBA需求（含 备货周期 列；同店铺ASIN多行）
    fba_p = os.path.join(TMP, 'fba.xlsx')
    fh = ['店铺ASIN', '备货周期', '日均销量(加权)', '需求量']
    frows = [
        ['店铺AASIN1', 30, 12, 100],
        ['店铺AASIN1', 45, 15, 200],
        ['店铺BASIN2', 30, 5, 50],
        ['店铺CASIN3', 60, 3, 30],
        ['店铺XASIN9', 30, 99, 999],   # 多余（ops 无）
    ]
    write_xlsx(fba_p, fh, frows)
    return sales_p, map_p, ops_p, fba_p

CFG = {
    'mode': 'rules',
    'rules': [
        {'name': '30天测试', 'type': '组合', 'weights': {3: 0.5, 7: 0.2, 14: 0.2, 30: 0.1}, 'days': 30},
        {'name': '45天测试', 'type': '组合', 'weights': {3: 0.5, 7: 0.2, 14: 0.2, 30: 0.1}, 'days': 45},
        {'name': '60天测试日均', 'type': '固定', 'fixed_window': 60, 'days': 60},
    ],
    'bindings': {'L1': ['30天测试', '45天测试'], 'L2': ['30天测试'], 'L3': ['60天测试日均']},
    'default_rule': '30天测试', 'default_op_level': '',
}

def run_rules():
    sales_p, map_p, ops_p, fba_p = build()
    mp = server.load_mapping(map_p)
    out_s = os.path.join(TMP, 'o_sales.xlsx')
    out_o = os.path.join(TMP, 'o_ops.xlsx')
    out_d = os.path.join(TMP, 'o_diff.xlsx')
    # 页签②已改为：销量原始文件不再单独处理，直接作为 lookup 源传给 process_ops
    h_o, n_o, meet, li = server.process_ops(ops_p, sales_p, out_o, 0.6, 0.3, 0.14, cfg=CFG, mapping=mp)
    h_d, n_d, summary, prev = server.process_diff(out_o, fba_p, out_d, cfg=CFG, mapping=mp)
    print('=== RULES 模式 ===')
    print('ops 行数:', n_o, '| meet:', meet)
    print('ops 表头:', h_o)
    print('diff 表头:', h_d)
    print('summary:', summary)
    print('diff 预览:')
    for r in prev:
        print('  ', r)
    # 断言
    assert n_o == 6, 'ops 应有 6 行(店铺A 2行 + 店铺D 2行 + 店铺B 1 + 店铺C 1)，实际 %d' % n_o
    # 店铺AASIN1 应出现 2 次（30天、45天）
    a_rows = [r for r in prev if r[0] == '店铺AASIN1']
    assert len(a_rows) == 2, '店铺AASIN1 应展开为 2 行，实际 %d' % len(a_rows)
    days = sorted(int(r[7]) for r in a_rows)
    assert days == [30, 45], '店铺AASIN1 天数应为 [30,45]，实际 %s' % days
    # 店铺CASIN3 L3 绑 60天固定，d60=0 → 不满足条件，备注 该规则加权日均=0
    c_rows = [r for r in prev if r[0] == '店铺CASIN3']
    assert len(c_rows) == 1 and c_rows[0][5] == '否', '店铺CASIN3 应不满足条件'
    assert '该规则加权日均=0' in c_rows[0][11], '店铺CASIN3 备注应写清规则不满足，实际 %r' % c_rows[0][11]
    # diff: use_day 应为 True
    assert summary['use_day'] is True, '应启用按天比对'
    assert summary['both'] == 3, '已生成应为3，实际 %s' % summary['both']
    assert summary['extra'] == 2, '多余应为2(店铺XASIN9 + 店铺CASIN3 的60天FBA无对应满足)，实际 %s' % summary['extra']
    # 店铺CASIN3 diff（60天）应为 多余，备注写清规则
    cdiff = [r for r in prev if r[0] == '店铺CASIN3' and str(r[7]) == '60']
    assert cdiff and cdiff[0][1].startswith('多余'), '店铺CASIN3 60天 应为多余'
    assert '该规则加权日均=0' in cdiff[0][11], 'diff 备注应标注规则不满足'
    print('RULES 断言: PASS')

def run_legacy():
    sales_p, map_p, ops_p, fba_p = build()
    out_s = os.path.join(TMP, 'l_sales.xlsx')
    out_o = os.path.join(TMP, 'l_ops.xlsx')
    out_d = os.path.join(TMP, 'l_diff.xlsx')
    # 页签② legacy：销量原始文件直接 lookup
    h_o, n_o, meet, li = server.process_ops(ops_p, sales_p, out_o, 0.6, 0.3, 0.14)
    h_d, n_d, summary, prev = server.process_diff(out_o, fba_p, out_d)
    print('\n=== LEGACY 模式 ===')
    print('ops 行数:', n_o, '| meet:', meet)
    print('summary:', summary)
    print('ops 表头:', h_o)
    print('diff 表头:', h_d)
    for r in prev:
        print('  ', r)
    assert n_o == 4, 'legacy ops 应单行(4店铺ASIN)，实际 %d' % n_o
    assert summary['use_day'] is False, 'legacy 不应按天比对'
    assert summary['both'] == 3, 'legacy 已生成应为3(店铺A/店铺B/店铺C满足；店铺D销量=0跳过；店铺X多余)，实际 %s' % summary['both']
    assert summary['extra'] == 1, 'legacy 多余应为1(店铺XASIN9)，实际 %s' % summary['extra']
    # legacy ops 表头不应含 规则/天数/备注
    assert '规则' not in h_o and '天数' not in h_o, 'legacy ops 不应有 规则/天数 列'
    print('LEGACY 断言: PASS')

if __name__ == '__main__':
    run_rules()
    run_legacy()
    print('\n全部断言通过 ✅')
