# -*- coding: utf-8 -*-
"""页签② 真实 HTTP 端到端：上传三文件+映射+cfg -> 跑 FBA 差异 -> 校验结果。
同时跑一次 legacy（无 cfg）回归。"""
import os, sys, json, time, tempfile, subprocess, urllib.request, urllib.parse
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = tempfile.mkdtemp(prefix='e2e_fba_')
PORT = 8021
BASE = 'http://127.0.0.1:%d' % PORT

def write_xlsx(path, header, rows):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(header)
    for r in rows: ws.append(r)
    wb.save(path)

def build():
    sales_p = os.path.join(TMP, 'sales.xlsx')
    write_xlsx(sales_p, ['店铺','ASIN','MSKU','状态','三天日均销量','七天日均销量','十四天日均销量','三十天日均销量','60天日均销量','90天日均销量'], [
        ['店铺A','ASIN1','M1','正常销售',10,10,10,10,5,5],
        ['店铺A','ASIN1','M2','正常销售',20,20,20,20,10,10],
        ['店铺B','ASIN2','M3','正常销售',5,5,5,0,0,0],
        ['店铺C','ASIN3','M4','正常销售',1,1,1,1,0,0],
        ['店铺D','ASIN4','M5','停止销售',9,9,9,9,9,9],
    ])
    map_p = os.path.join(TMP, 'map.xlsx')
    write_xlsx(map_p, ['店铺','ASIN','运营级别','补货方式'], [
        ['店铺A','ASIN1','L1','常规'],['店铺B','ASIN2','L2','常规'],
        ['店铺C','ASIN3','L3','常规'],['店铺D','ASIN4','L1','常规'],
    ])
    ops_p = os.path.join(TMP, 'ops.xlsx')
    write_xlsx(ops_p, ['店铺','ASIN','销售状态','配送类型','运营级别','店铺ASIN'], [
        ['店铺A','ASIN1','正常销售','FBA','L1','店铺AASIN1'],
        ['店铺B','ASIN2','正常销售','FBA','L2','店铺BASIN2'],
        ['店铺C','ASIN3','正常销售','FBA','L3','店铺CASIN3'],
        ['店铺D','ASIN4','正常销售','FBA','L1','店铺DASIN4'],
    ])
    fba_p = os.path.join(TMP, 'fba.xlsx')
    write_xlsx(fba_p, ['店铺ASIN','备货周期','日均销量(加权)','需求量'], [
        ['店铺AASIN1',30,12,100],['店铺AASIN1',45,15,200],
        ['店铺BASIN2',30,5,50],['店铺CASIN3',60,3,30],
        ['店铺XASIN9',30,99,999],
    ])
    return sales_p, map_p, ops_p, fba_p

CFG = {
    'mode':'rules',
    'rules':[
        {'name':'30天测试','type':'组合','weights':{3:0.5,7:0.2,14:0.2,30:0.1,60:0,90:0},'days':30},
        {'name':'45天测试','type':'组合','weights':{3:0.5,7:0.2,14:0.2,30:0.1,60:0,90:0},'days':45},
        {'name':'60天测试日均','type':'固定','fixed_window':60,'days':60},
    ],
    'bindings':{'L1':['30天测试','45天测试'],'L2':['30天测试'],'L3':['60天测试日均']},
    'default_rule':'30天测试','default_op_level':'',
}

def http(method, path, data=None, headers=None):
    req = urllib.request.Request(BASE+path, data=data, method=method)
    if headers:
        for k,v in headers.items(): req.add_header(k,v)
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.status, r.read().decode('utf-8')

def main():
    # 启动 server
    env = dict(os.environ)
    srv = subprocess.Popen([sys.executable, os.path.join(HERE,'server.py'), str(PORT)],
                           cwd=HERE, env=env)
    try:
        # 等 ping
        for _ in range(50):
            try:
                st, body = http('GET', '/api/ping')
                if st == 200: break
            except Exception: pass
            time.sleep(0.3)
        else:
            raise SystemExit('server 未启动')
        print('ping ok:', body.strip())

        sales_p, map_p, ops_p, fba_p = build()
        sid = 'E2E'

        def upload(sid_, slot, path):
            with open(path,'rb') as f: raw = f.read()
            st, body = http('POST', '/api/upload?sid=%s&slot=%s' % (sid_, slot), data=raw,
                            headers={'Content-Type':'application/octet-stream'})
            assert st == 200, 'upload %s -> %s' % (slot, body)
            print('upload', slot, 'ok')
        upload(sid, 'sales', sales_p); upload(sid, 'mapping', map_p); upload(sid, 'ops', ops_p); upload(sid, 'fba', fba_p)

        # ---- rules 模式 ----
        cfg_q = urllib.parse.urlencode({'cfg': json.dumps(CFG, ensure_ascii=False)})
        st, body = http('POST', '/api/tasks?type=fba&sid=%s&%s' % (sid, cfg_q))
        assert st == 202, 'run fba -> %s' % body
        tid = json.loads(body)['id']
        print('fba task id:', tid)
        res = None
        for _ in range(60):
            st, body = http('GET', '/api/tasks/%s' % tid)
            res = json.loads(body)
            if res['status'] in ('done','error'): break
            time.sleep(0.3)
        assert res['status'] == 'done', 'fba failed: %s' % res.get('error_msg')
        print('=== RULES (HTTP) ===')
        print('summary:', res['summary'])
        print('diff_header:', res['diff_header'])
        for r in res['diff_preview']: print('  ', r)
        s = res['summary']
        assert s['use_day'] is True, '应启用按天比对'
        assert s['both'] == 3, 'both 应=3，实际 %s' % s['both']
        assert s['extra'] == 2, 'extra 应=2，实际 %s' % s['extra']
        a_cnt = sum(1 for r in res['diff_preview'] if r[0]=='店铺AASIN1')
        assert a_cnt == 2, '店铺AASIN1 应展开2天，实际 %d' % a_cnt
        c = [r for r in res['diff_preview'] if r[0]=='店铺CASIN3' and str(r[7])=='60']
        assert c and c[0][1].startswith('多余') and '该规则加权日均=0' in c[0][11], '店铺C 60天 应标注规则不满足'
        print('RULES (HTTP) 断言: PASS')

        # ---- legacy 模式（上传新 sid，不传 cfg） ----
        sid2 = 'E2L'
        upload(sid2, 'sales', sales_p); upload(sid2, 'ops', ops_p); upload(sid2, 'fba', fba_p)
        st, body = http('POST', '/api/tasks?type=fba&sid=%s' % sid2)
        assert st == 202, body
        tid2 = json.loads(body)['id']
        res2 = None
        for _ in range(60):
            st, body = http('GET', '/api/tasks/%s' % tid2)
            res2 = json.loads(body)
            if res2['status'] in ('done','error'): break
            time.sleep(0.3)
        assert res2['status'] == 'done', res2.get('error_msg')
        s2 = res2['summary']
        print('\n=== LEGACY (HTTP) ===')
        print('summary:', s2)
        assert s2['use_day'] is False, 'legacy 不应按天比对'
        assert s2['both'] == 3, 'legacy both 应=3，实际 %s' % s2['both']
        assert s2['extra'] == 1, 'legacy extra 应=1，实际 %s' % s2['extra']
        print('LEGACY (HTTP) 断言: PASS')
        print('\nHTTP 端到端全部通过 ✅')
    finally:
        srv.terminate()
        try: srv.wait(timeout=5)
        except Exception: srv.kill()

if __name__ == '__main__':
    main()
