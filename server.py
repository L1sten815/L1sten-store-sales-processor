#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
产品化服务：上传 -> 后台任务 -> 下载
- POST /api/tasks?type=sales        创建销量处理任务并上传文件（请求体=原始 xlsx 字节）
- POST /api/upload?sid=&slot=        暂存 FBA 差异分析所需的单个文件（销量/运营级别/FBA需求）
- POST /api/tasks?type=fba&sid=      链式处理：销量结果 -> 运营级别结果 -> 差异比对（产出 3 份结果）
- GET  /api/tasks/<id>               查询任务状态/行数/错误/预览/汇总
- GET  /api/tasks/<id>/download?which=sales|ops|diff  下载结果 xlsx
- GET  /                          打开处理页面
启动：python server.py [端口]，默认 8000
"""
import http.server, threading, os, json, uuid, time, sys, re, urllib.parse, zipfile, inspect
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
# 端口：优先读环境变量 $PORT（云平台的标配），其次命令行参数，最后默认 8000
PORT = int(os.environ.get('PORT', sys.argv[1] if len(sys.argv) > 1 else 8000))
DEF = (0.6, 0.3, 0.1)
# 销量处理列名默认（可经前端高级配置覆盖）
DEF_COLS = {'store': '店铺', 'asin': 'ASIN', 'd3': '三天日均销量', 'd7': '七天日均销量', 'd14': '十四天日均销量'}
# 运营级别处理列名默认
DEF_OPS_COLS = {'store': '店铺', 'asin': 'ASIN', 'status': '销售状态', 'dtype': '配送类型',
                'asinKey': '店铺ASIN', 'salesKey': '日均销量加权', 'level': '运营级别'}

TASKS = {}      # id -> dict
UPLOADS = {}    # sid -> {slot: path}
LOCK = threading.Lock()


def fnum(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def col_letter(idx):
    # 0-based -> Excel 列字母（A, B, ... Z, AA, ...）
    s = ''; idx += 1
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _write_xlsx(out_path, sheet_tmp, sheet_name='sheet1'):
    """把流式生成的 sheet XML 打包成最小可用的 xlsx（inline strings，无共享字符串表）。"""
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          '<sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets></workbook>' % sheet_name)
    wbres = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
             '</Relationships>')
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels)
        z.writestr('xl/workbook.xml', wb)
        z.writestr('xl/_rels/workbook.xml.rels', wbres)
        # 工作表 XML 占文件体积 99%，用压缩等级 1（最快）：对高度重复的 XML 仍
        # 能压到 ~1.3x，但比默认等级 6 快 4~5 倍。小文件用默认即可。
        z.write(sheet_tmp, 'xl/worksheets/sheet1.xml', compresslevel=1)


def _sheet_tmp():
    return os.path.join(HERE, '_sheet_%d.tmp' % os.getpid())


def _stream_sheet(out_tmp, header, row_iter):
    """把表头 + 逐行数据流式拼成 worksheet XML（inline strings，无共享字符串表）。"""
    from xml.sax.saxutils import escape as _xesc
    def esc(s):
        return _xesc(str(s), {'"': '&quot;'})
    with open(out_tmp, 'w', encoding='utf-8', newline='') as f:
        f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')
        f.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
        def write_row(rownum, cells):
            parts = ['<row r="%d">' % rownum]
            for ci, v in enumerate(cells):
                ref = col_letter(ci) + str(rownum)
                if v is None or v == '':
                    parts.append('<c r="%s"/>' % ref)
                elif isinstance(v, bool):
                    parts.append('<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' % (ref, esc(v)))
                elif isinstance(v, float):
                    parts.append('<c r="%s"><v>%s</v></c>' % (ref, repr(v)))
                elif isinstance(v, int):
                    parts.append('<c r="%s"><v>%d</v></c>' % (ref, v))
                else:
                    parts.append('<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, esc(v)))
            parts.append('</row>')
            f.write(''.join(parts))
        write_row(1, header)
        rownum = 2
        for row in row_iter:
            write_row(rownum, row)
            rownum += 1
        f.write('</sheetData></worksheet>')


def _calamine_rows(path):
    """优先用 calamine（Rust 解析，惰性迭代，内存恒定）；失败回退 openpyxl read_only。"""
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(path)
    sheet = wb.get_sheet_by_index(0)
    it = sheet.iter_rows()                      # 惰性迭代器，首行即表头
    header = [str(h) if h is not None else '' for h in next(it)]
    if not header:
        raise ValueError("文件没有任何数据行（只有表头或为空）")
    return header, it


def process_file(in_path, w3, w7, w14, out_path, cols=None, on_progress=None):
    """销量处理：原始销量 -> 店铺ASIN + 加权日均销量（in_memory 计算，流式写出）。"""
    C = dict(DEF_COLS)
    if cols:
        for k, v in cols.items():
            if v:
                C[k] = v
    WSUM = w3 + w7 + w14
    if WSUM == 0:
        raise ValueError("三个权重不能同时为 0")

    try:
        header, data_iter = _calamine_rows(in_path)
    except Exception:
        wb_r = openpyxl.load_workbook(in_path, read_only=True, data_only=True)
        ws_r = wb_r.active
        it = ws_r.iter_rows(values_only=True)
        header = [str(h) if h is not None else '' for h in next(it)]
        if not header:
            raise ValueError("文件没有任何数据行（只有表头或为空）")
        data_iter = it

    def find(n):
        for i, h in enumerate(header):
            if h == n:
                return i
        return -1
    i_asin = find(C['asin']); i_store = find(C['store'])
    i_d3 = find(C['d3']); i_d7 = find(C['d7']); i_d14 = find(C['d14'])
    names = [('店铺', C['store']), ('ASIN', C['asin']),
             ('三天日均销量', C['d3']), ('七天日均销量', C['d7']), ('十四天日均销量', C['d14'])]
    miss = [label + '(' + name + ')' for label, name in names if find(name) < 0]
    if miss:
        raise ValueError("缺少必要列：" + "、".join(miss) + "。当前文件的列名为：" + "、".join(header))

    new_header = list(header)
    new_header.insert(0, '店铺ASIN')
    new_header.insert(i_d14 + 2, '日均销量加权')
    hdr_len = len(new_header)

    def get(row, idx):
        if 0 <= idx < len(row):
            return row[idx]
        return None

    preview_rows = []
    n = 0

    def row_gen():
        nonlocal n
        for row in data_iter:
            store = get(row, i_store); asin = get(row, i_asin)
            a = (str(store) if store is not None else '') + (str(asin) if asin is not None else '')
            m = (w3 * fnum(get(row, i_d3)) + w7 * fnum(get(row, i_d7)) + w14 * fnum(get(row, i_d14))) / WSUM
            m = round(m * 10000) / 10000
            nr = list(row); nr.insert(0, a); nr.insert(i_d14 + 2, m)
            while len(nr) < hdr_len:
                nr.append('')
            if len(preview_rows) < 200:
                preview_rows.append(nr)
            yield nr
            n += 1
            if on_progress and n % 5000 == 0:
                on_progress(n)

    tmp = _sheet_tmp()
    _stream_sheet(tmp, new_header, row_gen())
    _write_xlsx(out_path, tmp, 'sheet1')
    try:
        os.remove(tmp)
    except OSError:
        pass
    if on_progress:
        on_progress(n)
    return new_header, n, preview_rows


def process_ops(in_ops, in_sales, out_path, cols=None, on_progress=None, levels=None):
    """运营级别处理：按店铺ASIN 从销量结果取加权日均销量(L)，标记满足条件。
    输出表头=[店铺ASIN]+原表头+[销量,满足条件]。
    """
    C = dict(DEF_OPS_COLS)
    if cols:
        for k, v in cols.items():
            if v:
                C[k] = v

    # 1) 读销量结果 -> {店铺ASIN: 加权日均销量(按 MSKU 聚合)}
    #    同 店铺ASIN 可能对应多个 MSKU（同一 ASIN 的不同商户 SKU），需把各 MSKU 的
    #    加权日均销量 累加，得到该 ASIN 的总销量；只统计「正常销售」的 MSKU。
    sh, sit = _calamine_rows(in_sales) if _try_calamine(in_sales) else _openpyxl_rows(in_sales)
    def fidx(hdr, name):
        for i, h in enumerate(hdr):
            if h == name:
                return i
        return -1
    iA = fidx(sh, C['asinKey']); iM = fidx(sh, C['salesKey'])
    iStat = fidx(sh, '状态')   # 销量表的 状态 列（正常销售/停止销售）
    if iA < 0 or iM < 0:
        raise ValueError("销量结果缺少必要列：店铺ASIN / 日均销量加权（当前列：%s）" % sh)
    sales = {}
    for r in sit:
        a = r[iA]
        if a is None:
            continue
        a = str(a)
        # 仅汇总「正常销售」MSKU 的加权日均销量（无 状态 列时汇总全部，避免漏算）
        if iStat >= 0 and r[iStat] is not None and str(r[iStat]).strip() != '正常销售':
            continue
        sales[a] = sales.get(a, 0.0) + fnum(r[iM])

    # 2) 读运营级别
    oh, oit = _calamine_rows(in_ops) if _try_calamine(in_ops) else _openpyxl_rows(in_ops)
    i_store = fidx(oh, C['store']); i_asin = fidx(oh, C['asin'])
    i_status = fidx(oh, C['status']); i_dtype = fidx(oh, C['dtype'])
    i_key = fidx(oh, C['asinKey'])
    if i_status < 0 or i_dtype < 0:
        raise ValueError("运营级别缺少必要列：销售状态 / 配送类型（当前列：%s）" % oh)
    # 运营级别（白名单）：用户填入的「符合补货条件」的级别；为空表示不按运营级别筛选
    i_level = fidx(oh, C['level'])
    allowed = set()
    if levels:
        items = levels if isinstance(levels, (list, tuple, set)) else [levels]
        for lv in items:
            if not lv:
                continue
            for part in re.split(r'[,\s;；、]+', str(lv).strip()):
                part = part.strip()
                if part:
                    allowed.add(part)
    if allowed and i_level < 0:
        raise ValueError("运营级别缺少必要列：运营级别（当前列：%s）" % oh)
    distinct_levels = set()

    # 清洗表头：去掉名为空的列（用户遗留的空 M 列）；若原表已有 店铺ASIN/销量 则不重复添加
    clean_idx = [i for i, h in enumerate(oh) if h != '']
    clean_oh = [oh[i] for i in clean_idx]
    i_key2 = fidx(clean_oh, '店铺ASIN')
    iL2 = fidx(clean_oh, '销量')
    prepend_a = (i_key2 < 0)

    out_header = list(clean_oh)
    if prepend_a:
        out_header = ['店铺ASIN'] + out_header
    if iL2 < 0:
        out_header = out_header + ['销量', '满足条件']
    else:
        out_header = out_header + ['满足条件']
    sales_idx = (iL2 + (1 if prepend_a else 0)) if iL2 >= 0 else None

    n = 0
    meet = 0

    def row_gen():
        nonlocal n, meet
        for r in oit:
            store = r[i_store] if i_store >= 0 else ''
            asin = r[i_asin] if i_asin >= 0 else ''
            a = (str(store) if store is not None else '') + (str(asin) if asin is not None else '')
            if i_key >= 0 and r[i_key] is not None:
                a = str(r[i_key])
            L = sales.get(a, 0.0)
            level_val = str(r[i_level]).strip() if (i_level >= 0 and r[i_level] is not None) else ''
            if level_val:
                distinct_levels.add(level_val)
            level_ok = (not allowed) or (level_val in allowed)
            ok = (str(r[i_status]).strip() == '正常销售') and (str(r[i_dtype]).strip() == 'FBA') and L > 0 and level_ok
            if ok:
                meet += 1
            cr = [r[i] for i in clean_idx]
            if prepend_a:
                cr = [a] + cr
            if iL2 < 0:
                cr = cr + [round(L * 10000) / 10000]      # 原表无销量列 -> 追加
            else:
                cr[sales_idx] = round(L * 10000) / 10000   # 覆盖原销量列为当前关联值
            cr = cr + ['是' if ok else '否']
            yield cr
            n += 1
            if on_progress and n % 5000 == 0:
                on_progress(n)

    tmp = _sheet_tmp()
    _stream_sheet(tmp, out_header, row_gen())
    _write_xlsx(out_path, tmp, 'sheet1')
    try:
        os.remove(tmp)
    except OSError:
        pass
    if on_progress:
        on_progress(n)
    level_info = {'distinct': sorted(distinct_levels), 'allowed': sorted(allowed)}
    return out_header, n, meet, level_info


def process_diff(in_ops_result, in_fba, out_path, cols=None, on_progress=None, preview_cap=200):
    """差异比对：运营级别满足条件(唯一店铺ASIN) vs FBA需求(唯一店铺ASIN)。
    输出每行一个店铺ASIN + 差异状态 + 两侧上下文；返回 (header, n, summary, preview)。
    """
    # 1) 读运营级别结果 -> S_op + 上下文
    oh, oit = _calamine_rows(in_ops_result) if _try_calamine(in_ops_result) else _openpyxl_rows(in_ops_result)
    def fidx(hdr, name):
        for i, h in enumerate(hdr):
            if h == name:
                return i
        return -1
    iA = fidx(oh, '店铺ASIN'); iStore = fidx(oh, '店铺'); iAsin = fidx(oh, 'ASIN')
    iL = fidx(oh, '销量'); iOk = fidx(oh, '满足条件')
    if iA < 0:
        raise ValueError("运营级别结果缺少 店铺ASIN 列（当前列：%s）" % oh)
    S_op = set()
    op_ctx = {}
    for r in oit:
        a = str(r[iA]) if r[iA] is not None else ''
        if not a:
            continue
        if iOk >= 0:
            ok = str(r[iOk]).strip() == '是'
        else:
            ok = False
        if ok:
            S_op.add(a)
        if a not in op_ctx:
            store = r[iStore] if iStore >= 0 else ''
            asin = r[iAsin] if iAsin >= 0 else ''
            L = fnum(r[iL]) if iL >= 0 else 0.0
            op_ctx[a] = (str(store) if store is not None else '', str(asin) if asin is not None else '', L, ok)

    # 2) 读 FBA需求 -> S_fba + 上下文（按店铺ASIN 聚合）
    fh, fit = _calamine_rows(in_fba) if _try_calamine(in_fba) else _openpyxl_rows(in_fba)
    jA = fidx(fh, '店铺ASIN'); jStore = fidx(fh, '店铺'); jAsin = fidx(fh, 'ASIN')
    jD = fidx(fh, '日均销量(加权)'); jM = fidx(fh, '需求量')
    # 原表的 店铺ASIN 是用户后加的；没有就用「店铺」+「ASIN」现拼（与销量表的 A 列同一套算法）。
    if jA < 0 and (jStore < 0 or jAsin < 0):
        raise ValueError("FBA需求缺少 店铺ASIN 列，或缺少 店铺+ASIN 列（当前列：%s）" % fh)
    compute_key = (jA < 0)
    fba_ctx = {}
    for r in fit:
        if compute_key:
            a = (str(r[jStore]) if r[jStore] is not None else '') + (str(r[jAsin]) if r[jAsin] is not None else '')
        else:
            a = str(r[jA]) if r[jA] is not None else ''
        if not a:
            continue
        d = fnum(r[jD]) if jD >= 0 else 0.0
        q = fnum(r[jM]) if jM >= 0 else 0.0
        fba_ctx.setdefault(a, []).append((q, d))

    S_fba = set(fba_ctx.keys())
    order = {'缺失(运营满足条件、FBA需求未生成)': 0, '多余(FBA需求有、运营不满足条件)': 1, '已生成': 2}

    def status_of(a):
        in_op = a in S_op
        in_fba = a in S_fba
        if in_op and in_fba:
            return '已生成'
        if in_op:
            return '缺失(运营满足条件、FBA需求未生成)'
        return '多余(FBA需求有、运营不满足条件)'

    rows = []
    preview = []
    for a in (S_op | S_fba):
        st = status_of(a)
        oc = op_ctx.get(a)
        fc = fba_ctx.get(a, [])
        store = oc[0] if oc else ''
        asin = oc[1] if oc else ''
        L = oc[2] if oc else ''
        ok = oc[3] if oc else ''
        nc = len(fc)
        qsum = round(sum(x[0] for x in fc), 2)
        dfirst = fc[0][1] if fc else ''
        note = ''
        row = [a, st, store, asin, L, ('是' if ok else '否') if oc else '', nc, qsum, dfirst, note]
        rows.append(row)
        if len(preview) < preview_cap:
            preview.append(row)
    rows.sort(key=lambda x: (order.get(x[1], 9), x[0]))

    miss = sum(1 for x in rows if x[1].startswith('缺失'))
    extra = sum(1 for x in rows if x[1].startswith('多余'))
    both = sum(1 for x in rows if x[1] == '已生成')

    out_header = ['店铺ASIN', '差异状态', '运营-店铺', '运营-ASIN', '运营-销量', '运营满足条件',
                  'FBA需求-行数', 'FBA需求-需求量合计', 'FBA需求-日均销量(加权)首行', '备注']
    tmp = _sheet_tmp()
    _stream_sheet(tmp, out_header, iter(rows))
    _write_xlsx(out_path, tmp, 'sheet1')
    try:
        os.remove(tmp)
    except OSError:
        pass
    summary = {'op_meet': len(S_op), 'fba_unique': len(S_fba),
               'miss': miss, 'extra': extra, 'both': both, 'total': len(rows)}
    return out_header, len(rows), summary, preview


def _try_calamine(path):
    try:
        from python_calamine import CalamineWorkbook
        CalamineWorkbook.from_path(path)
        return True
    except Exception:
        return False


def _openpyxl_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    header = [str(h) if h is not None else '' for h in next(it)]
    if not header:
        raise ValueError("文件没有任何数据行（只有表头或为空）")
    return header, it


class H(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass  # 静默

    def _send(self, code, body, ctype, extra=None):
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Access-Control-Allow-Origin', '*')
        if extra:
            for k, v in extra.items():
                self.send_header(k, v)
        self.end_headers()
        if isinstance(body, (bytes, bytearray)):
            self.wfile.write(body)
        else:
            self.wfile.write(body.encode('utf-8'))

    def _json(self, code, d):
        self._send(code, json.dumps(d, ensure_ascii=False), 'application/json')

    def do_OPTIONS(self):
        # CORS 预检：处理跨源 POST（file:// 或非同源页面访问时浏览器会先发 OPTIONS）
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', '*')
        self.send_header('Access-Control-Max-Age', '86400')
        self.end_headers()

    def do_GET(self):
        p = urllib.parse.urlparse(self.path)
        if p.path == '/api/ping':
            self._json(200, {'ok': True, 'name': 'store-sales-processor'})
            return
        if p.path in ('/', '/index.html'):
            try:
                data = open(os.path.join(HERE, 'store-sales-processor.html'), 'rb').read()
                self._send(200, data, 'text/html; charset=utf-8')
            except FileNotFoundError:
                self._send(404, '页面文件 store-sales-processor.html 未找到', 'text/plain; charset=utf-8')
            return
        m = re.match(r'^/api/tasks/([^/]+)/download$', p.path)
        if m:
            tid = m.group(1)
            which = qs_get(p, 'which') or ''
            with LOCK:
                t = TASKS.get(tid)
            if not t:
                self._json(404, {'error': 'task not found'})
                return
            if which and t.get('outputs'):
                path = t['outputs'].get(which)
            else:
                path = t.get('output_path')
            if not path or not os.path.exists(path):
                self._send(404, 'result not ready', 'text/plain')
                return
            data = open(path, 'rb').read()
            fname = {'sales': 'sales_result.xlsx', 'ops': 'ops_result.xlsx', 'diff': 'diff_result.xlsx'}.get(which, 'store_sales_result.xlsx')
            disp = "attachment; filename=\"%s\"; filename*=UTF-8''%s" % (fname, fname)
            self._send(200, data,
                       'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                       {'Content-Disposition': disp})
            return
        m = re.match(r'^/api/tasks/([^/]+)$', p.path)
        if m:
            tid = m.group(1)
            with LOCK:
                t = TASKS.get(tid)
            if not t:
                self._json(404, {'error': 'task not found'})
                return
            self._json(200, {
                'id': tid,
                'status': t['status'],
                'filename': t.get('filename'),
                'n_rows': t.get('n_rows'),
                'header': t.get('header'),
                'preview_rows': t.get('preview_rows'),
                'error_msg': t.get('error_msg'),
                'type': t.get('type'),
                'summary': t.get('summary'),
                'has_output': bool(t.get('output_path') and os.path.exists(t['output_path'])),
                'has_outputs': bool(t.get('outputs')),
                'diff_header': t.get('diff_header'),
                'diff_preview': t.get('diff_preview'),
                'level_info': t.get('level_info'),
                'created': t.get('created'),
                'finished': t.get('finished'),
            })
            return
        if p.path == '/api/tasks':
            with LOCK:
                lst = [{'id': tid, 'status': t['status'], 'filename': t.get('filename'),
                        'n_rows': t.get('n_rows'), 'type': t.get('type')} for tid, t in TASKS.items()]
            self._json(200, {'tasks': lst})
            return
        self._send(404, 'not found', 'text/plain')

    def do_POST(self):
        p = urllib.parse.urlparse(self.path)
        qs = urllib.parse.parse_qs(p.query)

        # ---- 文件暂存（FBA 差异分析需要多文件） ----
        if p.path == '/api/upload':
            sid = qs_get(p, 'sid') or ''
            slot = qs_get(p, 'slot') or ''
            try:
                length = int(self.headers.get('Content-Length', 0))
            except ValueError:
                length = 0
            data = self.rfile.read(length) if length else b''
            if not sid or not slot:
                self._json(400, {'error': 'need sid and slot'})
                return
            d = UPLOADS.setdefault(sid, {})
            path = os.path.join(HERE, '_up_%s_%s.xlsx' % (sid, slot))
            open(path, 'wb').write(data)
            d[slot] = path
            self._json(200, {'ok': True, 'slot': slot, 'size': len(data)})
            return

        if p.path != '/api/tasks':
            self._send(404, 'not found', 'text/plain')
            return

        # ---- FBA 差异分析（链式：销量->运营级别->差异） ----
        typ = qs_get(p, 'type') or 'sales'
        if typ == 'fba':
            sid = qs_get(p, 'sid') or ''
            try:
                w3 = float(qs_get(p, 'w3') or DEF[0])
                w7 = float(qs_get(p, 'w7') or DEF[1])
                w14 = float(qs_get(p, 'w14') or DEF[2])
            except ValueError:
                self._json(400, {'error': 'invalid weights'})
                return
            up = UPLOADS.get(sid, {})
            if not (up.get('sales') and up.get('ops') and up.get('fba')):
                self._json(400, {'error': '请先上传 销量 / 运营级别 / FBA需求 三个文件'})
                return
            cols = {
                'store': (qs_get(p, 'colStore') or '').strip(),
                'asin': (qs_get(p, 'colAsin') or '').strip(),
                'd3': (qs_get(p, 'colD3') or '').strip(),
                'd7': (qs_get(p, 'colD7') or '').strip(),
                'd14': (qs_get(p, 'colD14') or '').strip(),
            }
            if not any(cols.values()):
                cols = None
            # 符合补货条件的运营级别（白名单），可重复传多个 levels= 参数
            levels = []
            for raw in qs_get_list(p, 'levels'):
                for part in re.split(r'[,\s;；、]+', raw):
                    part = part.strip()
                    if part:
                        levels.append(part)
            tid = uuid.uuid4().hex[:12]
            task = {
                'id': tid, 'status': 'running', 'type': 'fba', 'filename': 'fba',
                'n_rows': 0, 'header': None, 'preview_rows': None,
                'outputs': None, 'summary': None, 'diff_header': None, 'diff_preview': None,
                'error_msg': None, 'w': (w3, w7, w14), 'cols': cols, 'levels': levels, 'sid': sid,
                'created': time.time(), 'finished': None,
            }
            with LOCK:
                TASKS[tid] = task
            threading.Thread(target=run_fba, args=(sid, task), daemon=True).start()
            self._json(202, {'id': tid, 'status': 'running', 'n_rows': 0})
            return

        # ---- 销量处理（单文件） ----
        try:
            w3 = float(qs_get(p, 'w3') or DEF[0])
            w7 = float(qs_get(p, 'w7') or DEF[1])
            w14 = float(qs_get(p, 'w14') or DEF[2])
        except ValueError:
            self._json(400, {'error': 'invalid weights'})
            return
        try:
            length = int(self.headers.get('Content-Length', 0))
        except ValueError:
            length = 0
        body = self.rfile.read(length) if length else b''
        fname = (qs_get(p, 'filename') or self.headers.get('X-Filename') or 'upload.xlsx')
        cols = {
            'store': (qs_get(p, 'colStore') or '').strip(),
            'asin': (qs_get(p, 'colAsin') or '').strip(),
            'd3': (qs_get(p, 'colD3') or '').strip(),
            'd7': (qs_get(p, 'colD7') or '').strip(),
            'd14': (qs_get(p, 'colD14') or '').strip(),
        }
        if not any(cols.values()):
            cols = None
        tid = uuid.uuid4().hex[:12]
        in_path = os.path.join(HERE, '_in_' + tid + '.xlsx')
        out_path = os.path.join(HERE, '_out_' + tid + '.xlsx')
        open(in_path, 'wb').write(body)
        task = {
            'id': tid, 'status': 'running', 'type': 'sales', 'filename': fname, 'n_rows': 0,
            'header': None, 'preview_rows': None, 'error_msg': None,
            'input_path': in_path, 'output_path': out_path,
            'w': (w3, w7, w14), 'cols': cols,
            'created': time.time(), 'finished': None,
        }
        with LOCK:
            TASKS[tid] = task

        def run():
            try:
                def prog(n):
                    with LOCK:
                        task['n_rows'] = n
                h, n, prev = process_file(in_path, w3, w7, w14, out_path, cols=cols, on_progress=prog)
                with LOCK:
                    task['status'] = 'done'
                    task['n_rows'] = n
                    task['header'] = h
                    task['preview_rows'] = prev
                    task['finished'] = time.time()
                try:
                    os.unlink(in_path)
                except OSError:
                    pass
            except Exception as e:
                import traceback
                with LOCK:
                    task['status'] = 'error'
                    task['error_msg'] = str(e)
                    task['traceback'] = traceback.format_exc()
                    task['finished'] = time.time()

        threading.Thread(target=run, daemon=True).start()
        self._json(202, {'id': tid, 'status': 'running', 'n_rows': 0})


def qs_get(p, key):
    v = urllib.parse.parse_qs(p.query).get(key)
    return v[0] if v else ''


def qs_get_list(p, key):
    return urllib.parse.parse_qs(p.query).get(key, [])


def run_fba(sid, task):
    try:
        up = UPLOADS.get(sid, {})
        sales_p = up.get('sales'); ops_p = up.get('ops'); fba_p = up.get('fba')
        if not (sales_p and ops_p and fba_p):
            raise ValueError("缺少上传文件（销量 / 运营级别 / FBA需求）")
        tid = task['id']
        out_sales = os.path.join(HERE, '_fba_%s_sales.xlsx' % tid)
        out_ops = os.path.join(HERE, '_fba_%s_ops.xlsx' % tid)
        out_diff = os.path.join(HERE, '_fba_%s_diff.xlsx' % tid)

        def prog(n):
            with LOCK:
                task['n_rows'] = n

        hdr_s, n_s, prev = process_file(sales_p, task['w'][0], task['w'][1], task['w'][2],
                                         out_sales, cols=task.get('cols'), on_progress=prog)
        hdr_o, n_o, meet, level_info = process_ops(ops_p, out_sales, out_ops, cols=task.get('cols'), on_progress=prog, levels=task.get('levels'))
        hdr_d, n_d, summary, diff_preview = process_diff(out_ops, fba_p, out_diff, cols=task.get('cols'))

        with LOCK:
            task['status'] = 'done'
            task['outputs'] = {'sales': out_sales, 'ops': out_ops, 'diff': out_diff}
            task['summary'] = summary
            task['n_rows'] = n_s
            task['header'] = hdr_s
            task['preview_rows'] = prev
            task['diff_header'] = hdr_d
            task['diff_preview'] = diff_preview
            task['level_info'] = level_info
            task['finished'] = time.time()
        # 清理暂存文件
        for k in ('sales', 'ops', 'fba'):
            pp = up.get(k)
            if pp:
                try:
                    os.unlink(pp)
                except OSError:
                    pass
        UPLOADS.pop(sid, None)
    except Exception as e:
        import traceback
        with LOCK:
            task['status'] = 'error'
            task['error_msg'] = str(e)
            task['traceback'] = traceback.format_exc()
            task['finished'] = time.time()


class ThreadingServer(http.server.ThreadingHTTPServer):
    allow_reuse_address = True


if __name__ == '__main__':
    print("服务已启动： http://localhost:%d  （Ctrl+C 停止）" % PORT)
    # 自检：打印当前代码实际包含的关键功能。重启后看这一行，
    # 如果缺少 +FBA 上传 或 +FBA 任务，说明跑的还是旧版 server.py。
    extras = []
    if hasattr(H, 'do_OPTIONS'):
        extras.append('OPTIONS预检')
    try:
        src = inspect.getsource(H.do_POST)
        if '/api/upload' in src:
            extras.append('FBA上传')
        if 'run_fba' in src:
            extras.append('FBA任务')
    except Exception:
        pass
    print("当前版本： %s" % ('基础处理' if not extras else '基础处理 + ' + ' + '.join(extras)))
    ThreadingServer(('0.0.0.0', PORT), H).serve_forever()
