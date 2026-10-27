#!/usr/bin/env python3
"""Conservative shipment-to-customer workbook updater (openpyxl)."""
import argparse
import datetime as dt
import json
import re
import sys
from collections import defaultdict
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties


def norm(v):
    return re.sub(r'[^A-Z0-9]', '', str(v or '').upper()).replace('CHESNUT', 'CHESTNUT')


def po(v):
    x = norm(v)
    if x.startswith('PO'):
        x = x[2:]
    return x[:-2] if x.endswith('TP') else x


def style(v):
    x = norm(v)
    return x[:-2] if x.endswith('TP') else x


def date_value(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%Y.%m.%d'):
            try:
                return dt.datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def sheet_po(s):
    ids = {po(s.cell(r, 1).value) for r in range(2, s.max_row + 1)
           if s.cell(r, 1).value and s.cell(r, 2).value and s.cell(r, 4).value}
    ids.discard('')
    if len(ids) == 1:
        return next(iter(ids))
    m = re.search(r'(?:CA|UK)?\d{6}(?:TP)?', s.title.upper())
    return po(m.group()) if m and not ids else None


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def is_dated_header(value):
    text = str(value or '')
    return bool(re.search(r'(?<!\d)(?:\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}\s*/\s*\d{1,2})(?!\d)', text))


def is_shipment_measure(value, totals=False):
    prefix = r'(?:UNITS|CBM)' if totals else r'(?:(?:TTL|TOTAL)\s+)?(?:UNITS|CBM)'
    return bool(re.match(r'^\s*' + prefix + r'\s+SHIPPED(?=\b|\d)', str(value or ''), re.I))


def is_shipment_percent(value):
    return bool(re.match(r'^\s*%\s*TO\b.*\bSHIPPED(?=\b|\d)', str(value or ''), re.I))


def interval_from_header(value, year, historical_year=None):
    text = str(value or '')
    if 'SHIPPED' not in text.upper():
        return None
    before = re.search(r'\bBEFORE\s+(\d{1,2})\s*/\s*(\d{1,2})(?:\s*/\s*(\d{4}))?', text, re.I)
    if before:
        if not before.group(3) and historical_year is None:
            raise ValueError(f'历史累计表头 {text!r} 的 BEFORE 日期缺少年份，无法判断是否交叠，请用 --historical-year 指定确认后的年份')
        cutoff = dt.date(int(before.group(3) or historical_year), int(before.group(1)), int(before.group(2)))
        return dt.date.min, cutoff - dt.timedelta(days=1)
    iso = re.search(r'(\d{4}-\d{2}-\d{2})\s*~\s*(\d{4}-\d{2}-\d{2})', text)
    if iso:
        return tuple(dt.date.fromisoformat(x) for x in iso.groups())
    m = re.search(r'(?<!\d)(\d{1,2})\s*/\s*(\d{1,2})\s*-\s*(\d{1,2})\s*/\s*(\d{1,2})(?!\d)', text)
    if m:
        a,b,c,d = map(int,m.groups())
        finish = dt.date(year,c,d)
        begin = dt.date(year-(a>c),a,b)
        return begin,finish
    return None


def main():
    p = argparse.ArgumentParser(description='更新客户出货表副本，生成待确认 JSON')
    p.add_argument('--customer', required=True)
    p.add_argument('--shipment', required=True)
    p.add_argument('--output', required=True)
    p.add_argument('--days', required=True, type=int)
    p.add_argument('--today', help='测试或重跑时指定今天 YYYY-MM-DD')
    p.add_argument('--historical-year', type=int, help='人工确认的无年份历史 BEFORE 表头年份')
    args = p.parse_args()
    if args.days < 1:
        p.error('--days 必须大于 0')
    end = (dt.date.fromisoformat(args.today) if args.today else dt.date.today()) - dt.timedelta(days=1)
    start = end - dt.timedelta(days=args.days - 1)
    customer, shipment, output = map(lambda x: Path(x).expanduser().resolve(), (args.customer, args.shipment, args.output))
    if output in (customer, shipment) or output.exists():
        p.error('输出路径不得覆盖输入或现有文件')
    if output.suffix.lower() != '.xlsx':
        p.error('输出必须是 .xlsx')
    report_path = output.with_suffix('.issues.json')
    if report_path.exists():
        p.error('待确认报告已存在，不覆盖')
    wb = load_workbook(customer)
    cached_customer = load_workbook(customer, data_only=True)
    for sheet in wb:
        for cell in sheet[1]:
            interval = interval_from_header(cell.value, end.year, args.historical_year)
            if interval and interval[0] <= end and start <= interval[1]:
                raise ValueError(f'{sheet.title}!{cell.coordinate} 已有区间 {interval[0]}~{interval[1]} 与请求区间交叠，停止处理；无年份表头按结束日年份检查')
    sw = load_workbook(shipment, data_only=True)
    raw = load_workbook(shipment, data_only=False)
    if len(sw.worksheets) < 2 or sw.worksheets[1].title != '单证部--工厂实际提货数据汇总':
        raise ValueError('出货表第二页名称不符，停止处理')
    source, source_raw = sw.worksheets[1], raw.worksheets[1]
    expected = {5:'款号', 6:'产品描述', 7:'外合同', 11:'数量', 14:'工厂提货日', 19:'总体积', 25:'货代反馈情况', 32:'货代'}
    for col, name in expected.items():
        if name not in str(source.cell(1,col).value or ''):
            raise ValueError(f'出货表 {get_column_letter(col)} 列表头不符: {name}')
    if 'TOTALS' not in wb:
        raise ValueError('找不到 TOTALS')
    totals = wb['TOTALS']
    if str(totals['A1'].value).strip().upper() != 'PO#':
        raise ValueError('TOTALS 表头不符')
    issues = []
    def issue(kind, **kwargs):
        issues.append({'类型':kind, **kwargs})
    details = defaultdict(list)
    for s in wb.worksheets:
        if s == totals:
            continue
        keys = {po(s.cell(r,1).value) for r in range(2,s.max_row+1)
                if s.cell(r,1).value and s.cell(r,2).value}
        keys.discard('')
        for key in keys:
            details[key].append(s)
    for key, sheets in details.items():
        if len(sheets) > 1:
            issue('重复PO明细页', PO=key, 页=[s.title for s in sheets], 处理='该PO本次不自动更新，需确认')
    total_rows = defaultdict(list)
    for r in range(2, totals.max_row+1):
        if totals.cell(r,1).value:
            total_rows[po(totals.cell(r,1).value)].append(r)
    for s in wb:
        for row in s:
            for c in row:
                if c.data_type == 'e' or (isinstance(c.value,str) and (c.value.startswith('==') or '#REF!' in c.value)):
                    issue('原表公式错误或疑似错误', 位置=f'{s.title}!{c.coordinate}', 内容=str(c.value)[:120])
    grouped = defaultdict(lambda: [0, 0.0, []])
    new_items = defaultdict(dict)
    new_shipments = defaultdict(lambda: [0, 0.0])
    for r in range(2, source.max_row+1):
        key, st = po(source.cell(r,7).value), style(source.cell(r,5).value)
        if not key or not st:
            if any(source.cell(r,c).value is not None for c in (5,6,7,11,14,19,25,32)):
                issue('外合同或款号缺失', 出货表行=r, 外合同=str(source.cell(r,7).value or ''), 款号=str(source.cell(r,5).value or ''))
            continue
        day = date_value(source.cell(r,14).value)
        if day is None:
            issue('提货日期为空或无法识别', 出货表行=r, PO=key, 日期=str(source.cell(r,14).value or ''))
            continue
        if not start <= day <= end:
            continue
        note = source.cell(r,25).value
        if note:
            issue('货代反馈需确认', 出货表行=r, PO=key, 备注=str(note), 处理='本行暂不计入')
            continue
        forwarder = norm(source.cell(r,32).value)
        if forwarder not in ('LAURA','FLEXPORT'):
            issue('货代无法识别', 出货表行=r, PO=key, 货代=str(source.cell(r,32).value or ''))
            continue
        q, cbm = source.cell(r,11).value, source.cell(r,19).value
        if not is_number(q) or not is_number(cbm) or q < 0 or cbm < 0:
            issue('数量或体积无有效缓存或为负数', 出货表行=r, PO=key, 数量=str(q), 体积=str(cbm), 公式=[str(source_raw.cell(r,c).value) for c in (11,19)])
            continue
        desc = str(source.cell(r,6).value or '')
        if key not in details:
            if any(po(m.group()) == key for sheet in wb.worksheets[1:] for m in [re.search(r'(?:CA|UK)?\d{6}(?:TP)?', sheet.title.upper())] if m):
                issue('已有PO工作表但明细无法识别', 出货表行=r, PO=key, 处理='明细及TOTALS均不计入，请补全明细并确认')
                continue
            new_items[key][(st, desc)] = True
            new_shipments[(key, st, desc, forwarder)][0] += q
            new_shipments[(key, st, desc, forwarder)][1] += cbm
            issue('新PO明细待确认', 出货表行=r, PO=key, 款号=st, 描述=desc, 处理='订单数量、体积、预计提货日期和颜色待填写；按原描述暂列出货数据')
            continue
        if len(details[key]) != 1:
            issue('重复PO无法选择明细页', 出货表行=r, PO=key)
            continue
        s = details[key][0]
        if any(s in ss and len(ss)>1 for ss in details.values()):
            issue('明细页含重复PO无法自动更新', 出货表行=r, PO=key, 页=s.title)
            continue
        nd = norm(desc)
        matches = [rr for rr in range(2,s.max_row+1)
                   if po(s.cell(rr,1).value) == key and style(s.cell(rr,2).value) == st and norm(s.cell(rr,4).value)
                   and norm(s.cell(rr,4).value) in nd]
        if len(matches) != 1:
            issue('款号颜色无法唯一匹配', 出货表行=r, PO=key, 款号=st, 描述=desc, 候选行=matches)
            continue
        rr = matches[0]
        v = grouped[(key,rr,forwarder)]
        v[0] += q; v[1] += cbm; v[2].append(r)
    label = f'{start.isoformat()}~{end.isoformat()}'
    for s in wb:
        if any(label in str(c.value or '') for c in s[1]):
            raise ValueError(f'{s.title} 已有该区间，停止以避免重复统计')
    # Append new columns instead of inserting: openpyxl does not update cross-sheet formula references.
    processed_sheets = set()
    for key, sheets in list(details.items()):
        if len(sheets) != 1:
            continue
        s = sheets[0]
        if len([k for k, ss in details.items() if s in ss and len(ss)>1]) > 0:
            issue('含重复PO的明细页未更新', 页=s.title)
            continue
        if s.title in processed_sheets:
            continue
        processed_sheets.add(s.title)
        suspicious = [c.coordinate for c in s[1]
                      if (is_dated_header(c.value) or re.search(r'\b(?:SHIPPED|SHIPMENT|SHIPPING|SHIPPED\w*)\b|出货|出运', str(c.value or ''), re.I))
                      and not is_shipment_measure(c.value)
                      and not is_shipment_percent(c.value)
                      and (is_dated_header(c.value) or not re.search(r'\bPENDING\b|\bREMAINING\b', str(c.value or ''), re.I))]
        if suspicious:
            issue('历史出货表头无法识别', 页=s.title, 单元格=suspicious, 处理='保留待出货数量及体积原值，需确认')
        cols = {}
        for f in ('LAURA','FLEXPORT'):
            cols[f] = s.max_column + 1
            for offset, text in enumerate(('UNITS','% TO ORDER','CBM')):
                c=s.cell(1,cols[f]+offset, f'{text} SHIPPED {label} ({f})')
                c._style=copy(s.cell(1, max(1,cols[f]-1))._style)
            for rr in range(2,s.max_row+1):
                if not (s.cell(rr,2).value and s.cell(rr,4).value):
                    continue
                vals = grouped.get((po(s.cell(rr,1).value),rr,f))
                s.cell(rr,cols[f], vals[0] if vals else 0)
                s.cell(rr,cols[f]+1, f'=IF(E{rr}="","",IFERROR({get_column_letter(cols[f])}{rr}/E{rr},0))')
                s.cell(rr,cols[f]+2, vals[1] if vals else 0)
                s.cell(rr,cols[f]+1).number_format='0.00%'
        # Pending formula is replaced only where exactly one matching header exists.
        pending = [c.column for c in s[1] if re.fullmatch(r'UNITS PENDING SHIPMENT\s*', str(c.value or ''), re.I)]
        shipcols = [c.column for c in s[1] if is_shipment_measure(c.value) and re.match(r'^\s*(?:(?:TTL|TOTAL)\s+)?UNITS\b', str(c.value or ''), re.I)]
        cbmcols = [c.column for c in s[1] if is_shipment_measure(c.value) and re.match(r'^\s*(?:(?:TTL|TOTAL)\s+)?CBM\b', str(c.value or ''), re.I)]
        if not cbmcols:
            issue('缺少可识别历史出货体积列', 页=s.title, 处理='待出货体积无法确认')
        pending_cbm = [c.column for c in s[1] if re.fullmatch(r'CBM OF UNITS PENDING SHIPMENT\s*', str(c.value or ''), re.I)]
        if len(pending_cbm) != 1:
            issue('待出货体积列不唯一或缺失', 页=s.title, 列=pending_cbm, 处理='保留原值，需确认')
        for rr in range(2,s.max_row+1):
            if s.cell(rr,2).value and s.cell(rr,4).value and len(pending_cbm) == 1 and not suspicious:
                refs = ','.join(f'{get_column_letter(col)}{rr}' for col in cbmcols)
                s.cell(rr,pending_cbm[0],f'=IF(G{rr}="","",G{rr}-SUM({refs}))')
        if len(pending) != 1:
            issue('待出货列不唯一', 页=s.title, 列=pending, 处理='保留原公式，需人工确认')
        elif not suspicious:
            for rr in range(2,s.max_row+1):
                if not (s.cell(rr,2).value and s.cell(rr,4).value): continue
                refs=','.join(f'{get_column_letter(col)}{rr}' for col in shipcols)
                s.cell(rr,pending[0],f'=IF(E{rr}="","",E{rr}-SUM({refs}))')
                historical = []
                unknown = []
                for col in shipcols:
                    cell = s.cell(rr,col)
                    val = cell.value
                    if cell.data_type == 'f':
                        val = cached_customer[s.title].cell(rr,col).value if col <= cached_customer[s.title].max_column else None
                        if not is_number(val): unknown.append(cell.coordinate)
                    elif val is not None and not is_number(val): unknown.append(cell.coordinate)
                    historical.append(val if is_number(val) else 0)
                order = s.cell(rr,5).value
                if s.cell(rr,5).data_type == 'f':
                    order = cached_customer[s.title].cell(rr,5).value
                if not is_number(order): unknown.append(f'E{rr}')
                if unknown:
                    issue('超发无法自动核验', 页=s.title, 行=rr, 单元格=unknown, 处理='公式缓存或数值不可用，需人工确认')
                elif sum(historical) > order:
                    issue('疑似超发', 页=s.title, 行=rr, 处理='请核对退回及特殊备注')
    for key, items in new_items.items():
        title = ('PO#_'+key+'TP')[:31]
        if title in wb:
            issue('新PO工作表重名', PO=key, 处理='明细及TOTALS均不计入')
            for shipment_key in list(new_shipments):
                if shipment_key[0] == key:
                    del new_shipments[shipment_key]
            continue
        s=wb.create_sheet(title)
        s.append(['PO#','Style #','Description','Color','Ttl Units Orders','Date Ready for Pickup','Ttl CBM'])
        for st, desc in items:
            s.append([key+'TP',st,desc,None,None,None,None])
        for f in ('LAURA','FLEXPORT'):
            for text in ('UNITS','% TO ORDER','CBM'):
                s.cell(1,s.max_column+1,f'{text} SHIPPED {label} ({f})')
        s.cell(1,14,'Units Pending Shipment')
        s.cell(1,15,'% To Ttl Units Pending Shipment')
        s.cell(1,16,'CBM of Units Pending Shipment')
        for rr, (st, desc) in enumerate(items, 2):
            for offset, f in ((8, 'LAURA'), (11, 'FLEXPORT')):
                vals = new_shipments[(key, st, desc, f)]
                s.cell(rr, offset, vals[0]); s.cell(rr, offset+2, vals[1])
                s.cell(rr, offset+1, f'=IF(E{rr}="","",IFERROR({get_column_letter(offset)}{rr}/E{rr},0))').number_format='0.00%'
            s.cell(rr,14,f'=IF(E{rr}="","",E{rr}-SUM(H{rr},K{rr}))')
            s.cell(rr,15,f'=IF(E{rr}="","",IFERROR(N{rr}/E{rr},0))').number_format='0.00%'
            s.cell(rr,16,f'=IF(G{rr}="","",G{rr}-SUM(J{rr},M{rr}))')
        details[key]=[s]
        issue('新PO基础资料缺失', PO=key, 页=title, 处理='请填写颜色、订单数量、订单体积和预计提货日期并确认后重跑')
    # TOTALS new columns use shipment aggregates, independent of sheet layout.
    tc = totals.max_column+1
    for i, (f, measure) in enumerate((('LAURA','UNITS'),('LAURA','CBM'),('FLEXPORT','UNITS'),('FLEXPORT','CBM'))):
        totals.cell(1,tc+i,f'{measure} SHIPPED {label} ({f})')
        totals.cell(1,tc+i)._style=copy(totals.cell(1,tc-1)._style)
    added_total_rows = {}
    for key in details:
        if key not in total_rows and key not in new_items:
            r=totals.max_row+1; totals.cell(r,1,key+'TP'); total_rows[key].append(r); added_total_rows[key]=r
            issue('TOTALS新增PO基础资料缺失', PO=key, 行=r, 处理='请填写或核对订单数量、体积和预计日期')
    for key in new_items:
        if key not in total_rows:
            r=totals.max_row+1; totals.cell(r,1,key+'TP'); total_rows[key].append(r); added_total_rows[key]=r
    for key, r in added_total_rows.items():
        sheets = details.get(key, [])
        if len(sheets) != 1:
            continue
        s = sheets[0]
        rows = [rr for rr in range(2,s.max_row+1) if po(s.cell(rr,1).value)==key and s.cell(rr,2).value]
        if not rows:
            continue
        quoted = "'" + s.title.replace("'", "''") + "'!"
        for dest, col in ((8,'E'),(9,'G')):
            refs = ','.join(f'{quoted}{col}{rr}' for rr in rows)
            totals.cell(r,dest,f'=IF(COUNT({refs})<{len(rows)},"",SUM({refs}))')
        dates = [s.cell(rr,6).value for rr in rows if s.cell(rr,6).value is not None]
        if dates and len(dates)==len(rows) and all(x==dates[0] for x in dates) and not (isinstance(dates[0],str) and dates[0].startswith('=')):
            datecols = [c.column for c in totals[1] if str(c.value or '').strip().upper() == 'EXPECTED READY DATE OF PENDING GOODS']
            if len(datecols) == 1:
                totals.cell(r,datecols[0],dates[0])
            else:
                issue('TOTALS预计日期列不唯一或缺失', PO=key)
        else:
            issue('TOTALS预计日期无法安全确定', PO=key, 页=s.title)
    for key, rows in total_rows.items():
        if len(rows)!=1:
            issue('TOTALS重复PO', PO=key, 行=rows); continue
        r=rows[0]
        for i,(f,j) in enumerate((('LAURA',0),('LAURA',1),('FLEXPORT',0),('FLEXPORT',1))):
            value = sum(v[j] for (k,rr,ff),v in grouped.items() if k==key and ff==f)
            value += sum(v[j] for (k,st,desc,ff),v in new_shipments.items() if k==key and ff==f)
            totals.cell(r,tc+i,value)
    pending_cols = [c.column for c in totals[1] if str(c.value or '').strip().upper() == 'REMAINING UNITS TO SHIP']
    total_shipcols = [c.column for c in totals[1] if is_shipment_measure(c.value, totals=True) and str(c.value or '').strip().upper().startswith('UNITS')]
    total_suspicious = [c.coordinate for c in totals[1]
                        if (is_dated_header(c.value) or re.search(r'\b(?:SHIPPED|SHIPMENT|SHIPPING|SHIPPED\w*)\b|出货|出运', str(c.value or ''), re.I))
                        and not is_shipment_measure(c.value, totals=True)
                        and not is_shipment_percent(c.value)
                        and (is_dated_header(c.value) or not re.search(r'\bPENDING\b|\bREMAINING\b', str(c.value or ''), re.I))]
    if total_suspicious:
        issue('历史出货表头无法识别', 页=totals.title, 单元格=total_suspicious, 处理='保留待出货数量及体积原值，需确认')
    if len(pending_cols) == 1 and not total_suspicious:
        for key, rows in total_rows.items():
            if len(rows) == 1:
                r = rows[0]
                refs = ','.join(f'{get_column_letter(col)}{r}' for col in total_shipcols)
                totals.cell(r,pending_cols[0],f'=IF(H{r}="","",H{r}-SUM({refs}))')
    else:
        issue('TOTALS待出货列不唯一', 列=pending_cols)
    total_cbm_pending = [c.column for c in totals[1] if str(c.value or '').strip().upper() == 'REMAINING CBM TO SHIP']
    total_cbm_cols = [c.column for c in totals[1] if is_shipment_measure(c.value, totals=True) and str(c.value or '').strip().upper().startswith('CBM')]
    if len(total_cbm_pending)==1 and not total_suspicious:
        for key, rows in total_rows.items():
            if len(rows)==1:
                r=rows[0]; refs=','.join(f'{get_column_letter(col)}{r}' for col in total_cbm_cols)
                totals.cell(r,total_cbm_pending[0],f'=IF(I{r}="","",I{r}-SUM({refs}))')
    else:
        issue('TOTALS待出货体积列不唯一或缺失', 列=total_cbm_pending)
    wb.calculation=CalcProperties(calcMode='auto', fullCalcOnLoad=True, forceFullCalc=True)
    output.parent.mkdir(parents=True,exist_ok=True)
    wb.save(output)
    report={'区间':{'开始':str(start),'结束':str(end),'天数':args.days},'已匹配出货行':sum(len(x[2]) for x in grouped.values()),'待确认数量':len(issues),'待确认':issues,'说明':'新增列在最右侧；原公式错误未自动修复；Excel打开时需重算。'}
    report_path.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'output':str(output),'issues':str(report_path),'matched_source_rows':report['已匹配出货行'],'issue_count':len(issues)},ensure_ascii=False))


if __name__ == '__main__':
    try:
        main()
    except (ValueError, OSError) as exc:
        print(f'错误: {exc}',file=sys.stderr)
        sys.exit(2)
