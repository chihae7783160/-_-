"""Build the query page from the host workbook; never edit project fields."""
import copy
import datetime
import json
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl

DEPARTMENTS = ['高屏處', '南高處', '雲嘉處', '北三處']


def norm(value):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(value or ''))).replace('臺', '台').replace('舘', '館')


def text(value):
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime('%Y/%m/%d')
    return '' if value is None else str(value).strip()


def area_name(value):
    value = text(value)
    return '旗美區' if value == '仁湖區' else value


def build(source, template, output):
    html = Path(template).read_text(encoding='utf-8')
    match = re.search(r'const DATA=(.*?);\s*const LABELS', html, re.S)
    if not match:
        raise ValueError('查詢頁缺少 DATA 資料，停止更新')
    data = json.loads(match[1])
    for store in data:
        store['area'] = area_name(store['area'])
        if 'hostArea' in store:
            store['hostArea'] = area_name(store['hostArea'])
        for record in store['hostRecords']:
            record['area'] = area_name(record['area'])
    original = copy.deepcopy(data)
    exact, canonical, aliases = {}, {}, {}

    def index(table, name, i):
        if norm(name):
            table.setdefault(norm(name), set()).add(i)

    for i, store in enumerate(data):
        index(canonical, store['name'], i)
        for record in store['hostRecords']:
            index(exact, record['name'], i)
        for name in [store['name'], *store['aliases']]:
            index(aliases, name, i)

    def resolve(name):
        candidates = exact.get(norm(name)) or canonical.get(norm(name)) or aliases.get(norm(name))
        if not candidates:
            parts = re.split(r'[\n()（）]', name)
            candidates = set().union(*(aliases.get(norm(part), set()) for part in parts if norm(part)))
        if candidates and len(candidates) == 1:
            return next(iter(candidates))
        if candidates:
            raise ValueError(f'店名對應不唯一：{name}，請先核對，停止更新')
        return None

    workbook = openpyxl.load_workbook(source, data_only=True)
    if not all(dept in workbook.sheetnames for dept in DEPARTMENTS):
        raise ValueError('缺少必要的處別工作表，停止更新')
    parsed = []
    for dept in DEPARTMENTS:
        sheet = workbook[dept]
        starts = [c.column for c in sheet[4] if c.value == '店名']
        if not starts:
            raise ValueError(f'{dept}：第4列找不到店名欄，停止更新')
        for position, col in enumerate(starts):
            end = starts[position + 1] if position + 1 < len(starts) else sheet.max_column + 1
            notes = [c for c in range(col + 1, end) if sheet.cell(4, c).value == '備註']
            if len(notes) != 1 or not sheet.cell(3, col).value:
                raise ValueError(f'{dept}：區別或備註欄格式不符，停止更新')
            note_col = notes[0]
            for row in range(5, sheet.max_row + 1):
                name = text(sheet.cell(row, col).value)
                if not name:
                    continue
                hosts = []
                for c in range(col + 1, note_col):
                    values = [text(sheet.cell(row + offset, c).value) for offset in range(3)]
                    if any(values):
                        hosts.append(dict(zip(['location', 'model', 'history'], values), cell=sheet.cell(row, c).coordinate))
                if not hosts:
                    raise ValueError(f'{dept}：{name} 沒有可讀取的主機資料，停止更新')
                parsed.append(dict(name=name, dept=dept, area=area_name(sheet.cell(3, col).value),
                                   hosts=hosts, note=text(sheet.cell(row, note_col).value), row=row))
    old_count = sum(len(s['hostRecords']) for s in data)
    if len(parsed) < old_count * 0.8:
        raise ValueError('主機門市數量減少超過20%，請核對是否上傳完整總表')
    touched, pending = set(), []
    for record in parsed:
        i = resolve(record['name'])
        if i is None:
            i = len(data)
            data.append(dict(id='', name=record['name'], area=record['area'], dept=record['dept'],
                             base={}, small={}, hasBase=False, hasSmall=False, aliases=[record['name']],
                             hosts=[], hostRecords=[], hostSummaries=[], pending=True, closed=False))
            index(exact, record['name'], i)
            index(aliases, record['name'], i)
            pending.append(record['name'])
        store = data[i]
        if i not in touched:
            store['hosts'], store['hostRecords'] = [], []
            touched.add(i)
        store['hostRecords'].append(record)
        store['hosts'].extend(record['hosts'])
        store['dept'] = record['dept']
        store['hostSource'] = record['dept']
        store['hostName'] = record['name']
        store['hostArea'] = record['area']
        store['hostNote'] = '\n'.join(r['note'] for r in store['hostRecords'] if r['note'])
        if record['name'] not in store['aliases']:
            store['aliases'].append(record['name'])
    # Model summary tables use merged area cells; carry the area down.
    summaries = [s for s in workbook if s.title.startswith('門市監控統計')]
    if len(summaries) != 1:
        raise ValueError('找不到唯一的門市監控統計工作表，停止更新')
    sheet = summaries[0]
    if sheet['B3'].value != '店名' or sheet['J3'].value != '店名':
        raise ValueError('主機型號統計表格式不符，停止更新')
    for store in data:
        store['hostSummaries'] = []
    for col, model in [(2, 'G216'), (10, 'G5016')]:
        area = ''
        for row in range(4, sheet.max_row + 1):
            area = area_name(sheet.cell(row, col - 1).value) or area
            name = text(sheet.cell(row, col).value)
            if not name:
                continue
            i = resolve(name)
            if i is None:
                i = len(data)
                data.append(dict(id='', name=name, area=area, dept='未提供處別', base={}, small={},
                                 hasBase=False, hasSmall=False, aliases=[], hosts=[], hostRecords=[],
                                 hostSummaries=[], pending=True, closed=False))
                index(aliases, name, i)
                pending.append(name)
            data[i]['hostSummaries'].append(dict(model=model, total=text(sheet.cell(row, col + 1).value),
                count=text(sheet.cell(row, col + 2).value), note=text(sheet.cell(row, col + 3).value)))
    for before, after in zip(original, data):
        for field in ['id', 'name', 'base', 'small', 'hasBase', 'hasSmall', 'closed', 'area']:
            assert before.get(field) == after.get(field), (before['name'], field)
    for i, store in enumerate(data):
        store['hostMissingFromLatest'] = bool(store['hostRecords'] and i not in touched)
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('<', '\\u003c')
    html = html[:match.start(1)] + payload + html[match.end(1):]
    stamp = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime('%Y/%m/%d %H:%M')
    banner = f'<p class="hint" style="padding:0 16px">主機總表同步時間：{stamp}（台灣時間）</p>'
    html = re.sub(r'<p class="hint" style="padding:0 16px">主機總表同步時間：.*?</p>', '', html)
    html = html.replace('</header>', '</header>' + banner, 1)
    html = html.replace('此名稱僅出現在 G216／G5016 統計表，尚未確認對應門市及代號。獨立保留來源資料，不與其他門市合併、不列入營業所家數。', '此名稱尚未確認對應門市及代號。獨立保留來源資料，不與其他門市合併、不列入營業所家數。')
    marker = "if(r.closed)content.append"
    html = html.replace(marker, "if(r.hostMissingFromLatest)content.append(el('p','warning','最新主機總表未列出此門市；以下保留先前主機紀錄，請核對。'));" + marker, 1)
    html = html.replace('（全名單：營業所 361 間）', '（全名單：營業中 ${DATA.filter(s=>!s.closed&&!s.pending).length} 間）')
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    Path(output).write_text(html, encoding='utf-8')
    report = dict(source_records=len(parsed), matched_stores=len(touched), pending_names=pending,
                  retained_missing=[s['name'] for i,s in enumerate(original) if s['hostRecords'] and i not in touched],
                  updated_at=stamp)
    Path(output).with_name('sync-report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    build(*(sys.argv[1:] or ['全聯-監控主機總表.xlsx', 'index.html', '_site/index.html']))
