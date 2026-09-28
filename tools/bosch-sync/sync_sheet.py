"""Parse the BOSCH Retail Membership Google Sheet (xlsx export) into one JSON doc per month.

Usage: python3 sync_sheet.py <drive-tool-result.json | file.xlsx> <out_dir>

Each tab named YYYY-MM or YYYYMM (e.g. 2026-09) is one month; other tabs (copies etc.) are ignored.
The dashed form is what the page's live read needs; the output id is always YYYYMM.
Per tab: row 3 = DDS, row 4 = Showroom (cols B/C/D = cumulative / monthly new / shipped);
store rows start at row 8 and run until column A is blank or a section header
("Total Retail", "1-on-1 ...") begins. Blank or non-numeric cells -> 0.
Writes <out_dir>/<YYYYMM>.json shaped {dds:{cum,monthly,shipped}, showroom:{...}, stores:[...]},
plus chat:{dds,showroom} (1-on-1 chat windows) and tags (total tags used) for months whose tab
has a "1-on-1 Chat" / "Total Tags Used" block.
"""
import base64, json, os, re, sys
import openpyxl


def num(v):
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return 0


def triple(row):
    return {'cum': num(row[1]), 'monthly': num(row[2]), 'shipped': num(row[3])}


def chat_block(rows):
    """The optional "1-on-1 Chat" block below the store list: DDS / Showroom chat windows in column B."""
    label = lambda r: '' if r[0] is None else str(r[0]).strip()
    start = next((i for i, r in enumerate(rows) if label(r).lower().startswith('1-on-1')), None)
    if start is None:
        return None
    found = {}
    for r in rows[start + 1:]:
        key = {'DDS': 'dds', 'Showroom': 'showroom'}.get(label(r))
        if key and key not in found:
            found[key] = num(r[1])
    return found if len(found) == 2 else None


def tags_block(rows):
    """The optional "Total Tags Used" block: the number in column B of its "Total Retail" row."""
    label = lambda r: '' if r[0] is None else str(r[0]).strip()
    start = next((i for i, r in enumerate(rows) if label(r).lower().startswith('total tags used')), None)
    if start is None:
        return None
    for r in rows[start + 1:]:
        if label(r) == 'Total Retail' and isinstance(r[1], (int, float)):
            return num(r[1])
    return None


def load_workbook(src):
    if src.endswith('.xlsx'):
        return openpyxl.load_workbook(src, data_only=True)
    data = base64.b64decode(json.load(open(src))['content'])
    tmp = src + '.xlsx'
    open(tmp, 'wb').write(data)
    return openpyxl.load_workbook(tmp, data_only=True)


def main(src, out_dir):
    wb = load_workbook(src)
    os.makedirs(out_dir, exist_ok=True)
    for name in wb.sheetnames:
        m = re.fullmatch(r'(\d{4})-?(\d{2})', name.strip())
        if not m:
            continue
        month_id = m.group(1) + m.group(2)
        rows = list(wb[name].iter_rows(min_row=1, max_row=wb[name].max_row, values_only=True))
        doc = {'dds': triple(rows[2]), 'showroom': triple(rows[3]), 'stores': []}
        for r in rows[7:]:
            label = '' if r[0] is None else str(r[0]).strip()
            if label == '' or label == 'Total Retail' or label.lower().startswith('1-on-1'):
                break  # end of the store list: blank row or the next section (e.g. the 1-on-1 Chat block)
            doc['stores'].append({'name': str(r[0]).strip(), **triple(r)})
        chat = chat_block(rows)
        if chat:
            doc['chat'] = chat
        tags = tags_block(rows)
        if tags is not None:
            doc['tags'] = tags
        json.dump(doc, open(os.path.join(out_dir, month_id + '.json'), 'w', encoding='utf-8'), ensure_ascii=False)
        print(month_id, 'DDS', doc['dds'], 'Showroom', doc['showroom'], len(doc['stores']), 'stores')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
