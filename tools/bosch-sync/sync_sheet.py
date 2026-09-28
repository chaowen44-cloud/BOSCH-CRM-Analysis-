"""Parse the BOSCH Retail Membership Google Sheet (xlsx export) into one JSON doc per month.

Usage: python3 sync_sheet.py <drive-tool-result.json | file.xlsx> <out_dir>

Each tab named exactly YYYYMM (e.g. 202609) is one month; other tabs (copies etc.) are ignored.
Per tab: row 3 = DDS, row 4 = Showroom (cols B/C/D = cumulative / monthly new / shipped);
store rows start at row 8 and run until column A is blank or a section header
("Total Retail", "1-on-1 ...") begins. Blank or non-numeric cells -> 0.
Writes <out_dir>/<YYYYMM>.json shaped {dds:{cum,monthly,shipped}, showroom:{...}, stores:[...]}.
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
        if not re.fullmatch(r'\d{6}', name):
            continue
        rows = list(wb[name].iter_rows(min_row=1, max_row=wb[name].max_row, values_only=True))
        doc = {'dds': triple(rows[2]), 'showroom': triple(rows[3]), 'stores': []}
        for r in rows[7:]:
            label = '' if r[0] is None else str(r[0]).strip()
            if label == '' or label == 'Total Retail' or label.lower().startswith('1-on-1'):
                break  # end of the store list: blank row or the next section (e.g. the 1-on-1 Chat block)
            doc['stores'].append({'name': str(r[0]).strip(), **triple(r)})
        json.dump(doc, open(os.path.join(out_dir, name + '.json'), 'w', encoding='utf-8'), ensure_ascii=False)
        print(name, 'DDS', doc['dds'], 'Showroom', doc['showroom'], len(doc['stores']), 'stores')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
