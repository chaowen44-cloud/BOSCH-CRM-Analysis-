"""Write parsed month JSONs (from sync_sheet.py) into the dashboard's <script id="liveData"> block.

Usage: python3 embed_data.py <months_dir> <dashboard.html>
Prints CHANGED (file rewritten, syncedAt bumped) or UNCHANGED (file untouched).
"""
import datetime, glob, json, os, re, sys

BLOCK = re.compile(r'(<script type="application/json" id="liveData">)(.*?)(</script>)', re.S)


def main(months_dir, html_path):
    months = {}
    for f in sorted(glob.glob(os.path.join(months_dir, '*.json'))):
        months[os.path.splitext(os.path.basename(f))[0]] = json.load(open(f, encoding='utf-8'))
    if not months:
        sys.exit('no month files found — refusing to blank the dashboard')
    html = open(html_path, encoding='utf-8').read()
    m = BLOCK.search(html)
    if not m:
        sys.exit('liveData block not found in ' + html_path)
    current = json.loads(m.group(2).replace('<\\/', '</'))
    if current.get('months') == months:
        print('UNCHANGED')
        return
    payload = {'syncedAt': datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'), 'months': months}
    body = json.dumps(payload, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    html = html[:m.start(2)] + body + html[m.end(2):]
    open(html_path, 'w', encoding='utf-8').write(html)
    print('CHANGED', sorted(months))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
