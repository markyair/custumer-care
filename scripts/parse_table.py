"""Convert a Claude Docs `view` read of the tracking table into JSON.

Usage: python3 -I parse_table.py <tool-result-file> <out.json>
Output: {"rev": N, "rows": [{"id": "...", "cells": [13 strings]}, ...]}; row 0 is the header.
A dropdown chip cell looks like "[6] איום משפטי – להסלים".
"""
import json, re, html, sys
src, dst = sys.argv[1], sys.argv[2]
raw = open(src, encoding='utf-8').read()
try:
    d = json.loads(raw); xml = d['data']['xml']; rev = d.get('rev') or d['data'].get('rev')
except Exception:
    m = re.search(r'"xml":"(.*?)","complete"', raw, re.S)
    xml = json.loads('"' + m.group(1) + '"'); rev = int(re.search(r'"rev":(\d+)', raw).group(1))
rows = []
for rm in re.finditer(r"<row id='([^']*)'[^>]*>(.*?)</row>", xml, re.S):
    cells = []
    for cm in re.finditer(r"<cell id='[^']*'[^>]*>(.*?)</cell>", rm.group(2), re.S):
        c = cm.group(1)
        dd = re.search(r"<dropdown[^>]*index='(\d+)'[^>]*>(.*?)</dropdown>", c)
        if dd:
            cells.append('[' + dd.group(1) + '] ' + html.unescape(dd.group(2)))
        else:
            cells.append(''.join(html.unescape(t) for t in re.findall(r"<text[^>]*>(.*?)</text>", c, re.S)))
    rows.append({'id': rm.group(1), 'cells': cells})
json.dump({'rev': rev, 'rows': rows}, open(dst, 'w'), ensure_ascii=False)
print('REV', rev, 'ROWS', len(rows) - 1)
st = {}
for r in rows[1:]:
    k = r['cells'][7][:4]; st[k] = st.get(k, 0) + 1
print('status counts', st)
