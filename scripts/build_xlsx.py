"""Build the shortened Excel file from the parsed table JSON.

Usage: python3 -I build_xlsx.py <table.json> <outdir>
Shortening (Excel only): summary 200 chars, action 160, last update = latest entry only, 220 chars.
On 47 rows this gave 16,144 bytes and uploaded successfully on 4.11.2026.
Writes <outdir>/טבלת פניות שירות ובדק.xlsx and <outdir>/b64.txt.
"""
import json, re, zipfile, io, base64, hashlib, sys, os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

STATUS = ["תקלה חדשה","ממתין למענה שלנו – חריגה","ביקור/פגישה לתיאום","קבלן לא זמין / סירב לביקור",
          "נדרשת הכרעת מהנדס/הנהלה","דרישה כספית/פיצוי","איום משפטי – להסלים","לא בדק / צד ג'",
          "טופל – נסגר","נענה – ממתין לביצוע","בביצוע"]
SENDER = ["דייר","מעצב/נציג דייר","יזם/לקוח","ועד/חברת ניהול","קבלן/ספק","מוקד טלפוני","צד ג'"]
FILL = {"איום משפטי – להסלים":"F4B6B6","ממתין למענה שלנו – חריגה":"F8CBAD","דרישה כספית/פיצוי":"FFE699",
        "נדרשת הכרעת מהנדס/הנהלה":"FFF2CC","קבלן לא זמין / סירב לביקור":"FFE699","ביקור/פגישה לתיאום":"DDEBF7",
        "נענה – ממתין לביצוע":"DDEBF7","בביצוע":"C6E0B4","טופל – נסגר":"E2EFDA","תקלה חדשה":"EDEDED",
        "לא בדק / צד ג'":"D9D9D9"}

src, outdir = sys.argv[1], sys.argv[2]
t = json.load(open(src)); rows = [r['cells'][:] for r in t['rows']]

def chip(v, lst):
    m = re.match(r'\[(\d+)\]', v or ''); return lst[int(m.group(1))] if m else (v or '')
def cap(s, n):
    s = s or ''; return s if len(s) <= n else s[:n-1].rstrip() + '…'
def short1(note, n):
    parts = [p.strip() for p in (note or '').split('; ') if p.strip()]
    return (('(…) ' if len(parts) > 1 else '') + cap(parts[-1], n)) if parts else ''

hdr = rows[0][:]
hdr[6] = 'תמצית (מקוצר)'; hdr[8] = 'פעולה מוצעת (מקוצר)'; hdr[12] = 'עדכון אחרון (מקוצר; המלא במסמך החי)'
wb = Workbook(); ws = wb.active; ws.title = 'פניות'; ws.sheet_view.rightToLeft = True
ws.append(hdr)
for c in ws[1]:
    c.fill = PatternFill('solid', fgColor='1F3864'); c.font = Font(bold=True, color='FFFFFF')
    c.alignment = Alignment(wrap_text=True, vertical='top')
cnt = {}
for r in rows[1:]:
    out = r[:]
    out[5] = chip(r[5], SENDER); out[7] = chip(r[7], STATUS)
    out[6] = cap(r[6], 200); out[8] = cap(r[8], 160); out[12] = short1(r[12], 220)
    try: out[0] = int(out[0])
    except Exception: pass
    ws.append(out); cnt[out[7]] = cnt.get(out[7], 0) + 1
    f = FILL.get(out[7])
    if f: ws.cell(row=ws.max_row, column=8).fill = PatternFill('solid', fgColor=f)
al = Alignment(wrap_text=True, vertical='top')
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment = al
for i, w in enumerate([5,11,26,14,24,14,55,24,45,14,11,9,45], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions

buf = io.BytesIO(); wb.save(buf); data = buf.getvalue()
zin = zipfile.ZipFile(io.BytesIO(data)); out = io.BytesIO()
zout = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED, compresslevel=9)
for item in zin.infolist():
    n = item.filename
    if n.startswith('xl/theme/') or n.startswith('docProps/'): continue
    b = zin.read(n)
    if n == '[Content_Types].xml':
        s = b.decode(); s = re.sub(r'<Override[^>]*theme[^>]*/>', '', s); s = re.sub(r'<Override[^>]*docProps[^>]*/>', '', s); b = s.encode()
    if n == 'xl/_rels/workbook.xml.rels':
        s = b.decode(); s = re.sub(r'<Relationship[^>]*theme[^>]*/>', '', s); b = s.encode()
    if n == '_rels/.rels':
        s = b.decode(); s = re.sub(r'<Relationship[^>]*docProps[^>]*/>', '', s); b = s.encode()
    zout.writestr(item, b, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
zout.close(); final = out.getvalue()
load_workbook(io.BytesIO(final))  # sanity check: the file opens
open(os.path.join(outdir, 'טבלת פניות שירות ובדק.xlsx'), 'wb').write(final)
b64 = base64.b64encode(final).decode()
open(os.path.join(outdir, 'b64.txt'), 'w').write(b64)
print('BYTES', len(final), 'B64LEN', len(b64), 'SHA', hashlib.sha256(final).hexdigest()[:12])
print(cnt)
