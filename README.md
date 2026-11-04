# סוכן שירות ובדק, לאונרדו גרופ

גיבוי של הסקריפטים והפרומפטים של סוכן השירות והבדק, כדי שישרדו גם בלי SharePoint.‏ המקור המלא נמצא בתיקיית הפיילוט ב-SharePoint (פרויקטים / 2. פתוחים / הנהלה וכלליות / פיילוט שירות ובדק), בקבצים 02 ו-04.‏

## תוכן

- `prompts/morning-scan.md`: הפרומפט של "סריקת בוקר: פניות חדשות בשירות ובדק", ראשון עד חמישי 07:50 שעון ישראל (`CRON_TZ=Asia/Jerusalem 50 7 * * 0-4`).
- `prompts/weekly-schedule.md`: הפרומפט של "לו"ז שבועי: התחייבויות שירות ובדק", יום חמישי 16:34 (`CRON_TZ=Asia/Jerusalem 34 16 * * 4`).
- `scripts/parse_table.py`: ממיר קריאת view של טבלת המעקב במסמך Claude Docs ל-JSON.
- `scripts/build_xlsx.py`: בונה מה-JSON את קובץ ה-Excel המקוצר (כ-16 KB על 47 שורות) ואת ה-base64 להעלאה.

## הרצה

```bash
python3 -I scripts/parse_table.py <קובץ תוצאת ה-view> table.json
```

```bash
python3 -I scripts/build_xlsx.py table.json <תיקיית פלט>
```

דורש openpyxl.‏ פרטי ההעלאה והאימות נמצאים בסעיף 7 של כל פרומפט.
