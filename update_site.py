"""
Admin Updater for mythicheatsheet.netlify.app
Usage:
    python update_site.py [timetable.xlsx] [datesheet.xlsx]
"""
import json, re, sys
from datetime import datetime
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parent
DEPT_ORDER = ["CS", "SE", "DS", "AI", "CY", "CI"]
SHORT_DAYS = {"mon": "Monday", "tue": "Tuesday", "wed": "Wednesday", "thu": "Thursday", "fri": "Friday", "sat": "Saturday", "sun": "Sunday"}

def clean(val):
    return re.sub(r"\s+", " ", str(val or "")).strip()

def pad_time(val):
    m = re.match(r"^(\d{1,2}):(\d{2})$", val)
    return f"{int(m.group(1)):02d}:{int(m.group(2)):02d}" if m else None

def add_minutes(start, minutes):
    h, m = map(int, start.split(":"))
    total = (h * 60 + m + int(round(minutes))) % (24 * 60)
    return f"{total // 60:02d}:{total % 60:02d}"

def infer_credits(raw_cr, code, title, duration):
    cr_str = clean(raw_cr).upper()
    if cr_str in {"1", "2", "3", "4", "NC", "1+1"}:
        return 0 if cr_str == "NC" else int(cr_str[0])
    is_lab = "lab" in title.lower() or bool(re.match(r"^(CL|EL|SL|AL|DL)\d", code.upper()))
    if is_lab:
        return 1
    if 75 <= duration <= 95:
        return 3
    if duration >= 150:
        return 3
    if 100 <= duration <= 130:
        return 1 if "quran" in title.lower() else 2
    if duration <= 65:
        return 0 if "seerat" in title.lower() or "sirat" in title.lower() else 1
    return 3

def parse_timetable(tt_path):
    wb = openpyxl.load_workbook(tt_path, data_only=True)
    entries = []
    for dept in DEPT_ORDER:
        if dept not in wb.sheetnames:
            continue
        ws = wb[dept]
        for r in range(3, ws.max_row + 1):
            code = clean(ws.cell(r, 1).value)
            title = clean(ws.cell(r, 2).value)
            section = clean(ws.cell(r, 3).value)
            teacher = clean(ws.cell(r, 4).value)
            raw_cr = ws.cell(r, 5).value
            try:
                duration = float(ws.cell(r, 12).value)
            except (TypeError, ValueError):
                duration = 80.0
            if not code or not title:
                continue
            credits = infer_credits(raw_cr, code, title, duration)
            for idx, (dc, sc, rc) in enumerate([(13, 14, 15), (16, 17, 18)], 1):
                day = SHORT_DAYS.get(clean(ws.cell(r, dc).value).lower()[:3])
                start = pad_time(clean(ws.cell(r, sc).value).replace(".", ":"))
                room = clean(ws.cell(r, rc).value)
                if day and start:
                    entries.append({
                        "id": f"fsc-{dept.lower()}-{r}-{idx}",
                        "day": day,
                        "slot": f"{start} - {add_minutes(start, duration)}",
                        "course": title,
                        "code": code,
                        "dept": dept,
                        "section": section,
                        "room": room,
                        "teacher": teacher,
                        "credits": credits,
                        "duration": int(duration)
                    })
    return entries

def parse_datesheet(ds_path):
    wb = openpyxl.load_workbook(ds_path, data_only=True)
    ws = wb.active
    banner = clean(ws.cell(1, 1).value) or "MIDTERM-II"
    exam_type = "MIDTERM-I" if "MIDTERM-I " in banner.upper() or banner.upper().endswith("MIDTERM-I") else (
                "FINAL EXAMS" if "FINAL" in banner.upper() else "MIDTERM-II")
    header_row = 3
    for r in range(1, 8):
        if any(":" in clean(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)):
            header_row = r
            break
    slots = []
    for c in range(3, ws.max_column + 1, 2):
        s_val = clean(ws.cell(header_row, c + 1).value or ws.cell(header_row, c).value)
        if ":" in s_val:
            slots.append((c, c + 1, s_val))
    exams, cur_day, cur_date = [], "", ""
    for r in range(header_row + 1, ws.max_row + 1):
        d, dt = clean(ws.cell(r, 1).value), clean(ws.cell(r, 2).value)
        if d: cur_day = d
        if dt: cur_date = dt
        for c_code, c_name, slot_str in slots:
            code, name = clean(ws.cell(r, c_code).value), clean(ws.cell(r, c_name).value)
            if code and name:
                exams.append({"date": cur_date, "day": cur_day, "slot": slot_str, "code": code, "course": name})
    return exam_type, exams
def validate_workbook(wb, required_sheets, min_rows=10):
    """Ensures the admin uploaded a real FAST file, not a random or corrupted document."""
    for sheet in required_sheets:
        if sheet not in wb.sheetnames:
            raise ValueError(f"CRITICAL ERROR: Missing expected sheet '{sheet}'. Did you upload the wrong file?")
        ws = wb[sheet]
        if ws.max_row < min_rows:
            raise ValueError(f"CRITICAL ERROR: Sheet '{sheet}' only has {ws.max_row} rows. The file appears incomplete.")
    print("✓ Workbook schema validated successfully.")
if __name__ == "__main__":
    tt_file = Path(sys.argv[1]) if len(sys.argv) > 1 else next(ROOT.glob("*TT*.xlsx"))
    ds_file = Path(sys.argv[2]) if len(sys.argv) > 2 else next(ROOT.glob("*date*sheet*.xlsx"))
    entries = parse_timetable(tt_file)
    exam_type, exams = parse_datesheet(ds_file)
    now_str = datetime.now().strftime("%d %B %Y %I:%M %p").upper()
    v_match = re.search(r"v(\d+(?:\.\d+)*)", tt_file.name, re.I)
    meta = {
        "site": "mythicheatsheet.netlify.app",
        "lastUpdated": now_str,
        "version": f"Fall 2026 FSC v{v_match.group(1) if v_match else '1.1.5'}",
        "examType": exam_type,
        "count": len(entries),
        "examCount": len(exams)
    }
    payload = f"window.TIMETABLE_DATA = {json.dumps({'meta': meta, 'entries': entries, 'exams': exams}, indent=2)};\n"
    (ROOT / "data.js").write_text(payload, encoding="utf-8")
    print(f"Updated data.js -> {len(entries)} classes, {len(exams)} exams ({exam_type}) @ {now_str}")
    