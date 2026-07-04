import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
import re
import os

def extract_roll_name(text):
    if not text:
        return "", ""

    text = str(text).strip()

    cleaned = re.sub(r'[^A-Za-z0-9 ]+', ' ', text)

    match = re.search(r'\d+[A-Za-z]+\d+', cleaned)

    if match:
        roll = match.group().upper()
        name = cleaned.replace(match.group(), "").strip()
        return roll, name

    return "", cleaned.strip()

def is_red_or_yellow(cell):
    fill = cell.fill
    if not fill or not fill.start_color:
        return False

    rgb = fill.start_color.rgb
    if not rgb:
        return False

    rgb = rgb.upper()

    RED = "00FF7F7F"
    YELLOW = "00FFD966"
    GREEN = "0090EE90"

    if rgb == GREEN:
        return False

    return rgb in (RED, YELLOW)


def extract_att_data(file_path):
    wb = load_workbook(file_path)
    ws = wb["ATT"]

    data = []

    for row in ws.iter_rows(min_row=2):
        name_cell = row[0]
        duration_cell = row[1]

        if not name_cell.value:
            continue

        if is_red_or_yellow(duration_cell):
            roll, name = extract_roll_name(name_cell.value)

            data.append({
                "Roll Number": roll,
                "Name": name,
                "Issue Type": "LESS_TIME"
            })

    return data


def extract_late_data(file_path):
    df = pd.read_excel(file_path, sheet_name="LATE")

    data = []

    for value in df.iloc[:, 0]:
        if pd.isna(value):
            continue

        roll, name = extract_roll_name(value)

        data.append({
            "Roll Number": roll,
            "Name": name,
            "Issue Type": "LATE_ARRIVAL"
        })

    return data


def export_to_csv(data, save_path, file_name):
    if not data:
        return None

    df = pd.DataFrame(data)

    df = df.rename(columns={
        "Roll Number": "roll_number",
        "Name": "student_name",
        "Issue Type": "issue_type"
    })

    df = df[["roll_number", "student_name", "issue_type"]]
    df.fillna("", inplace=True)

    full_path = os.path.join(save_path, f"{file_name}.csv")
    df.to_csv(full_path, index=False)

    return full_path

# +----------------+
# | EXCEL STYLES   |
# +----------------+
HEADER_FILL = PatternFill("solid", fgColor="d1d1d1")
HEADER_FONT = Font(color="000000", bold=True)

ROW1 = PatternFill("solid", fgColor="FFFFFF")
ROW2 = PatternFill("solid", fgColor="d1d1d1")

GREEN = PatternFill("solid", fgColor="90EE90")
RED = PatternFill("solid", fgColor="FF7F7F")
YELLOW = PatternFill("solid", fgColor="FFD966")


# +----------------+
# | REGEX / CONST  |
# +----------------+
ROLL_PATTERN = re.compile(r'([A-Z]{1,4}\s*[_\-]?\s*\d{2,4})', re.IGNORECASE)
NON_WORD_RE = re.compile(r"[^\w\s]", re.UNICODE)
PAREN_RE = re.compile(r"\(.*?\)")
APOSTROPHE_RE = re.compile(r"[’‘`]")

BLACKLIST_WORDS = {"ZOOM USER"}
TITLE_WORDS = {"HOSTS"}

# +----------------+
# | HELPERS        |
# +----------------+
def find_column(df, keys):
    for c in df.columns:
        if any(k in c.lower() for k in keys):
            return c
    raise ValueError(f"Missing required column: {keys}")

def parse_datetime(col):
    return pd.to_datetime(col, errors="coerce", format="mixed")

def format_datetime_12h(dt):
    if pd.isna(dt):
        return ""
    return dt.strftime("%m/%d/%Y %I:%M:%S %p")

def extract_roll(name):
    if not name:
        return None
    m = ROLL_PATTERN.search(str(name))
    if not m:
        return None
    return m.group(1).upper().replace(" ", "").replace("_", "").replace("-", "")

def normalize_name(name):
    if not name:
        return ""
    s = PAREN_RE.sub("", str(name))
    s = ROLL_PATTERN.sub("", s)
    s = NON_WORD_RE.sub("", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()

def extract_alias(name):
    if not name:
        return None
    m = re.search(r"\(([^)]+)\)", str(name))
    return normalize_name(m.group(1)) if m else None

def is_blacklisted(name):
    if not name:
        return False
    s = APOSTROPHE_RE.sub("'", str(name).lower())
    if any(w in s for w in BLACKLIST_WORDS):
        return True
    parts = s.split()
    return len(parts) >= 2 and parts[-1] in TITLE_WORDS

def format_table(ws):
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        col[0].fill = HEADER_FILL
        col[0].font = HEADER_FONT
        width = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = width + 3
    for r in range(2, ws.max_row + 1):
        fill = ROW1 if r % 2 == 0 else ROW2
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).fill = fill

# +----------------+
# | MAIN PROCESS   |
# +----------------+
def process(
    csv_path,
    do_att,
    do_late,
    late_time,
    late_ampm,
    total_att,
    absent_limit,
    out_path
):
    # ---------- LOAD ----------
    df = pd.read_csv(csv_path)
    df.columns = df.columns.str.strip()

    name_col  = find_column(df, {"name", "participant", "user"})
    join_col  = find_column(df, {"join"})
    leave_col = find_column(df, {"leave"})
    dur_col   = find_column(df, {"duration"})

    df["_JOIN_DT"]  = parse_datetime(df[join_col])
    df["_LEAVE_DT"] = parse_datetime(df[leave_col])
    df["_DUR"] = pd.to_numeric(df[dur_col], errors="coerce").fillna(0)

    df["_ROLL"]      = df[name_col].apply(extract_roll)
    df["_NAME_KEY"]  = df[name_col].apply(normalize_name)
    df["_ALIAS_KEY"] = df[name_col].apply(extract_alias)

    df_original = df.copy()

    # ---------- CF / BLACKLIST ----------
    cf_mask = df_original[name_col].apply(is_blacklisted)
    cf = df_original.loc[cf_mask].copy()
    df = df_original.loc[~cf_mask].copy()

    # ---------- MEETING DURATION ----------
    vt = df.dropna(subset=["_JOIN_DT", "_LEAVE_DT"])
    meeting_minutes = (
        vt["_LEAVE_DT"].max() - vt["_JOIN_DT"].min()
    ).total_seconds() / 60.0 if not vt.empty else 0

    # ---------- PERSON MODEL ----------
    class Person:
        def __init__(self, r):
            self.roll = r["_ROLL"]
            self.name = r["_NAME_KEY"]
            self.alias = r["_ALIAS_KEY"]
            self.display = r[name_col]
            self.duration = r["_DUR"]
            self.first_join = r["_JOIN_DT"]

        def priority_name(self):
            # brackets > roll > plain
            if "(" in self.display and ")" in self.display:
                return (0, self.display)
            if self.roll:
                return (1, self.display)
            return (2, self.display)

    persons = []

    # ---------- LEVEL 1: EXACT DUPLICATES ----------
    for _, r in df.iterrows():
        matched = False
        for p in persons:
            if p.roll == r["_ROLL"] and p.name == r["_NAME_KEY"] and p.first_join == r["_JOIN_DT"]:
                p.duration += r["_DUR"]
                matched = True
                break
        if not matched:
            persons.append(Person(r))

    # ---------- LEVEL 1.5: NO-ROLL NAME DUPLICATES ----------
    merged = []
    for p in persons:
        if not p.roll:
            found = False
            for m in merged:
                if not m.roll and m.name == p.name:
                    m.duration += p.duration
                    found = True
                    break
            if not found:
                merged.append(p)
        else:
            merged.append(p)

    persons = merged


    # ---------- LEVEL 2: BRACKET ALIAS ----------
    merged = []
    for p in persons:
        done = False
        for m in merged:
            if p.alias and p.alias == m.name and m.roll:
                m.duration += p.duration
                done = True
                break
        if not done:
            merged.append(p)
    persons = merged

    # ---------- LEVEL 3 & 4: ROLL BASED (100% RULE) ----------
    final = []
    for p in persons:
        merged = False
        for f in final:
            if p.roll and f.roll and p.roll == f.roll:
                tokens = set(p.name.split()) & set(f.name.split())
                name_match = any(len(t) >= 4 for t in tokens)

                if name_match:
                    f.duration += p.duration
                    merged = True
                    break
                else:
                    if (p.duration + f.duration) < meeting_minutes:
                        f.duration += p.duration
                        merged = True
                        break

        if not merged:
            final.append(p)

    # ---------- FINAL TABLE ----------
    rows = []
    for p in final:
        rows.append({
            "Name (original name)": p.display,
            "Duration (minutes)": round(p.duration, 2),
            "First Join": p.first_join,
            "_PRIORITY": p.priority_name()
        })

    out = pd.DataFrame(rows)

    out = out.sort_values(
        by=["_PRIORITY", "Name (original name)"],
        na_position="last"
    )
    out.drop(columns="_PRIORITY", inplace=True)


    absent = out[out["Duration (minutes)"] < absent_limit]

    # ---------- LATE ----------
    cutoff = datetime.strptime(f"{late_time} {late_ampm}", "%I:%M:%S %p").time()
    late = out[out["First Join"].dt.time > cutoff].copy()
    late = late.sort_values(
        by=["Name (original name)", "First Join"],
        ascending=[True, True],
        na_position="last"
    )
    late["Join Time"] = late["First Join"].apply(format_datetime_12h)
    late["Status"] = "LATE"
    late = late[["Name (original name)", "Join Time", "Status"]]

    # ---------- RAW ----------
    raw = df_original[[name_col, join_col, leave_col, dur_col]].copy()
    raw["Join Time"] = df_original["_JOIN_DT"].apply(format_datetime_12h)
    raw.rename(columns={name_col: "Name (original name)"}, inplace=True)

    raw["_JOIN_SORT"] = df_original["_JOIN_DT"]
    raw = raw.sort_values(
        by=["Name (original name)", "_JOIN_SORT"],
        ascending=[True, True],
        na_position="last"
    )
    raw = raw[["Name (original name)", "Join Time", leave_col, dur_col]]

    # ---------- CF ----------
    cf["Join Time"] = cf["_JOIN_DT"].apply(format_datetime_12h)
    cf = cf[[name_col, "Join Time", leave_col, dur_col]]
    cf.rename(columns={name_col: "Name (original name)"}, inplace=True)

    # ---------- WRITE ----------
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        raw.to_excel(writer, sheet_name="RAW", index=False)

        if do_att:
            out[["Name (original name)", "Duration (minutes)"]].to_excel(
                writer, sheet_name="ATT", index=False
            )
            if not absent.empty:
                absent[["Name (original name)", "Duration (minutes)"]].to_excel(
                    writer, sheet_name="ABSENT", index=False
                )

        if do_late:
            late.to_excel(writer, sheet_name="LATE", index=False)

        if not cf.empty:
            cf.to_excel(writer, sheet_name="CF", index=False)

    # ---------- FORMAT & HIGHLIGHT ----------
    wb = load_workbook(out_path)
    for s in wb.sheetnames:
        format_table(wb[s])

    # RAW highlights (respect user selections)
    ws = wb["RAW"]
    for r in range(2, ws.max_row + 1):

        # ---- LATE highlight (Join Time) ----
        if do_late:
            try:
                t = datetime.strptime(
                    ws[f"B{r}"].value,
                    "%m/%d/%Y %I:%M:%S %p"
                ).time()
                ws[f"B{r}"].fill = GREEN if t <= cutoff else RED
            except:
                pass

        # ---- ATT / ABSENT highlight (Duration) ----
        if do_att:
            try:
                d = float(ws[f"D{r}"].value)
                ws[f"D{r}"].fill = (
                    YELLOW if d < absent_limit else
                    RED if d < total_att else
                    GREEN
                )
            except:
                pass

    # ATT / ABSENT highlights
    if do_att:
        ws = wb["ATT"]
        for r in range(2, ws.max_row + 1):
            d = float(ws[f"B{r}"].value)
            ws[f"B{r}"].fill = (
                YELLOW if d < absent_limit else
                RED if d < total_att else
                GREEN
            )
        if "ABSENT" in wb.sheetnames:
            ws = wb["ABSENT"]
            for r in range(2, ws.max_row + 1):
                ws[f"B{r}"].fill = YELLOW

    # LATE highlights
    if do_late and "LATE" in wb.sheetnames:
        ws = wb["LATE"]
        for r in range(2, ws.max_row + 1):
            ws[f"B{r}"].fill = RED
            ws[f"C{r}"].fill = RED

    wb.save(out_path)
