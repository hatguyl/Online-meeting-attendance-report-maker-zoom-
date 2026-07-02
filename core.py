import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
import re


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
ROLL_PATTERN = re.compile(r'([A-Z]{1,3}\s*[_\-]?\s*\d{2,4})', re.IGNORECASE)
NON_WORD_RE = re.compile(r"[^\w\s]", re.UNICODE)
APOSTROPHE_RE = re.compile(r"[’‘`]")
PAREN_RE = re.compile(r"\(.*?\)")

BLACKLIST_WORDS = {"Zoom User"}
TITLE_WORDS = {"Host"}


# +----------------+
# | NAME PARSING   |
# +----------------+
def name_prefix(name, n=5):
    if not name:
        return ""
    return name.replace(" ", "")[:n]

def extract_roll(name):
    if not name:
        return None

    s = str(name).replace("{", "").replace("}", "")
    m = ROLL_PATTERN.search(s)
    if not m:
        return None

    raw = m.group(1).upper()
    raw = raw.replace("_", "").replace("-", "").replace(" ", "")

    m2 = re.match(r"([A-Z]+)([A-Z0-9]+)", raw)
    if not m2:
        return None
    letters, tail = m2.groups()

    tail = (tail.replace("O", "0"))

    if not tail.isdigit():
        return None
    return f"{letters}{tail.zfill(3)}"

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

def bracket_links(name_key, alias_key):
    if not name_key or not alias_key:
        return False
    return name_key == alias_key


def names_related(name1, name2):
    if not name1 or not name2:
        return False

    if name1 in name2 or name2 in name1:
        return True

    p1 = set(name1.split())
    p2 = set(name2.split())

    return any(len(t) >= 4 for t in p1 & p2)


def is_blacklisted(name):
    if not name:
        return False
    s = APOSTROPHE_RE.sub("'", str(name).lower())
    if any(w in s for w in BLACKLIST_WORDS):
        return True
    parts = s.split()
    return len(parts) >= 2 and parts[-1] in TITLE_WORDS


# +----------------+
# | CSV UTILITIES  |
# +----------------+
def find_column(df, keywords):
    for col in df.columns:
        if any(k in col.lower() for k in keywords):
            return col
    raise ValueError(f"Missing required column: {keywords}")


# +----------------+
# | TIME HANDLING  |
# +----------------+
def parse_datetime(series):
    return pd.to_datetime(series, errors="coerce", format="mixed")


def format_datetime_12h(dt):
    if pd.isna(dt):
        return ""
    return dt.strftime("%m/%d/%Y %I:%M:%S %p")


# +----------------+
# | TABLE FORMAT   |
# +----------------+
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
    df = pd.read_csv(csv_path)
    df.columns = df.columns.str.strip()

    name_col = find_column(df, {"name", "participant", "user"})
    join_col = find_column(df, {"join"})
    leave_col = find_column(df, {"leave"})
    dur_col = find_column(df, {"duration"})

    df["_JOIN_DT"] = parse_datetime(df[join_col])
    df["_JOIN_FMT"] = df["_JOIN_DT"].apply(format_datetime_12h)
    df_original = df.copy()   

    # ---------- ID RESOLUTION ----------
    df["_ROLL"] = df[name_col].apply(extract_roll)
    df["_NAME_KEY"] = df[name_col].apply(normalize_name)
    df["_ALIAS_KEY"] = df[name_col].apply(extract_alias)

    # +---------------------------+
    # | GROUP KEY RESOLUTION      |
    # +---------------------------+

    group_map = []   
    final_keys = []

    for _, r in df.iterrows():
        roll   = r["_ROLL"]
        name   = r["_NAME_KEY"]
        alias  = r["_ALIAS_KEY"]

        matched = None

        if roll:
            for g in group_map:
                if g["roll"] == roll:
                    matched = g
                    break

        if matched is None and alias:
            for g in group_map:
                if g["name"] == alias:
                    matched = g
                    break

        if matched is None:
            matched = {
                "roll": roll,
                "name": alias or name   
            }
            group_map.append(matched)

        if matched["roll"]:
            final_keys.append(f"{matched['roll']}::{matched['name']}")
        else:
            final_keys.append(matched["name"])

    df["_GROUP_KEY"] = final_keys


    def pick_display_name(group, name_col):
        names = group[name_col].dropna().astype(str)

        bracketed = names[names.str.contains(r"\(.*\)", regex=True)]
        if not bracketed.empty:

            return bracketed.iloc[bracketed.str.len().argmax()]

        return names.iloc[names.str.len().argmax()]


    cutoff = datetime.strptime(
        f"{late_time} {late_ampm}", "%I:%M:%S %p"
    ).time()

    # BLACKLIST 
    blacklist_mask = df[name_col].apply(is_blacklisted)
    cf = df.loc[blacklist_mask].copy()
    cf["Join Time"] = cf["_JOIN_FMT"]
    df = df.loc[~blacklist_mask].copy()

    # RAW
    raw = df_original[[name_col, join_col, leave_col, dur_col]].copy()
    raw["Join Time"] = df_original["_JOIN_FMT"]
    raw.rename(columns={name_col: "Name (original name)"}, inplace=True)

    raw["_JOIN_SORT"] = df_original["_JOIN_DT"]
    raw = raw.sort_values(
        by=["Name (original name)", "_JOIN_SORT"],
        ascending=[True, True],
        na_position="last"
    )

    raw = raw[["Name (original name)", "Join Time", leave_col, dur_col]]

    name_map = (
        df.groupby("_GROUP_KEY", group_keys=False)
        .apply(lambda g: pick_display_name(g, name_col), include_groups=False)
    )

    # ATT 
    att = (
        df.groupby("_GROUP_KEY", group_keys=False)
        .agg({dur_col: "sum"})
        .reset_index()
    )

    att["Name (original name)"] = att["_GROUP_KEY"].map(name_map)

    att[dur_col] = pd.to_numeric(att[dur_col], errors="coerce").fillna(0)

    att = att.sort_values(
        by="Name (original name)",
        ascending=True,
        na_position="last"
    )

    # ABSENT 
    absent = att[att[dur_col] < absent_limit].copy()

    # LATE

    late = (
        df.sort_values("_JOIN_DT")
        .groupby("_GROUP_KEY", as_index=False)
        .apply(
            lambda g:
                g.loc[g["_ROLL"].notna()].iloc[0]
                if g["_ROLL"].notna().any()
                else g.iloc[0],
            include_groups=False
        )
    )

    late["Name (original name)"] = late["_GROUP_KEY"].map(name_map)
    late = late[late["_JOIN_DT"].dt.time > cutoff]
    late["Join Time"] = late["_JOIN_FMT"]
    late["Status"] = "LATE"
    late = late.sort_values(
        by="Name (original name)",
        ascending=True,
        na_position="last"
    )
    late = late[["Name (original name)", "Join Time", "Status"]]


    # WRITE 
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        raw.to_excel(writer, sheet_name="RAW", index=False)

        if do_att:
            att[["Name (original name)", dur_col]].to_excel(writer, sheet_name="ATT", index=False)

            if not absent.empty:
                absent[["Name (original name)", dur_col]].to_excel(
                    writer, sheet_name="ABSENT", index=False
                )
            
        if do_late:
            late.to_excel(writer, sheet_name="LATE", index=False)

        cf[[name_col, "Join Time", leave_col, dur_col]].rename(
            columns={name_col: "Name (original name)"}
        ).to_excel(writer, sheet_name="CF", index=False)


    # POST FORMAT 
    wb = load_workbook(out_path)

    for sheet in wb.sheetnames:
        format_table(wb[sheet])

    # RAW highlighting
    ws = wb["RAW"]
    for r in range(2, ws.max_row + 1):
        jt = ws[f"B{r}"].value
        dv = ws[f"D{r}"].value

        if jt:
            try:
                t = datetime.strptime(jt, "%m/%d/%Y %I:%M:%S %p").time()
                ws[f"B{r}"].fill = GREEN if t <= cutoff else RED
            except Exception:
                pass

        try:
            d = float(dv)
            ws[f"D{r}"].fill = (
                YELLOW if d < absent_limit else
                RED if d < total_att else
                GREEN
            )
        except Exception:
            pass

    # ATT / ABSENT HIGHLIGHTER 
    if do_att:
        ws = wb["ATT"]
        for r in range(2, ws.max_row + 1):
            try:
                d = float(ws[f"B{r}"].value)
                ws[f"B{r}"].fill = (
                    YELLOW if d < absent_limit else
                    RED if d < total_att else
                    GREEN
                )
            except Exception:
                pass

        if "ABSENT" in wb.sheetnames:
            ws = wb["ABSENT"]
            for r in range(2, ws.max_row + 1):
                ws[f"B{r}"].fill = YELLOW

    # LATE HIGHLIGHTER 
    if do_late:
        ws = wb["LATE"]
        for r in range(2, ws.max_row + 1):
            ws[f"B{r}"].fill = RED
            ws[f"C{r}"].fill = RED

    wb.save(out_path)
