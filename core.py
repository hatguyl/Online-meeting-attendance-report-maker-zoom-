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


# +--------+
# | REGEX  |
# +--------+
ROLL_PATTERN = re.compile(r'\b\d{2}[A-Z]{2}\d{3}\b', re.IGNORECASE)
NON_WORD_RE = re.compile(r"[^\w\s]", re.UNICODE)
APOSTROPHE_RE = re.compile(r"[’‘`]")
PAREN_RE = re.compile(r"\(.*?\)")

#BLACKLISTED WORDS

BLACKLIST_WORDS = {"user"}
TITLE_WORDS = {"guest"}


# +----------------+
# | NAME PARSING   |
# +----------------+
def extract_roll(name):
    if not name:
        return None
    m = ROLL_PATTERN.search(str(name))
    return m.group(0).upper() if m else None


def normalize_name(name):
    if not name:
        return ""
    s = PAREN_RE.sub("", str(name))
    s = NON_WORD_RE.sub("", s)
    return s.strip().lower()


def clean_display_name(name):
    if not name:
        return ""
    s = PAREN_RE.sub("", str(name))
    s = ROLL_PATTERN.sub("", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


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

    # +----------------+
    # | ID RESOLUTION  |
    # +----------------+
    df["_ROLL"] = df[name_col].apply(extract_roll)
    df["_NAME_KEY"] = df[name_col].apply(normalize_name)

    name_to_roll = (
        df.loc[df["_ROLL"].notna(), ["_NAME_KEY", "_ROLL"]]
        .drop_duplicates("_NAME_KEY")
        .set_index("_NAME_KEY")["_ROLL"]
        .to_dict()
    )

    df["_ROLL"] = df.apply(
        lambda r: r["_ROLL"] if pd.notna(r["_ROLL"]) else name_to_roll.get(r["_NAME_KEY"]),
        axis=1
    )

    df["_GROUP_KEY"] = df["_ROLL"].fillna("NO_ROLL_" + df["_NAME_KEY"].fillna("unknown"))
    df["_DISPLAY_NAME"] = df[name_col].apply(clean_display_name)

    cutoff = datetime.strptime(
        f"{late_time} {late_ampm}", "%I:%M:%S %p"
    ).time()

    blacklist_mask = df[name_col].apply(is_blacklisted)

    cf = df.loc[blacklist_mask].copy()
    cf["Join Time"] = cf["_JOIN_FMT"]

    df = df.loc[~blacklist_mask].copy()

    raw = df[["_DISPLAY_NAME", dur_col]].copy()
    raw["Join Time"] = df["_JOIN_FMT"]

    att = (
        df.groupby("_GROUP_KEY", as_index=False)
        .agg({"_DISPLAY_NAME": "first", dur_col: "sum"})
    )

    absent = att[att[dur_col] < absent_limit].copy()
    has_absent = not absent.empty

    late = (
        df.sort_values("_JOIN_DT")
        .groupby("_GROUP_KEY", as_index=False)
        .first()
    )
    late = late[late["_JOIN_DT"].dt.time > cutoff]
    late["Join Time"] = late["_JOIN_FMT"]
    late["Status"] = "LATE"

    # +----------------+
    # | WRITE EXCEL    |
    # +----------------+
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        raw[["_DISPLAY_NAME", "Join Time", dur_col]].to_excel(
            writer, sheet_name="RAW", index=False
        )

        if do_att:
            att[["_DISPLAY_NAME", dur_col]].to_excel(
                writer, sheet_name="ATT", index=False
            )

            if has_absent:
                absent[["_DISPLAY_NAME", dur_col]].to_excel(
                    writer, sheet_name="ABSENT", index=False
                )

        if do_late:
            late[["_DISPLAY_NAME", "Join Time", "Status"]].to_excel(
                writer, sheet_name="LATE", index=False
            )

        cf[["_DISPLAY_NAME", "Join Time", leave_col, dur_col]].to_excel(
            writer, sheet_name="CF", index=False
        )

    # +----------------+
    # | POST FORMAT    |
    # +----------------+
    wb = load_workbook(out_path)

    ws = wb["RAW"]
    format_table(ws)

    for r, dt in enumerate(df["_JOIN_DT"], start=2):
        if pd.notna(dt):
            ws[f"B{r}"].fill = GREEN if dt.time() <= cutoff else RED

    if do_att:
        ws = wb["ATT"]
        format_table(ws)
        for r in range(2, ws.max_row + 1):
            d = ws[f"B{r}"].value
            ws[f"B{r}"].fill = (
                YELLOW if d < absent_limit else
                RED if d < total_att else
                GREEN
            )

        if has_absent:
            ws = wb["ABSENT"]
            format_table(ws)
            for r in range(2, ws.max_row + 1):
                ws[f"B{r}"].fill = YELLOW

    if do_late:
        ws = wb["LATE"]
        format_table(ws)
        for r in range(2, ws.max_row + 1):
            ws[f"B{r}"].fill = RED
            ws[f"C{r}"].fill = RED

    ws = wb["CF"]
    format_table(ws)

    wb.save(out_path)
