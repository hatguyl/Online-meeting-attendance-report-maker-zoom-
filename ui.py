import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
import datetime
from tkinter import filedialog, messagebox
from ui2 import MultiCSVConfirm
from PIL import Image, ImageTk
import os, sys, threading
import subprocess
import pandas as pd 
from core import (
    process,
    extract_att_data,
    extract_late_data,
    export_to_csv,
    combine_csv_files,
    group_csv_files,
    extract_meeting_metadata
)


# +-------------------+
# | RESOURCE HANDLING |
# +-------------------+
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# +----------------+
# | UI INTERFACE   |
# +----------------+
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        # +----------------+
        # | WINDOW CONFIG  |
        # +----------------+
        self.title("Report tool")
        self.geometry("400x540")
        self.resizable(False, False)
        self.minsize(400, 540)
        self.maxsize(400, 540)
        self.configure(bg="#111111")

        # +----------------+
        # | COLOR PALETTE  |
        # +----------------+
        self.BG = "#111111"
        self.CARD = "#1c1c1c"
        self.FIELD = "#2a2a2a"
        self.FIELD_HOVER = "#343434"
        self.TEXT = "#ffffff"
        self.MUTED = "#bdbdbd"
        self.ACCENT = "#f4c430"

        # +----------------+
        # | STATE VARIABLES|
        # +----------------+
        self.csv_files = []
        self.do_att = ctk.BooleanVar(value=True)
        self.do_late = ctk.BooleanVar(value=True)
        self.do_csv = ctk.BooleanVar(value=True)

        self.batch = ctk.StringVar(value="NOT SELECTED")
        self.late_time = ctk.StringVar(value="06:05:55")
        self.late_ampm = ctk.StringVar(value="AM")

        # +----------------+
        # | MAIN CARD      |
        # +----------------+
        self.card = ctk.CTkFrame(self, corner_radius=22, fg_color=self.CARD)
        self.card.pack(padx=14, pady=12, fill="both", expand=False)

        # +----------------+
        # | LOGO & ICON    |
        # +----------------+
        logo_img = Image.open(resource_path("mentorbeelogo.png")).resize((160, 55), Image.LANCZOS)
        self.logo = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(160, 55))

        icon_img = Image.open(resource_path("mbeelogo.png")).resize((64, 64), Image.LANCZOS)
        self.iconphoto(False, ImageTk.PhotoImage(icon_img))

        ctk.CTkLabel(self.card, image=self.logo, text="").pack(pady=(10, 5))

        # +----------------+
        # | CSV DROP FILE  |
        # +----------------+
        self.drop = ctk.CTkFrame(self.card, height=95, corner_radius=16, fg_color=self.FIELD)
        self.drop.pack(padx=20, pady=12, fill="x")

        self.drop_label = ctk.CTkLabel(
            self.drop,
            text="📄\nDrop Zoom generated CSV here\nor click to browse",
            font=("Segoe UI", 13),
            text_color=self.MUTED,
            justify="center"
        )
        self.drop_label.pack(expand=True, pady=10)

        self.drop_label.bind("<Button-1>", lambda e: self.browse())
        self.drop_label.bind("<Enter>", lambda e: self.drop.configure(fg_color=self.FIELD_HOVER))
        self.drop_label.bind("<Leave>", lambda e: self.drop.configure(fg_color=self.FIELD))

        self.drop.drop_target_register(DND_FILES)
        self.drop.dnd_bind("<<Drop>>", self.on_drop)

        # +-----------------------------+
        # | GLOBAL DRAG & DROP SUPPORT  |
        # +-----------------------------+
        for widget in (self, self.card, self.drop_label):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self.on_drop)


        # +----------------+
        # | MODE TOGGLES   |
        # +----------------+
        chk_row = ctk.CTkFrame(self.card, fg_color="transparent")
        chk_row.pack(pady=8)

        ctk.CTkCheckBox(
            chk_row,
            text="CSV",
            variable=self.do_csv,
            checkbox_width=20, checkbox_height=20,
            fg_color=self.ACCENT,
            hover_color="#ffd95a"
        ).pack(side="left", padx=10)

        ctk.CTkCheckBox(
            chk_row,
            text="ATT",
            variable=self.do_att,
            checkbox_width=20, checkbox_height=20,
            fg_color=self.ACCENT,
            hover_color="#ffd95a",
            command=self.update_input_states
        ).pack(side="left", padx=10)

        ctk.CTkCheckBox(
            chk_row,
            text="LATE",
            variable=self.do_late,
            checkbox_width=20, checkbox_height=20,
            fg_color=self.ACCENT,
            hover_color="#ffd95a",
            command=self.update_input_states
        ).pack(side="left", padx=10)

        # +----------------+
        # | BATCH SELECT   |
        # +----------------+
        ctk.CTkLabel(
            self.card,
            text="Batch Timing",
            text_color=self.MUTED,
            anchor="center"
        ).pack(pady=(10, 5))

        self.batch_menu = ctk.CTkOptionMenu(
            self.card,
        values=[
            "AUTOMATIC",
            "NOT SELECTED",
            "MRNG (6 to 7:30)",
            "MRNG (6 to 8)",
            "EVNG (7 to 8:30)",     
            "EVNG (7 to 9:30)",
            "EVNG (7:30 to 9:30)",
            "3 HOURS CLASS",     
            "REVISE & EXAM",
            "OTHERS"
        ],
            variable=self.batch,
            command=self.apply_batch,
            fg_color=self.FIELD,
            button_color=self.ACCENT,
            button_hover_color="#ffd95a"
        )
        self.batch_menu.pack()
        self.batch_menu.configure(state="disabled")

        # +----------------+
        # | ATT / ABS BOX  |
        # +----------------+
        ctk.CTkLabel(self.card, text="Attendance  &  Absent Due (minutes)", text_color=self.MUTED)\
            .pack(pady=(6, 0))

        att_row = ctk.CTkFrame(self.card, fg_color="transparent")
        att_row.pack(padx=30, pady=4)

        self.att_entry = ctk.CTkEntry(
            att_row,
            height=34,
            width=100,
            fg_color=self.FIELD,
            border_color="#555555",
            justify="center"   
        )

        self.att_entry.pack(side="left", padx=(0, 12))
        self.att_entry.insert(0, "75")

        self.abs_entry = ctk.CTkEntry(
            att_row,
            height=34,
            width=100,
            fg_color=self.FIELD,
            border_color="#555555",
            justify="center"
        )
        self.abs_entry.pack(side="left")
        self.abs_entry.insert(0, "10")
       
       
        # +----------------+
        # | ISSUED DATE    |
        # +----------------+
        self.issued_date_label = ctk.CTkLabel(
            self.card,
            text="Issued Date: —",
            text_color=self.MUTED,
            font=("Segoe UI", 12),
            anchor="center"
        )
        self.issued_date_label.pack(pady=(2, 2))

        # +----------------+
        # | LATE INPUTS    |
        # +----------------+
        ctk.CTkLabel(
            self.card,
            text="Late Cutoff Time (HH:MM:SS)",
            text_color=self.MUTED
        ).pack(pady=(4, 0))

        time_row = ctk.CTkFrame(self.card, fg_color="transparent")
        time_row.pack(padx=30, pady=4)

        self.time_entry = ctk.CTkEntry(
            time_row,
            textvariable=self.late_time,
            height=36,
            width=140,
            fg_color=self.FIELD,
            border_color="#555555",
            justify="center"   
        )

        self.time_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.ampm = ctk.CTkOptionMenu(
            time_row, values=["AM", "PM"],
            variable=self.late_ampm, width=80,
            fg_color=self.FIELD,
            button_color=self.ACCENT,
            button_hover_color="#ffd95a"
        )
        self.ampm.pack(side="right")

        # +----------------+
        # | ACTION BUTTON  |
        # +----------------+
        ctk.CTkButton(
            self.card,
            text="GENERATE REPORT",
            height=40,
            fg_color=self.ACCENT,
            hover_color="#ffd95a",
            text_color="#000000",
            command=self.run
        ).pack(pady=(8, 10), fill="x", padx=30)
        self.update_input_states()

        # +----------------+
        # | VERSION LABEL  |
        # +----------------+
        self.version_label = ctk.CTkLabel(
            self,
            text="v3.2.1",
            font=("Segoe UI", 9),
            text_color=self.MUTED
        )

        self.version_label.place(
            relx=1.0,
            rely=1.0,
            anchor="se",
            x=-12,
            y=-10
        )

        # +----------------+
        # | SAFE SHUTDOWN  |
        # +----------------+
        self.protocol("WM_DELETE_WINDOW", self.safe_close)

    # +----------------+
    # | SAFE CLOSE     |
    # +----------------+
    def safe_close(self):
        self.destroy()
    
    def open_file(self, path):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.call(["open", path])
            else:
                import subprocess
                subprocess.call(["xdg-open", path])
        except Exception as e:
            print("Open failed:", e)

    # +----------------+
    # | UI HELPERS     |
    # +----------------+
    def update_input_states(self):
        state = "normal" if self.do_att.get() else "disabled"
        self.att_entry.configure(state=state)
        self.abs_entry.configure(state=state)

        late_state = "normal" if self.do_late.get() else "disabled"
        self.time_entry.configure(state=late_state)
        self.ampm.configure(state=late_state)

    def labeled_entry(self, label):
        ctk.CTkLabel(self.card, text=label, text_color=self.MUTED).pack(
            anchor="w", padx=30, pady=(6, 0)
        )
        entry = ctk.CTkEntry(
            self.card,
            height=36,
            fg_color=self.FIELD,
            border_color="#555555"
        )
        entry.pack(fill="x", padx=30)
        return entry
    
    def update_issued_date(self, csv_path):
        try:
            df = combine_csv_files([csv_path])
            join_col = None
            for c in df.columns:
                if "join" in c.lower():
                    join_col = c
                    break

            if not join_col:
                self.issued_date_label.configure(text="Issued Date: —")
                return

            join_times = pd.to_datetime(
                df[join_col], errors="coerce", format="mixed"
            ).dropna()

            if join_times.empty:
                self.issued_date_label.configure(text="Issued Date: —")
                return

            # ---------- DATE ----------
            first_date = join_times.iloc[0].date()
            today = pd.Timestamp.now().date()

            date_str = join_times.iloc[0].strftime("%d/%m/%Y")

            if first_date != today:
                self.issued_date_label.configure(
                    text=f"⚠️ Issued Date: {date_str}",
                    text_color="#FF4C4C"   # red
                )
            else:
                self.issued_date_label.configure(
                    text=f"Issued Date: {date_str}",
                    text_color=self.MUTED
                )

            # ---------- AM/PM AUTO DETECT ----------
            hours = join_times.dt.hour

            am_count = (hours < 12).sum()
            pm_count = (hours >= 12).sum()

            if pm_count > am_count:
                self.late_ampm.set("PM")
            else:
                self.late_ampm.set("AM")
        except Exception:
            self.issued_date_label.configure(text="Issued Date: —")
    
    def auto_fill_from_csv(self, csv_path):
        try:
            df = combine_csv_files([csv_path])

            # ---------- FIND COLUMNS ----------
            join_col = next(
                (c for c in df.columns if "join" in c.lower()),
                None
            )

            leave_col = next(
                (c for c in df.columns if "leave" in c.lower()),
                None
            )

            dur_col = next(
                (c for c in df.columns if "duration" in c.lower()),
                None
            )

            if not join_col or not leave_col:
                return

            # ---------- PARSE ----------
            join_times = pd.to_datetime(
                df[join_col],
                errors="coerce",
                format="mixed"
            ).dropna()

            leave_times = pd.to_datetime(
                df[leave_col],
                errors="coerce",
                format="mixed"
            ).dropna()

            durations = pd.to_numeric(
                df[dur_col],
                errors="coerce"
            ).dropna()

            if join_times.empty or leave_times.empty:
                return

            # ==========================================
            # EARLIEST STRONG CLUSTER
            # ==========================================

            join_floor = join_times.dt.floor("min")

            counts = (
                join_floor.value_counts()
            )

            participant_count = len(join_times)

            if participant_count <= 15:
                threshold = 1
            elif participant_count <= 40:
                threshold = 2
            else:
                threshold = max(3, int(participant_count * 0.03))

            strong_clusters = counts[
                counts >= threshold
            ]

            if not strong_clusters.empty:

                dominant = strong_clusters.index.min()

            else:

                dominant = counts.sort_index().index[0]

            # ==========================================
            # ROUND TO VALID SLOT
            # ==========================================

            hour = dominant.hour
            minute = dominant.minute

            if minute < 15:
                snapped_minute = 0

            elif minute < 45:
                snapped_minute = 30

            else:
                snapped_minute = 0
                hour += 1

            class_start = dominant.replace(
                hour=hour,
                minute=snapped_minute,
                second=0
            )

            # ==================================================
            # SMART DUE DETECTION
            # ==================================================

            top_users = durations.sort_values(ascending=False)

            top_count = max(5, int(len(top_users) * 0.15))

            strong_duration = top_users.head(top_count).mean()

            duration_slots = [
                (60,  "1 HOUR CLASS"),
                (90,  "1.5 HOUR CLASS"),
                (120, "2 HOUR CLASS"),
                (150, "2.5 HOUR CLASS"),
                (180, "3 HOUR CLASS"),
                (210, "3.5 HOUR CLASS"),
                (240, "4 HOUR CLASS"),
                (270, "4.5 HOUR CLASS"),
            ]

            class_minutes = min(
                duration_slots,
                key=lambda x: abs(x[0] - strong_duration)
            )[0]

            batch_name = dict(duration_slots)[class_minutes]

            # ==================================================
            # LATE CUTOFF
            # ==================================================

            late_cutoff = class_start + pd.Timedelta(
                minutes=5,
                seconds=55
            )

            # ==================================================
            # ATTENDANCE RULES
            # ==================================================

            attendance_rules = {
                60:  (50, 5),
                90:  (75, 10),
                120: (100, 15),
                150: (130, 20),
                180: (140, 30),
                210: (170, 30),
                240: (200, 40),
                270: (230, 40),
            }

            attendance, absent = attendance_rules[class_minutes]

            # ==================================================
            # APPLY TO UI
            # ==================================================

            if len(self.csv_files) <= 1:
                self.batch.set(batch_name)

            self.late_time.set(
                late_cutoff.strftime("%I:%M:%S")
            )

            self.late_ampm.set(
                late_cutoff.strftime("%p")
            )

            self.att_entry.delete(0, "end")
            self.att_entry.insert(0, str(attendance))

            self.abs_entry.delete(0, "end")
            self.abs_entry.insert(0, str(absent))

        except Exception as e:
            print("Auto-fill failed:", e)

    # +----------------+
    # | FILE HANDLING  |
    # +----------------+
    def browse(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("CSV Files", "*.csv")]
        )

        if paths:
            self.csv_files = list(paths)
            if len(self.csv_files) > 1:
                self.batch.set("AUTOMATIC")

            count = len(self.csv_files)

            if count == 1:
                label = os.path.basename(self.csv_files[0])
            else:
                label = (
                    "1 CSV file selected"
                    if count == 1
                    else f"{count} CSV files selected"
                )

            self.drop_label.configure(
                text=f"📄\n{label}"
            )

            self.update_issued_date(self.csv_files[0])
            self.auto_fill_from_csv(self.csv_files[0])

            self.batch_menu.configure(state="normal")

    def on_drop(self, event):

        files = self.tk.splitlist(event.data)

        valid = []

        for path in files:
            path = path.strip()

            if (
                path.lower().endswith(".csv")
                and os.path.exists(path)
            ):
                valid.append(path)

        if not valid:
            messagebox.showerror(
                "Invalid file",
                "Please drop Zoom-generated CSV files only."
            )
            return

        self.csv_files = valid
        if len(self.csv_files) > 1:
            self.batch.set("AUTOMATIC")

        count = len(valid)

        if count == 1:
            label = os.path.basename(valid[0])
        else:
            label = f"{count} CSV files selected"

        self.drop_label.configure(
            text=f"📄\n{label}"
        )

        self.update_issued_date(valid[0])
        self.auto_fill_from_csv(valid[0])

        self.batch_menu.configure(state="normal")

    # +----------------+
    # | BATCH LOGIC    |
    # +----------------+
    def apply_batch(self, _):
        if self.batch.get() == "NOT SELECTED":
            return
        self.do_late.set(True)
        self.abs_entry.delete(0, "end")
        self.abs_entry.insert(0, "10")

        if self.batch.get() == "MRNG (6 to 7:30)":
            self.att_entry.delete(0, "end")
            self.att_entry.insert(0, "75")
            self.late_time.set("06:05:55")
            self.late_ampm.set("AM")
        elif self.batch.get() == "MRNG (6 to 8)":
            self.att_entry.delete(0, "end")
            self.att_entry.insert(0, "100")
            self.late_time.set("06:05:55")
            self.late_ampm.set("AM")
        elif self.batch.get() == "EVNG (7:30 to 9:30)":
            self.att_entry.delete(0, "end")
            self.att_entry.insert(0, "100")
            self.late_time.set("07:35:55")
            self.late_ampm.set("PM")
        elif self.batch.get() == "EVNG (7 to 9:30)":
            self.att_entry.delete(0, "end")
            self.att_entry.insert(0, "130")
            self.late_time.set("07:05:55")
            self.late_ampm.set("PM")
        elif self.batch.get() == "EVNG (7 to 8:30)":
            self.att_entry.delete(0, "end")
            self.att_entry.insert(0, "70")
            self.late_time.set("07:05:55")
            self.late_ampm.set("PM")
        elif self.batch.get() == "3 HOURS CLASS":
            self.att_entry.delete(0, "end")
            self.att_entry.insert(0, "140")
            self.abs_entry.delete(0, "end")
            self.late_time.set("12:00:00")
            self.late_ampm.set("AM")
            self.abs_entry.insert(0, "30")
        elif self.batch.get() == "REVISE & EXAM":
            self.att_entry.delete(0, "end")
            self.att_entry.insert(0, "0")
            self.abs_entry.delete(0, "end")
            self.abs_entry.insert(0, "0")
            self.do_late.set(False)

        self.update_input_states()

    # +----------------+
    # | MAIN EXEC      |
    # +----------------+
    
    # +---------------------------+
    # | SHARED LATE CUTOFF LOGIC  |
    # +---------------------------+
    def _compute_late_cutoff(self, csv_group):
        """Return (late_time_str, late_ampm_str) using the same cluster logic as auto_fill."""
        try:
            df = combine_csv_files(csv_group)
            join_col = next((c for c in df.columns if "join" in c.lower()), None)
            if not join_col:
                return self.late_time.get(), self.late_ampm.get()
            join_times = pd.to_datetime(df[join_col], errors="coerce", format="mixed").dropna()
            if join_times.empty:
                return self.late_time.get(), self.late_ampm.get()
            counts = join_times.dt.floor("min").value_counts()
            n = len(join_times)
            threshold = 1 if n <= 15 else 2 if n <= 40 else max(3, int(n * 0.03))
            strong = counts[counts >= threshold]
            dominant = strong.index.min() if not strong.empty else counts.sort_index().index[0]
            h, m = dominant.hour, dominant.minute
            if m < 15:   snapped = 0
            elif m < 45: snapped = 30
            else:        snapped = 0; h += 1
            cutoff = dominant.replace(hour=h, minute=snapped, second=0) + pd.Timedelta(minutes=5, seconds=55)
            return cutoff.strftime("%I:%M:%S"), cutoff.strftime("%p")
        except Exception:
            return self.late_time.get(), self.late_ampm.get()

    def open_multi_confirm(self, groups, _save_jobs):

        reports = []
        for group in groups:
            try:
                meta  = extract_meeting_metadata(group[0])
                topic = (
                    str(meta.get("topic", "REPORT"))
                    .replace("MENTORBEE", "").replace("EDUVERSE", "").strip()
                )

                # ---------- AUTO DETECT ATTENDANCE ----------
                temp_df  = combine_csv_files(group)
                dur_col  = next((c for c in temp_df.columns if "duration" in c.lower()), None)
                attendance, absent = 75, 10
                if dur_col:
                    durations = pd.to_numeric(temp_df[dur_col], errors="coerce").dropna()
                    if not durations.empty:
                        top_count      = max(5, int(len(durations) * 0.15))
                        strong_dur     = durations.sort_values(ascending=False).head(top_count).mean()
                        for thresh, att, ab in [
                            (75, 50, 5), (105, 75, 10), (135, 100, 15), (165, 130, 20)
                        ]:
                            if strong_dur <= thresh:
                                attendance, absent = att, ab
                                break
                        else:
                            attendance, absent = 140, 30

                # ---------- ACCURATE LATE CUTOFF ----------
                late_time, late_ampm = self._compute_late_cutoff(group)

                reports.append({
                    "group": group, "topic": topic,
                    "attendance": attendance, "absent": absent,
                    "late_time": late_time, "late_ampm": late_ampm,
                })
            except Exception as e:
                print("Config build failed:", e)

        # ==========================================
        # CALLBACK — returns True only when all saves confirmed (bug 7 fix)
        # ==========================================
        def generate_callback(configs):
            save_jobs = []
            for cfg in configs:
                group = cfg["group"]
                try:
                    meta  = extract_meeting_metadata(group[0])
                    topic = (
                        str(meta.get("topic", "REPORT"))
                        .replace("MENTORBEE", "").replace("EDUVERSE", "")
                        .replace("Mentorbee", "").replace("Eduverse", "")
                        .replace("/", "-").replace("\\", "-").replace(":", "-")
                        .replace("*", "").replace("?", "").replace('"', "")
                        .replace("<", "").replace(">", "").replace("|", "")
                        .strip()
                    )
                    held_date = "UNKNOWN_DATE"
                    try:
                        held_date = pd.to_datetime(meta.get("start", ""), format="mixed").strftime("%d-%m-%Y")
                    except Exception:
                        pass
                    mode = (
                        "ATT & LATE" if self.do_att.get() and self.do_late.get() else
                        "ATT"        if self.do_att.get() else
                        "LATE"       if self.do_late.get() else "REPORT"
                    )
                    topic = f"[{held_date}] {topic} {mode}"
                except Exception:
                    topic = "REPORT"

                out = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=f"{topic}.xlsx",
                    title=f"Save report for {topic}"
                )
                if not out:
                    return False   # user cancelled — keep ui2 open
                save_jobs.append((cfg, out))

            self.pending_save_jobs = save_jobs
            threading.Thread(
                target=lambda: self._run_multi_reports(configs, save_jobs),
                daemon=True
            ).start()
            return True   # all saves confirmed → ui2 can close

        MultiCSVConfirm(self, reports, generate_callback,
                        do_att=self.do_att.get(),
                        do_late=self.do_late.get(),
                        do_csv=self.do_csv.get())

    def _run_multi_reports(self, configs, save_jobs):
        success, failed, last_out = 0, 0, None
        for cfg, out in save_jobs:
            try:
                do_att  = cfg.get("do_att",  self.do_att.get())  or cfg.get("do_csv", self.do_csv.get())
                do_late = cfg.get("do_late", self.do_late.get()) or cfg.get("do_csv", self.do_csv.get())
                process(
                    group,
                    do_att,
                    do_late,
                    late_time,
                    late_ampm,
                    total_att,
                    absent_limit,
                    out,
                    selected_data=selected_data,
                    batch_name=self.batch.get(),
                    issued_date=datetime.datetime.now().strftime("%d/%m/%Y")
                )

                # ---------- CSV EXPORT ----------
                if cfg.get("do_csv", self.do_csv.get()):
                    try:
                        final_data = [
                            r for r in (extract_att_data(out) or []) + (extract_late_data(out) or [])
                            if r.get("Roll Number") or r.get("Name")
                        ]
                        if final_data:
                            base = (
                                os.path.splitext(os.path.basename(out))[0]
                                .replace("ATT & LATE", "").replace("ATT", "")
                                .replace("LATE", "").replace("REPORT", "").strip()
                            )
                            export_to_csv(final_data, os.path.dirname(out), f"{base} CSV")
                    except Exception as e:
                        print("CSV export failed:", e)

                last_out = out
                success += 1
            except Exception as e:
                print("Process failed:", e)
                failed += 1

        self.after(0, lambda: messagebox.showinfo(
            "Completed", f"{success} report(s) generated\n{failed} failed"
        ))
        try:
            if last_out:
                self.open_file(last_out if len(configs) == 1 else os.path.dirname(last_out))
        except Exception:
            pass
        
    def run(self):
        if not self.csv_files:
            messagebox.showerror("Error", "Select Zoom generated CSV")
            return

        try:
            total_att    = int(self.att_entry.get())
            absent_limit = int(self.abs_entry.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Attendance and Absent values must be integers")
            return

        groups = group_csv_files(self.csv_files)

        # ---- multi-group → open confirmation window ----
        if len(groups) > 1:
            self.open_multi_confirm(groups, [])
            return

        # ---- single group ----
        group = groups[0]
        try:
            meta  = extract_meeting_metadata(group[0])
            topic = (
                str(meta.get("topic", "REPORT"))
                .replace("MENTORBEE", "").replace("EDUVERSE", "")
                .replace("/", "-").replace("\\", "-").replace(":", "-").strip()
            )
            held_date = "UNKNOWN_DATE"
            try:
                held_date = pd.to_datetime(meta.get("start", ""), format="mixed").strftime("%d-%m-%Y")
            except Exception:
                pass
            mode = (
                "ATT & LATE" if self.do_att.get() and self.do_late.get() else
                "ATT"        if self.do_att.get() else
                "LATE"       if self.do_late.get() else "REPORT"
            )
            filename = f"[{held_date}] {topic} {mode}"
        except Exception:
            filename = "REPORT"

        out = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"{filename}.xlsx",
            title="Save Report"
        )
        if not out:
            return

        # Compute late cutoff — use cluster logic for AUTOMATIC, else UI value
        if self.batch.get() == "AUTOMATIC":
            late_time, late_ampm = self._compute_late_cutoff(group)
        else:
            late_time = self.late_time.get()
            late_ampm = self.late_ampm.get()

        do_att  = self.do_att.get()  or self.do_csv.get()
        do_late = self.do_late.get() or self.do_csv.get()

        selected = []
        if self.do_csv.get():
            selected.append("CSV")

        if self.do_att.get():
            selected.append("ATT")

        if self.do_late.get():
            selected.append("LATE")

        selected_data = ", ".join(selected)

        def task():
            try:
                process(
                    group,
                    do_att,
                    do_late,
                    late_time,
                    late_ampm,
                    total_att,
                    absent_limit,
                    out,
                    selected_data=selected_data,
                    batch_name=self.batch.get(),
                    issued_date=datetime.datetime.now().strftime("%d/%m/%Y")
                )

                # ---------- CSV EXPORT ----------
                csv_out = None
                if self.do_csv.get():
                    try:
                        final_data = [
                            r for r in (extract_att_data(out) or []) + (extract_late_data(out) or [])
                            if r.get("Roll Number") or r.get("Name")
                        ]
                        if final_data:
                            base = (
                                os.path.splitext(os.path.basename(out))[0]
                                .replace("ATT & LATE", "").replace("ATT", "")
                                .replace("LATE", "").replace("REPORT", "").strip()
                            )
                            csv_out = export_to_csv(final_data, os.path.dirname(out), f"{base} CSV")
                    except Exception as e:
                        print("CSV export failed:", e)

                # ---------- RATIO CHECK ----------
                try:
                    total_people = len(pd.read_excel(out, sheet_name="ATT"))
                    issue_count  = (
                        len(extract_att_data(out) or []) * self.do_att.get() +
                        len(extract_late_data(out) or []) * self.do_late.get()
                    )
                    if total_people > 0 and issue_count / total_people > 0.55:
                        self.after(0, lambda: messagebox.showwarning(
                            "Warning",
                            "Something wrong on the values you have entered please do check."
                        ))
                except Exception as e:
                    print("Ratio check failed:", e)

                self.after(0, lambda: messagebox.showinfo("Completed", "Report generated"))
                self.after(0, lambda: self.open_file(out))
                if csv_out:
                    self.after(0, lambda p=csv_out: self.open_file(p))

            except Exception as e:
                err = str(e)
                self.after(0, lambda err=err: messagebox.showerror("Error", err))

        threading.Thread(target=task, daemon=True).start()