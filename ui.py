import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import os, sys, threading
import subprocess
import pandas as pd 
from core import process, extract_att_data, extract_late_data, export_to_csv


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
        self.MUTED = "#ffffff"
        self.ACCENT = "#9c30f4"

        # +----------------+
        # | STATE VARIABLES|
        # +----------------+
        self.csv_path = ""
        self.do_att = ctk.BooleanVar(value=True)
        self.do_late = ctk.BooleanVar(value=True)
        self.do_csv = ctk.BooleanVar(value=True)

        self.batch = ctk.StringVar(value="MRNG (6 to 7:30)")
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
        logo_img = Image.open(resource_path("logo.png")).resize((160, 55), Image.LANCZOS)
        self.logo = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(160, 55))

        icon_img = Image.open(resource_path("logo.png")).resize((64, 64), Image.LANCZOS)
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
            values=["MRNG (6 to 7:30)",
                    "MRNG (6 to 8)", 
                    "EVNG (7 to 9:30)",
                    "EVNG (7:30 to 9:30)",
                    "REVISE & EXAM",
                    "OTHERS"],
            variable=self.batch,
            command=self.apply_batch,
            fg_color=self.FIELD,
            button_color=self.ACCENT,
            button_hover_color="#ffd95a"
        )
        self.batch_menu.pack()

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

        self.apply_batch("MRNG")
        self.update_input_states()

        # +----------------+
        # | VERSION LABEL  |
        # +----------------+
        self.version_label = ctk.CTkLabel(
            self,
            text="v2.0.0",
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
            df = pd.read_csv(csv_path)

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


    # +----------------+
    # | FILE HANDLING  |
    # +----------------+
    def browse(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path:
            self.csv_path = path
            self.drop_label.configure(text=f"📄\n{os.path.basename(path)}")
            self.update_issued_date(path)

    def on_drop(self, event):
        files = self.tk.splitlist(event.data)

        for path in files:
            path = path.strip()

            if path.lower().endswith(".csv") and os.path.exists(path):
                self.csv_path = path
                self.drop_label.configure(
                    text=f"📄\n{os.path.basename(path)}"
                )
                self.update_issued_date(path)
                return


        messagebox.showerror(
            "Invalid file",
            "Please drop a Zoom-generated CSV file only."
        )

    # +----------------+
    # | BATCH LOGIC    |
    # +----------------+
    def apply_batch(self, _):
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
    def run(self):
        if not self.csv_path:
            messagebox.showerror("Error", "Select Zoom generated CSV")
            return

        try:
            total_att = int(self.att_entry.get())
            absent_limit = int(self.abs_entry.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Attendance and Absent values must be integers")
            return

        out = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if not out:
            return
        
        do_att = self.do_att.get() or self.do_csv.get()
        do_late = self.do_late.get() or self.do_csv.get()

        def task():
            try:
                process(
                    self.csv_path,
                    do_att,
                    do_late,
                    self.late_time.get(),
                    self.late_ampm.get(),
                    total_att,
                    absent_limit,
                    out
                )
                # ---------- CSV EXPORT ----------
                if self.do_csv.get():
                    try:
                        att_data = extract_att_data(out) or []
                        late_data = extract_late_data(out) or []

                        final_data = att_data + late_data

                        final_data = [
                            row for row in final_data
                            if row.get("Roll Number") or row.get("Name")
                        ]

                        if final_data:
                            folder = os.path.dirname(out)
                            filename = os.path.splitext(os.path.basename(out))[0] + "_C"

                            export_to_csv(final_data, folder, filename)

                    except Exception as e:
                        print("CSV export failed:", e)
                def done():
                    self.open_file(out)

                    csv_path = os.path.join(
                        os.path.dirname(out),
                        os.path.splitext(os.path.basename(out))[0] + "_C.csv"
                    )

                    if os.path.exists(csv_path):
                        self.open_file(csv_path)

                    messagebox.showinfo("Success", "Report generated successfully")

                self.after(0, done)
                # ---------- ISSUE RATIO CHECK ----------
                try:
                    total_people = len(pd.read_excel(out, sheet_name="ATT"))

                    issue_count = 0
                    if self.do_att.get():
                        issue_count += len(extract_att_data(out) or [])
                    if self.do_late.get():
                        issue_count += len(extract_late_data(out) or [])

                    if total_people > 0:
                        ratio = issue_count / total_people

                        if ratio > 0.55:
                            self.after(
                                0,
                                lambda: messagebox.showwarning(
                                    "Warning",
                                    "Something wrong on the values you have entered please do check."
                                )
                            )
                except Exception as e:
                    print("Ratio check failed:", e)
            except Exception as e:
                err = str(e)
                self.after(
                    0,
                    lambda err=err: messagebox.showerror("Error", err)
                )


        threading.Thread(target=task, daemon=True).start()

