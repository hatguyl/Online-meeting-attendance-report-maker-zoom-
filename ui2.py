import customtkinter as ctk
from PIL import Image, ImageTk
import os, sys

ctk.set_appearance_mode("dark")

def resource_path(rel):
    try: base = sys._MEIPASS
    except: base = os.path.abspath(".")
    return os.path.join(base, rel)

BG    = "#111111"
CARD  = "#1c1c1c"
FIELD = "#2a2a2a"
TEXT  = "#ffffff"
MUTED = "#bdbdbd"
ACCENT= "#b630f4"

ROW_H = 54   

class MultiCSVConfirm(ctk.CTkToplevel):

    def __init__(self, parent, reports, on_generate, do_att=True, do_late=True, do_csv=True):
        super().__init__(parent)
        self.title("Report tool")
        self.resizable(False, False)
        self.configure(fg_color=BG)
        self.on_generate = on_generate
        self.do_att_global  = do_att
        self.do_late_global = do_late
        self.do_csv_global  = do_csv
        self.vars = []
        self.enabled_vars = []

        # ---- window icon ----

        def _apply_icon(event=None):
            try:
                import numpy as np
                raw  = Image.open(resource_path("logo.png")).convert("RGBA")
                data = np.array(raw)
                data[(data[:,:,0]<40)&(data[:,:,1]<40)&(data[:,:,2]<40), 3] = 0
                icon = ImageTk.PhotoImage(Image.fromarray(data).resize((32,32), Image.LANCZOS))
                self._tk_icon = icon          
                self.iconphoto(False, icon)
            except Exception:
                pass

        _apply_icon()
        self.bind("<Map>", lambda e: (_apply_icon(), self.unbind("<Map>")))
        self.after(350, _apply_icon)

        n     = len(reports)
        win_h = 50 + 30 + min(n, 8) * (ROW_H + 8) + 70 + 30
        self.geometry(f"560x{win_h}")

        # ---- main card ----
        card = ctk.CTkFrame(self, corner_radius=18, fg_color=CARD)
        card.pack(padx=12, pady=10, fill="both", expand=True)

        ctk.CTkLabel(card, text="CONFIRMATION", font=("Segoe UI", 18, "bold"),
                     text_color=TEXT).pack(pady=(10, 6))

        # ---- column headers ----
        hdr = ctk.CTkFrame(card, fg_color="transparent")
        hdr.pack(fill="x", padx=20, pady=(0, 2))
        ctk.CTkLabel(hdr, text="",       width=36                              ).pack(side="left")
        ctk.CTkLabel(hdr, text="Topic",  width=205, anchor="w",  text_color=MUTED, font=("Segoe UI",11,"bold")).pack(side="left")
        ctk.CTkLabel(hdr, text="ATT",    width=55,  anchor="center", text_color=MUTED, font=("Segoe UI",11,"bold")).pack(side="left")
        ctk.CTkLabel(hdr, text="ABSENT", width=68,  anchor="center", text_color=MUTED, font=("Segoe UI",11,"bold")).pack(side="left")
        ctk.CTkLabel(hdr, text="LATE",   width=106, anchor="center", text_color=MUTED, font=("Segoe UI",11,"bold")).pack(side="left")

        # ---- BUTTON first ----
        ctk.CTkButton(
            card, text="GENERATE ALL", height=38, corner_radius=12,
            fg_color=ACCENT, hover_color="#ffd95a", text_color="#000000",
            font=("Segoe UI", 13, "bold"), command=self._generate
        ).pack(side="bottom", padx=28, pady=(0, 14), fill="x")

        # ---- scroll area ----
        scroll = ctk.CTkScrollableFrame(card, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        for report in reports:
            self._add_row(scroll, report)

    def _add_row(self, parent, report):
        enabled  = ctk.BooleanVar(value=True)
        att_var  = ctk.StringVar(value=str(report["attendance"]))
        ab_var   = ctk.StringVar(value=str(report["absent"]))
        late_var = ctk.StringVar(value=report["late_time"])

        row = ctk.CTkFrame(parent, height=46, fg_color="#232323", corner_radius=14)
        row.pack(fill="x", pady=4, padx=6)
        row.pack_propagate(False)

        # checkbox
        ctk.CTkCheckBox(row, text="", variable=enabled, width=22,
                        checkbox_width=18, checkbox_height=18,
                        fg_color=ACCENT, hover_color="#ffd95a", border_color="#666"
                       ).pack(side="left", padx=(10, 2))

        # topic
        topic = report["topic"][:26] + "…" if len(report["topic"]) > 26 else report["topic"]
        ctk.CTkLabel(row, text=topic, width=200, anchor="w",
                     text_color="#e0e0e0", font=("Segoe UI", 11)
                    ).pack(side="left", padx=(4, 0))

        # ATT + ABSENT entries
        for var, w in [(att_var, 46), (ab_var, 60)]:
            ctk.CTkEntry(row, width=w, height=30, textvariable=var,
                         justify="center", fg_color=FIELD,
                         border_color="#555", text_color=TEXT
                        ).pack(side="left", padx=4)

        # late time + AM/PM
        lf = ctk.CTkFrame(row, fg_color="transparent")
        lf.pack(side="left", padx=(4, 10))
        ctk.CTkEntry(lf, width=72, height=30, textvariable=late_var,
                     justify="center", fg_color="#2c2c2c", border_color="#4a4a4a"
                    ).pack(side="left")
        ctk.CTkLabel(lf, text=report["late_ampm"], width=30,
                     text_color=MUTED, font=("Segoe UI", 10, "bold")
                    ).pack(side="left", padx=(4, 0))

        self.enabled_vars.append(enabled)
        self.vars.append({
            "group":      report["group"],
            "topic":      report["topic"],
            "attendance": att_var,
            "absent":     ab_var,
            "late":       late_var,
            "late_ampm":  report["late_ampm"],
            "do_att":     self.do_att_global,
            "do_late":    self.do_late_global,
            "do_csv":     self.do_csv_global,
        })

    def _generate(self):
        configs = []
        for enabled, item in zip(self.enabled_vars, self.vars):
            if not enabled.get():
                continue
            configs.append({
                "group":      item["group"],
                "topic":      item["topic"],
                "attendance": int(item["attendance"].get()),
                "absent":     int(item["absent"].get()),
                "late_time":  item["late"].get().strip(),
                "late_ampm":  item["late_ampm"],
                "do_att":     item["do_att"],
                "do_late":    item["do_late"],
                "do_csv":     item["do_csv"],
            })
        if configs:
            result = self.on_generate(configs)
            if result:
                self.destroy()