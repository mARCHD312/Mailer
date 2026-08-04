import customtkinter as ctk
import tkinter.filedialog as fd
import json
import os
from mail_engine import MailEngine
import database

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

CONFIG_FILE = r"E:\POSAO\Skripta 2.0 - mejl\config.json"

class MailApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Zoho Mailer 2.0 - Professional")
        self.geometry("900x700")

        self.engine = None
        self.load_config()

        # Layout konfiguracija
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Leva kolona (Podesavanja)
        self.frame_left = ctk.CTkFrame(self)
        self.frame_left.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Desna kolona (Sadrzaj i Logovi)
        self.frame_right = ctk.CTkFrame(self)
        self.frame_right.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        self.setup_left_panel()
        self.setup_right_panel()

    def load_config(self):
        self.config = {
            "email": "",
            "password": "",
            "excel_path": "",
            "limit": "50",
            "subject": "Predlog za saradnju",
            "body": "Poštovani,\n\nPišemo ispred kompanije...\nZa {company_name}.\n\nPozdrav",
            "start_hour": "8",
            "end_hour": "17"
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.config.update(json.load(f))
            except:
                pass

    def save_config(self):
        self.config["email"] = self.entry_email.get()
        self.config["password"] = self.entry_pass.get()
        self.config["excel_path"] = self.entry_excel.get()
        self.config["limit"] = self.entry_limit.get()
        self.config["subject"] = self.entry_subject.get()
        self.config["body"] = self.text_body.get("1.0", "end-1c")
        self.config["start_hour"] = self.entry_start.get()
        self.config["end_hour"] = self.entry_end.get()
        
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=4)

    def setup_left_panel(self):
        lbl_title = ctk.CTkLabel(self.frame_left, text="1. SMTP Konekcija (Zoho)", font=("Arial", 16, "bold"))
        lbl_title.pack(pady=10)

        self.entry_email = ctk.CTkEntry(self.frame_left, placeholder_text="Tvoj Zoho Email")
        self.entry_email.insert(0, self.config["email"])
        self.entry_email.pack(fill="x", padx=20, pady=5)

        self.entry_pass = ctk.CTkEntry(self.frame_left, placeholder_text="Lozinka / App Password", show="*")
        self.entry_pass.insert(0, self.config["password"])
        self.entry_pass.pack(fill="x", padx=20, pady=5)

        lbl_kampanja = ctk.CTkLabel(self.frame_left, text="2. Podešavanje Kampanje", font=("Arial", 16, "bold"))
        lbl_kampanja.pack(pady=(20, 10))

        frame_excel = ctk.CTkFrame(self.frame_left, fg_color="transparent")
        frame_excel.pack(fill="x", padx=20, pady=5)
        
        self.entry_excel = ctk.CTkEntry(frame_excel, placeholder_text="Putanja do Excel baze")
        self.entry_excel.insert(0, self.config["excel_path"])
        self.entry_excel.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        btn_browse = ctk.CTkButton(frame_excel, text="Odaberi fajl", width=90, command=self.browse_excel)
        btn_browse.pack(side="right")

        frame_limits = ctk.CTkFrame(self.frame_left, fg_color="transparent")
        frame_limits.pack(fill="x", padx=20, pady=5)
        
        ctk.CTkLabel(frame_limits, text="Limit/dan:").pack(side="left")
        self.entry_limit = ctk.CTkEntry(frame_limits, width=50)
        self.entry_limit.insert(0, self.config["limit"])
        self.entry_limit.pack(side="left", padx=(5, 15))

        ctk.CTkLabel(frame_limits, text="Od h:").pack(side="left")
        self.entry_start = ctk.CTkEntry(frame_limits, width=40)
        self.entry_start.insert(0, self.config["start_hour"])
        self.entry_start.pack(side="left", padx=5)

        ctk.CTkLabel(frame_limits, text="Do h:").pack(side="left")
        self.entry_end = ctk.CTkEntry(frame_limits, width=40)
        self.entry_end.insert(0, self.config["end_hour"])
        self.entry_end.pack(side="left", padx=5)

        lbl_db = ctk.CTkLabel(self.frame_left, text="3. Zaštita od duplog slanja", font=("Arial", 16, "bold"))
        lbl_db.pack(pady=(20, 10))
        
        total = database.get_total_sent()
        self.lbl_db_info = ctk.CTkLabel(self.frame_left, text=f"Ukupno jedinstvenih adresa u bazi: {total}", text_color="#00FF00")
        self.lbl_db_info.pack()

    def setup_right_panel(self):
        lbl_title = ctk.CTkLabel(self.frame_right, text="Tekst Emaila", font=("Arial", 16, "bold"))
        lbl_title.pack(pady=10)

        self.entry_subject = ctk.CTkEntry(self.frame_right, placeholder_text="Naslov mejla")
        self.entry_subject.insert(0, self.config["subject"])
        self.entry_subject.pack(fill="x", padx=20, pady=5)

        ctk.CTkLabel(self.frame_right, text="Telo poruke (koristi {company_name} za ime firme):", font=("Arial", 12)).pack(anchor="w", padx=20)
        self.text_body = ctk.CTkTextbox(self.frame_right, height=200)
        self.text_body.insert("1.0", self.config["body"])
        self.text_body.pack(fill="x", padx=20, pady=5)

        # Kontrole
        frame_controls = ctk.CTkFrame(self.frame_right, fg_color="transparent")
        frame_controls.pack(pady=20)

        self.btn_start = ctk.CTkButton(frame_controls, text="▶ START", fg_color="green", hover_color="darkgreen", font=("Arial", 14, "bold"), command=self.start_campaign)
        self.btn_start.pack(side="left", padx=10)

        self.btn_pause = ctk.CTkButton(frame_controls, text="⏸ PAUZA", fg_color="orange", hover_color="darkorange", font=("Arial", 14, "bold"), state="disabled", command=self.pause_campaign)
        self.btn_pause.pack(side="left", padx=10)

        self.btn_stop = ctk.CTkButton(frame_controls, text="⏹ STOP", fg_color="red", hover_color="darkred", font=("Arial", 14, "bold"), state="disabled", command=self.stop_campaign)
        self.btn_stop.pack(side="left", padx=10)

        # Progress i log
        self.progress = ctk.CTkProgressBar(self.frame_right)
        self.progress.pack(fill="x", padx=20, pady=10)
        self.progress.set(0)

        self.text_log = ctk.CTkTextbox(self.frame_right, height=180, state="disabled", fg_color="black", text_color="#00FF00", font=("Consolas", 12))
        self.text_log.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    def browse_excel(self):
        filepaths = fd.askopenfilenames(filetypes=[("Excel fajlovi", "*.xlsx")])
        if filepaths:
            self.entry_excel.delete(0, "end")
            # Spajamo ih znakom | kako bi UI mogao to da prikaze (ili sacuva)
            self.entry_excel.insert(0, "|".join(filepaths))

    def log(self, msg):
        self.text_log.configure(state="normal")
        self.text_log.insert("end", msg + "\n")
        self.text_log.see("end")
        self.text_log.configure(state="disabled")

    def update_progress(self, val):
        self.progress.set(val)

    def on_engine_finish(self):
        self.btn_start.configure(state="normal")
        self.btn_pause.configure(state="disabled", text="⏸ PAUZA")
        self.btn_stop.configure(state="disabled")
        
        # Update db counter
        total = database.get_total_sent()
        self.lbl_db_info.configure(text=f"Ukupno jedinstvenih adresa u bazi: {total}")

    def start_campaign(self):
        self.save_config()
        self.text_log.configure(state="normal")
        self.text_log.delete("1.0", "end")
        self.text_log.configure(state="disabled")
        
        self.progress.set(0)

        try:
            limit = int(self.entry_limit.get())
            sh = int(self.entry_start.get())
            eh = int(self.entry_end.get())
        except:
            self.log("❌ Greška: Limit i radni sati moraju biti brojevi!")
            return

        excel_val = self.entry_excel.get()
        if not excel_val:
            self.log("❌ Greška: Morate odabrati bar jedan Excel fajl!")
            return
            
        paths_list = [p.strip() for p in excel_val.split('|') if p.strip()]

        self.engine = MailEngine(
            excel_paths=paths_list,
            smtp_server="smtp.zoho.eu", # Mozes izvuci u UI kasnije ako menjas Zoho u nesto drugo
            port=465,
            email=self.entry_email.get(),
            password=self.entry_pass.get(),
            limit=limit,
            start_hour=sh,
            end_hour=eh,
            log_callback=self.log,
            progress_callback=self.update_progress,
            finish_callback=self.on_engine_finish
        )
        
        self.btn_start.configure(state="disabled")
        self.btn_pause.configure(state="normal")
        self.btn_stop.configure(state="normal")

        self.engine.start(
            subject=self.entry_subject.get(),
            body_template=self.text_body.get("1.0", "end-1c"),
            col_email="Email",
            col_name="Naziv"
        )

    def pause_campaign(self):
        if self.engine:
            self.engine.pause()
            if self.engine.is_paused:
                self.btn_pause.configure(text="▶ NASTAVI")
            else:
                self.btn_pause.configure(text="⏸ PAUZA")

    def stop_campaign(self):
        if self.engine:
            self.engine.stop()

if __name__ == "__main__":
    app = MailApp()
    app.mainloop()
