import pandas as pd
import smtplib
from email.message import EmailMessage
import time
import random
from datetime import datetime
import os
import database
import threading

class MailEngine:
    def __init__(self, excel_paths, smtp_server, port, email, password, limit, start_hour, end_hour, log_callback, progress_callback, finish_callback):
        self.excel_paths = excel_paths
        self.smtp_server = smtp_server
        self.port = port
        self.email = email
        self.password = password
        self.limit = limit
        self.start_hour = start_hour
        self.end_hour = end_hour
        self.log = log_callback
        self.update_progress = progress_callback
        self.on_finish = finish_callback
        
        self.is_running = False
        self.is_paused = False
        self.thread = None

    def start(self, subject, body_template, col_email, col_name):
        if self.is_running:
            return
        self.is_running = True
        self.is_paused = False
        self.thread = threading.Thread(target=self._run, args=(subject, body_template, col_email, col_name))
        self.thread.daemon = True
        self.thread.start()

    def pause(self):
        if self.is_running:
            self.is_paused = not self.is_paused
            if self.is_paused:
                self.log("[PAUZA] Kampanja je pauzirana.")
            else:
                self.log("[NASTAVAK] Kampanja se nastavlja...")

    def stop(self):
        if self.is_running:
            self.is_running = False
            self.log("[STOP] Zaustavljam kampanju, sačekaj...")

    def _connect(self):
        server = smtplib.SMTP_SSL(self.smtp_server, self.port)
        server.login(self.email, self.password)
        return server

    def _run(self, subject, body_template, col_email, col_name):
        try:
            server = self._connect()
            self.log("✅ Uspesna konekcija sa Zoho serverom.")
        except Exception as e:
            self.log(f"❌ Greska pri konekciji: {e}")
            self.is_running = False
            self.on_finish()
            return

        sent_count = 0
        total_seconds = (self.end_hour - self.start_hour) * 3600
        average_sleep = total_seconds / max(1, self.limit)

        for excel_path in self.excel_paths:
            if not self.is_running:
                break
            if sent_count >= self.limit:
                break
                
            self.log(f"\n--- Otvaram fajl: {os.path.basename(excel_path)} ---")
            if not os.path.exists(excel_path):
                self.log(f"GRESKA: Fajl nije pronađen: {excel_path}")
                continue
                
            try:
                df = pd.read_excel(excel_path)
            except Exception as e:
                self.log(f"GRESKA pri učitavanju Excela: {e}")
                continue

            if 'Status' not in df.columns:
                df['Status'] = ''
            if 'Datum Slanja' not in df.columns:
                df['Datum Slanja'] = ''
                
            df['Status'] = df['Status'].astype(str)
            df['Datum Slanja'] = df['Datum Slanja'].astype(str)
            
            unsent_indices = []
            for index, row in df.iterrows():
                email_val = str(row.get(col_email, '')).strip()
                status = str(row.get('Status', '')).strip()
                
                if status != 'Poslato' and email_val and email_val.lower() != 'nan':
                    if not database.is_email_sent(email_val):
                        unsent_indices.append(index)
                    else:
                        df.at[index, 'Status'] = 'Poslato (Ranije)'
                        
            self.log(f"Pronadjeno {len(unsent_indices)} novih adresa u ovom fajlu.")
            
            if len(unsent_indices) == 0:
                self.log("Sve adrese u ovom fajlu su već obrađene.")
                try:
                    df.to_excel(excel_path, index=False)
                except:
                    pass
                continue

            to_send = unsent_indices
            
            for index in to_send:
                if sent_count >= self.limit:
                    self.log(f"\n[INFO] Dostignut dnevni limit od {self.limit} mejlova.")
                    break
                    
                while self.is_paused and self.is_running:
                    time.sleep(1)
                    
                if not self.is_running:
                    break
                    
                row = df.loc[index]
                email_addr = str(row[col_email]).strip()
                company_name = str(row[col_name]).strip()
                if company_name == 'nan' or not company_name:
                    company_name = "kolege"

                hour = datetime.now().hour
                if not (self.start_hour <= hour < self.end_hour):
                    self.log(f"Trenutno vreme ({hour}h) je van radnog vremena ({self.start_hour}-{self.end_hour}h).")
                    self.log("Zaustavljam kampanju za danas.")
                    self.is_running = False
                    break

                self.log(f"[{sent_count+1}/{self.limit}] Saljem za: {company_name} ({email_addr})")
                
                msg = EmailMessage()
                msg['Subject'] = subject
                msg['From'] = self.email
                msg['To'] = email_addr
                
                body = body_template.replace("{company_name}", company_name)
                msg.set_content(body)
                
                try:
                    server.send_message(msg)
                    self.log(f"✅ Uspesno poslato.")
                    df.at[index, 'Status'] = 'Poslato'
                    df.at[index, 'Datum Slanja'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    database.add_sent_email(email_addr)
                    sent_count += 1
                except Exception as e:
                    self.log(f"❌ Greska pri slanju: {e}")
                    df.at[index, 'Status'] = 'Greška'
                    
                try:
                    df.to_excel(excel_path, index=False)
                except Exception as e:
                    self.log(f"⚠️ Upozorenje: Excel fajl je otvoren! Zatvori ga da bih mogao da sacuvam statuse.")

                self.update_progress(sent_count / self.limit)

                if not self.is_running or sent_count >= self.limit:
                    break
                    
                # Pause between emails
                sleep_time = random.uniform(average_sleep * 0.8, average_sleep * 1.2)
                mins, secs = divmod(sleep_time, 60)
                self.log(f"⏳ Cekam {int(mins)} min {int(secs)} sek do sledeceg...")
                
                slept = 0
                while slept < sleep_time and self.is_running:
                    time.sleep(1)
                    if not self.is_paused:
                        slept += 1
                        
                if sent_count % 3 == 0 and self.is_running:
                    try:
                        server.quit()
                        server = self._connect()
                    except:
                        pass
                        
        try:
            server.quit()
        except:
            pass
            
        self.log(f"\n🎯 Kampanja završena! Ukupno poslato u ovoj sesiji: {sent_count}")
        self.is_running = False
        self.on_finish()
