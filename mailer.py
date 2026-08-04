import pandas as pd
import smtplib
from email.message import EmailMessage
import time
import random
from datetime import datetime
import os

# ================= KONEKCIJA =================
SMTP_SERVER = "smtp.zoho.eu"  # Promeni u smtp.zoho.com ako ti je tamo registrovan nalog
SMTP_PORT = 465 # SSL port
EMAIL_ADDRESS = "info@3dmarch.org"
# PREPORUKA: Umesto glavne lozinke, kreiraj "App Password" na Zoho nalogu zbog sigurnosti
EMAIL_PASSWORD = "Test135!+" 

# ================= PODEŠAVANJA =================
EXCEL_FILE = r"E:\POSAO\Skripta 2.0 - mejl\TEST\Arhitektura_i_projektovanje.xlsx" # Promeni na ime tvog fajla
KOLONA_EMAIL = "Email" # Kako ti se tačno zove kolona sa mejlovima u Excelu
KOLONA_FIRMA = "Naziv" # Kako ti se tačno zove kolona sa imenom firme
EMAILS_PER_DAY = 50
START_HOUR = 8
END_HOUR = 17

def connect_to_smtp():
    """Konektovanje na Zoho server"""
    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    return server

def send_email(server, to_email, company_name):
    msg = EmailMessage()
    msg['Subject'] = "Predlog za saradnju" # Promeni naslov mejla
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email
    
    # Personalizovani tekst mejla - prilagodi ga sebi
    body = f"""Poštovani,

Pišemo ispred kompanije [Tvoje Ime] za kompaniju {company_name}.
Imamo odličan predlog za vas...

Srdačan pozdrav,
Tvoj Tim
"""
    msg.set_content(body)
    
    try:
        server.send_message(msg)
        return True
    except Exception as e:
        print(f"[GRESKA] Greska pri slanju na {to_email}: {e}")
        return False

def main():
    # 1. Provera fajla
    if not os.path.exists(EXCEL_FILE):
        print(f"[GRESKA] Fajl '{EXCEL_FILE}' ne postoji. Proveri putanju i ime fajla.")
        return
        
    print(f"Ucitavam bazu: {EXCEL_FILE}")
    df = pd.read_excel(EXCEL_FILE)
    
    # 2. Pravimo 'Status' i 'Datum Slanja' kolone ako ne postoje
    if 'Status' not in df.columns:
        df['Status'] = ''
    if 'Datum Slanja' not in df.columns:
        df['Datum Slanja'] = ''
        
    # Osiguravamo da kolone budu tekstualne (string) da bi mogao da se upiše tekst
    df['Status'] = df['Status'].astype(str)
    df['Datum Slanja'] = df['Datum Slanja'].astype(str)
        
    # 3. Tražimo one koji nemaju status "Poslato" i koji uopšte imaju mejl
    unsent = df[(df['Status'] != 'Poslato') & (df[KOLONA_EMAIL].notna())]
    
    if unsent.empty:
        print("[INFO] Svi mejlovi iz baze su vec poslati!")
        return
        
    to_send = unsent.head(EMAILS_PER_DAY)
    print(f"Pronadjeno {len(unsent)} neposlatih firmi. Danas saljemo na {len(to_send)} adresa.")
    
    # 4. Računanje pauze između 2 mejla (kako bi popunili vreme od 8h do 17h)
    total_seconds = (END_HOUR - START_HOUR) * 3600
    average_sleep = total_seconds / EMAILS_PER_DAY
    
    try:
        server = connect_to_smtp()
        print("[OK] Uspostavljena konekcija sa Zoho serverom.")
    except Exception as e:
        print(f"[GRESKA] Ne mogu da se povezem na Zoho: {e}")
        return

    sent_count = 0
    
    for index, row in to_send.iterrows():
        # Proveravamo da li smo još uvek u dozvoljenom radnom vremenu
        current_hour = datetime.now().hour
        if not (START_HOUR <= current_hour < END_HOUR):
            print(f"\nTrenutno vreme ({current_hour}h) je van radnog vremena ({START_HOUR}-{END_HOUR}h).")
            print("Prekidam slanje. Mozes pokrenuti skriptu sutra!")
            break
            
        # Uzimamo podatke
        company_name = str(row[KOLONA_FIRMA]).strip()
        if company_name == 'nan' or not company_name:
            company_name = "kolege" # Ako nema imena firme
            
        email_addr = str(row[KOLONA_EMAIL]).strip()
        
        print(f"[{sent_count+1}/{EMAILS_PER_DAY}] Saljem mejl za: {company_name} ({email_addr})...")
        
        success = send_email(server, email_addr, company_name)
        
        if success:
            print(f"[OK] Uspesno.")
            # Upisujemo status da znamo za sledeći put
            df.at[index, 'Status'] = 'Poslato'
            df.at[index, 'Datum Slanja'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            # Čuvamo excel odmah, ako pukne program da se ne izgubi
            df.to_excel(EXCEL_FILE, index=False)
            sent_count += 1
            
            # Ako nije poslednji, čekamo
            if sent_count < EMAILS_PER_DAY:
                # Malo variramo vreme (+/- 20%) da izgleda kao da čovek šalje
                sleep_time = random.uniform(average_sleep * 0.8, average_sleep * 1.2)
                mins, secs = divmod(sleep_time, 60)
                print(f"[CEKAM] {int(mins)} min {int(secs)} sekundi do sledeceg...\n")
                time.sleep(sleep_time)
                
                # Relogin povremeno jer SMTP serveri diskonektuju ako se dugo čeka
                if sent_count % 3 == 0:
                    try:
                        server.quit()
                        server = connect_to_smtp()
                    except:
                        pass
        else:
            df.at[index, 'Status'] = 'Greška'
            df.to_excel(EXCEL_FILE, index=False)
            
    try:
        server.quit()
    except:
        pass
        
    print(f"\n[ZAVRSENO] Za danas gotovo! Ukupno poslato: {sent_count} mejlova.")

if __name__ == "__main__":
    main()
