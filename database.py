import sqlite3
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sent_emails.db")

def init_db():
    """Kreira bazu i tabelu ako ne postoje."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sent_emails (
            email TEXT PRIMARY KEY,
            sent_date TEXT
        )
    ''')
    conn.commit()
    conn.close()

def is_email_sent(email):
    """Proverava da li je mejl već u bazi (da li mu je već slato)."""
    email = email.lower().strip()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('SELECT 1 FROM sent_emails WHERE email = ?', (email,))
    result = cursor.fetchone()
    conn.close()
    return result is not None

def add_sent_email(email):
    """Dodaje mejl u bazu poslatih."""
    email = email.lower().strip()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        cursor.execute('INSERT INTO sent_emails (email, sent_date) VALUES (?, ?)', (email, now_str))
        conn.commit()
    except sqlite3.IntegrityError:
        pass # Vec postoji
    finally:
        conn.close()

def get_total_sent():
    """Vraća ukupan broj mejlova u bazi."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) FROM sent_emails')
    count = cursor.fetchone()[0]
    conn.close()
    return count

# Kada se fajl prvi put uveze, inicijalizuj bazu
init_db()
