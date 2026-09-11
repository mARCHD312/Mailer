import os
import json
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
from mail_engine import MailEngine
import threading

app = Flask(__name__, template_folder='web_templates')
app.secret_key = 'tajna_sifra_za_sesije_123'

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

class AppState:
    engine = None
    logs = []
    progress = 0

def log_callback(msg):
    AppState.logs.append(msg)
    if len(AppState.logs) > 150:
        AppState.logs.pop(0)

def progress_callback(val):
    AppState.progress = val

def finish_callback():
    pass

@app.route('/')
def index():
    config = {
        "email": "",
        "password": "",
        "limit": "50",
        "start_hour": "8",
        "end_hour": "17",
        "language": "Engleski"
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config.update(json.load(f))
        except:
            pass
    return render_template('index.html', config=config)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "Nema fajla u zahtevu"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "Nije izabran fajl"}), 400
    if file and file.filename.endswith('.xlsx'):
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        return jsonify({"success": True, "filepath": filepath})
    return jsonify({"error": "Samo .xlsx fajlovi su dozvoljeni"}), 400

@app.route('/start', methods=['POST'])
def start_campaign():
    if AppState.engine and AppState.engine.is_running:
        return jsonify({"error": "Kampanja već u toku"}), 400

    data = request.json
    
    config = {
        "email": data.get("email"),
        "password": data.get("password"),
        "excel_path": "", 
        "limit": data.get("limit"),
        "start_hour": data.get("start_hour"),
        "end_hour": data.get("end_hour"),
        "language": data.get("language")
    }
    
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)

    excel_path = data.get("filepath")
    if not excel_path or not os.path.exists(excel_path):
         return jsonify({"error": "Morate prvo otpremiti Excel fajl!"}), 400

    AppState.logs = []
    AppState.progress = 0

    AppState.engine = MailEngine(
        excel_paths=[excel_path],
        smtp_server="smtp.zoho.eu",
        port=465,
        email=config["email"],
        password=config["password"],
        limit=int(config["limit"]),
        start_hour=int(config["start_hour"]),
        end_hour=int(config["end_hour"]),
        language=config["language"],
        log_callback=log_callback,
        progress_callback=progress_callback,
        finish_callback=finish_callback
    )

    AppState.engine.start(col_email="Email", col_name="Naziv")

    return jsonify({"success": True})

@app.route('/pause', methods=['POST'])
def pause_campaign():
    if AppState.engine:
        AppState.engine.pause()
        return jsonify({"success": True, "is_paused": AppState.engine.is_paused})
    return jsonify({"error": "Nema aktivne kampanje"}), 400

@app.route('/stop', methods=['POST'])
def stop_campaign():
    if AppState.engine:
        AppState.engine.stop()
        return jsonify({"success": True})
    return jsonify({"error": "Nema aktivne kampanje"}), 400

@app.route('/status', methods=['GET'])
def get_status():
    import database
    return jsonify({
        "logs": AppState.logs,
        "progress": AppState.progress,
        "is_running": AppState.engine.is_running if AppState.engine else False,
        "is_paused": AppState.engine.is_paused if AppState.engine else False,
        "total_db": database.get_total_sent()
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
