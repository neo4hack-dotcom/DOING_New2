
import os
import json
import time
import smtplib
import ssl
import threading
from datetime import datetime, timezone
from email.message import EmailMessage
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

app = Flask(__name__, static_folder='dist')
CORS(app)  # Autoriser les requêtes Cross-Origin (nécessaire pour le développement)

PORT = 3001
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVER_CONFIG_FILE = os.path.join(BASE_DIR, 'server-config.json')
DEFAULT_DB_FILE = os.path.join(BASE_DIR, 'db.json')
SCHEDULER_INTERVAL_SECONDS = 20
DB_LOCK = threading.Lock()

# --- FONCTIONS UTILITAIRES ---

def get_db_path():
    """Récupère le chemin de la base de données depuis la config ou utilise le chemin par défaut."""
    try:
        if os.path.exists(SERVER_CONFIG_FILE):
            with open(SERVER_CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                if config.get('dbPath'):
                    return config['dbPath']
    except Exception as e:
        print(f"Error reading config: {e}")
    return DEFAULT_DB_FILE


def default_smtp_config():
    return {
        "host": "",
        "port": 587,
        "user": "",
        "password": "",
        "security": "starttls",
        "clientHostname": ""
    }


def normalize_smtp_config(raw):
    raw = raw or {}
    try:
        port = int(raw.get('port', 587))
    except Exception:
        port = 587
    if port <= 0:
        port = 587
    security = raw.get('security')
    if security not in ('none', 'ssl', 'starttls'):
        security = 'starttls'
    return {
        "host": str(raw.get('host', '')).strip(),
        "port": port,
        "user": str(raw.get('user', '')).strip(),
        "password": str(raw.get('password', '')),
        "security": security,
        "clientHostname": str(raw.get('clientHostname', '')).strip()
    }


def parse_iso_datetime(value):
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.endswith('Z'):
        text = text[:-1] + '+00:00'
    try:
        dt = datetime.fromisoformat(text)
    except Exception:
        return None
    if dt.tzinfo is None:
        local_tz = datetime.now().astimezone().tzinfo
        dt = dt.replace(tzinfo=local_tz)
    return dt.astimezone(timezone.utc)


def recipients_from_payload(payload):
    def normalize(value):
        if isinstance(value, list):
            return [str(v).strip() for v in value if str(v).strip()]
        if isinstance(value, str):
            parts = value.replace(';', ',').split(',')
            return [p.strip() for p in parts if p.strip()]
        return []

    recipients = {
        "to": normalize(payload.get('to')),
        "cc": normalize(payload.get('cc')),
        "bcc": normalize(payload.get('bcc')),
    }
    if not recipients["to"] and not recipients["cc"] and not recipients["bcc"]:
        raise ValueError("At least one recipient is required (to, cc, or bcc).")
    return recipients


def recipients_from_job(job):
    return recipients_from_payload({
        "to": (job.get('recipients') or {}).get('to', []),
        "cc": (job.get('recipients') or {}).get('cc', []),
        "bcc": (job.get('recipients') or {}).get('bcc', []),
    })


def smtp_connect_and_auth(config):
    host = config.get('host', '').strip()
    port = int(config.get('port', 0) or 0)
    user = config.get('user', '').strip()
    password = config.get('password', '')
    security = config.get('security', 'starttls')
    client_hostname = config.get('clientHostname') or None

    if not host:
        raise ValueError("SMTP host is required.")
    if port <= 0:
        raise ValueError("SMTP port is required.")
    if not user:
        raise ValueError("SMTP user is required.")
    if not password:
        raise ValueError("SMTP password is required.")

    if security == 'ssl':
        server = smtplib.SMTP_SSL(host=host, port=port, timeout=30, local_hostname=client_hostname)
        server.ehlo()
    else:
        server = smtplib.SMTP(host=host, port=port, timeout=30, local_hostname=client_hostname)
        server.ehlo()
        if security == 'starttls':
            context = ssl.create_default_context()
            server.starttls(context=context)
            server.ehlo()

    server.login(user, password)
    return server


def send_email_via_smtp(config, subject, html_body, recipients):
    sender = config.get('user', '').strip()
    if not sender:
        raise ValueError("SMTP user is empty.")

    to_list = recipients.get('to', [])
    cc_list = recipients.get('cc', [])
    bcc_list = recipients.get('bcc', [])
    all_recipients = [*to_list, *cc_list, *bcc_list]
    if len(all_recipients) == 0:
        raise ValueError("No recipients provided.")

    msg = EmailMessage()
    msg['Subject'] = subject or 'DOINg PM report'
    msg['From'] = sender
    msg['To'] = ', '.join(to_list) if to_list else 'undisclosed-recipients:;'
    if cc_list:
        msg['Cc'] = ', '.join(cc_list)
    msg.set_content("This message contains an HTML PM report. Please view using an HTML-capable email client.")
    msg.add_alternative(html_body or '<p>No report content.</p>', subtype='html')

    server = smtp_connect_and_auth(config)
    try:
        server.send_message(msg, from_addr=sender, to_addrs=all_recipients)
    finally:
        try:
            server.quit()
        except Exception:
            pass


def process_due_pm_email_jobs():
    # Step 1: snapshot due jobs + smtp config
    with DB_LOCK:
        if not os.path.exists(DB_FILE):
            return
        with open(DB_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        jobs = data.get('pmEmailJobs', []) or []
        smtp_config = normalize_smtp_config(data.get('smtpConfig') or {})
        now_utc = datetime.now(timezone.utc)
        due_jobs = []
        for job in jobs:
            if job.get('status') != 'pending':
                continue
            schedule_at = parse_iso_datetime(job.get('scheduleAt'))
            if schedule_at is None:
                continue
            if schedule_at <= now_utc:
                due_jobs.append({
                    "id": job.get('id'),
                    "subject": job.get('subject', ''),
                    "htmlBody": job.get('htmlBody', ''),
                    "recipients": (job.get('recipients') or {})
                })

    if len(due_jobs) == 0:
        return

    # Step 2: send emails without DB lock
    now_iso = datetime.now(timezone.utc).isoformat()
    outcomes = {}
    for job in due_jobs:
        job_id = job.get('id')
        if not job_id:
            continue
        try:
            recipients = recipients_from_payload(job.get('recipients') or {})
            send_email_via_smtp(smtp_config, job.get('subject', ''), job.get('htmlBody', ''), recipients)
            outcomes[job_id] = {"status": "sent", "sentAt": now_iso, "lastError": ""}
        except Exception as e:
            outcomes[job_id] = {"status": "failed", "lastError": str(e)[:500]}

    if len(outcomes) == 0:
        return

    # Step 3: persist statuses atomically
    with DB_LOCK:
        if not os.path.exists(DB_FILE):
            return
        with open(DB_FILE, 'r', encoding='utf-8') as f:
            latest = json.load(f)
        changed = False
        for job in (latest.get('pmEmailJobs', []) or []):
            job_id = job.get('id')
            if not job_id or job_id not in outcomes:
                continue
            # Do not overwrite if user already edited/cancelled the job.
            if job.get('status') != 'pending':
                continue
            result = outcomes[job_id]
            job['status'] = result.get('status', 'failed')
            job['lastTriedAt'] = now_iso
            if result.get('status') == 'sent':
                job['sentAt'] = result.get('sentAt', now_iso)
                job['lastError'] = ''
            else:
                job['lastError'] = result.get('lastError', 'Unknown SMTP error')
            changed = True
        if changed:
            latest['lastUpdated'] = int(time.time() * 1000)
            with open(DB_FILE, 'w', encoding='utf-8') as f:
                json.dump(latest, f, indent=2)


def scheduler_loop():
    print(f"[Mailer] PM email scheduler started (interval: {SCHEDULER_INTERVAL_SECONDS}s)")
    while True:
        try:
            process_due_pm_email_jobs()
        except Exception as e:
            print(f"[Mailer] Scheduler error: {e}")
        time.sleep(SCHEDULER_INTERVAL_SECONDS)

# Variable globale pour le chemin de la base de données
DB_FILE = get_db_path()

def init_db_if_needed():
    """Initialise le fichier db.json avec la structure par défaut s'il n'existe pas."""
    directory = os.path.dirname(DB_FILE)
    if directory and not os.path.exists(directory):
        try:
            os.makedirs(directory, exist_ok=True)
            print(f"[Init] Created directory: {directory}")
        except Exception as e:
            print(f"[Error] Could not create directory {directory}: {e}")

    if not os.path.exists(DB_FILE):
        # Données initiales avec l'utilisateur administrateur système
        initial_data = {
            "users": [{
                "id": "u1", "uid": "Admin", "firstName": "System", "lastName": "Admin",
                "functionTitle": "Administrator", "role": "Admin", "password": "admin"
            }],
            "teams": [],  # Équipes (contiennent les projets)
            "meetings": [],  # Réunions enregistrées
            "weeklyReports": [],  # Rapports hebdomadaires
            "workingGroups": [],  # Groupes de travail
            "smartTodos": [],  # Smart To Do
            "oneOffQueries": [],  # One Off Queries
            "pmGantData": [],  # PM Gant
            "pmReportData": [],  # PM Status Report
            "smtpConfig": default_smtp_config(),  # SMTP config for real email delivery
            "pmEmailJobs": [],  # Scheduled PM email jobs + delivery history
            "notifications": [],  # Notifications système
            "dismissedAlerts": {},  # Alertes rejetées
            "systemMessage": { "active": False, "content": "", "level": "info" },  # Message système global
            "notes": [],
            "lastUpdated": int(time.time() * 1000)  # Timestamp de création
        }
        try:
            with open(DB_FILE, 'w', encoding='utf-8') as f:
                json.dump(initial_data, f, indent=2)
            print(f"[Init] ✅ Base de données créée avec succès : {DB_FILE}")
        except Exception as e:
            print(f"[Error] Impossible de créer db.json : {e}")
    else:
        print(f"[Init] Base de données existante trouvée : {DB_FILE}")

# Initialisation au démarrage du serveur
init_db_if_needed()

# --- ROUTES API PRINCIPALES ---

@app.route('/api/data', methods=['GET'])
def get_data():
    """Endpoint de LECTURE: Récupère toutes les données de db.json."""
    try:
        with DB_LOCK:
            if os.path.exists(DB_FILE):
                with open(DB_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            else:
                data = {}
        if data:
            # Ensure new top-level fields exist for backward compatibility
            if not isinstance(data.get('smtpConfig'), dict):
                data['smtpConfig'] = default_smtp_config()
            else:
                data['smtpConfig'] = normalize_smtp_config(data.get('smtpConfig'))
            if not isinstance(data.get('pmEmailJobs'), list):
                data['pmEmailJobs'] = []
            return jsonify(data)
        return jsonify({})
    except Exception as e:
        print(f"Erreur lecture fichier: {e}")
        return jsonify({"error": "Erreur lecture données"}), 500

@app.route('/api/data', methods=['POST'])
def save_data():
    """Endpoint d'ÉCRITURE: Sauvegarde les données dans db.json avec gestion des conflits de concurrence."""
    try:
        new_data = request.json or {}
        client_base_version = request.headers.get('X-Base-Version')
        new_data['smtpConfig'] = normalize_smtp_config(new_data.get('smtpConfig') or {})
        if not isinstance(new_data.get('pmEmailJobs'), list):
            new_data['pmEmailJobs'] = []

        with DB_LOCK:
            # Charge la version actuelle de la DB
            current_db_data = {}
            if os.path.exists(DB_FILE):
                with open(DB_FILE, 'r', encoding='utf-8') as f:
                    current_db_data = json.load(f)

            # Vérification de Concurrence (Optimistic Locking)
            # Si le client envoie une version de base, on vérifie si la DB n'a pas été modifiée entre temps
            if client_base_version and client_base_version != 'force':
                server_version = str(current_db_data.get('lastUpdated', 0))
                if server_version != str(client_base_version):
                    print(f"[Conflit] Version Client: {client_base_version} vs Serveur: {server_version}")
                    # Retourne les données serveur pour permettre une fusion
                    return jsonify({
                        "error": "Conflict detected",
                        "serverData": current_db_data
                    }), 409

            # Mise à jour du timestamp de la dernière modification
            new_data['lastUpdated'] = int(time.time() * 1000)

            with open(DB_FILE, 'w', encoding='utf-8') as f:
                json.dump(new_data, f, indent=2)
            
        print(f"[Sauvegarde] Données mises à jour à {time.strftime('%H:%M:%S')}")
        return jsonify({"success": True, "timestamp": new_data['lastUpdated']})
    except Exception as e:
        print(f"Erreur écriture fichier: {e}")
        return jsonify({"error": "Erreur sauvegarde données"}), 500

# --- ENDPOINTS DE CONFIGURATION ---
# Permettent de modifier le chemin de la base de données

@app.route('/api/config/db-path', methods=['GET'])
def get_db_config_path():
    """Retourne le chemin actuel de la base de données."""
    return jsonify({"path": DB_FILE})

@app.route('/api/config/db-path', methods=['POST'])
def update_db_config_path():
    """Modifie le chemin de la base de données et crée les répertoires nécessaires."""
    global DB_FILE
    new_path = request.json.get('path')
    if not new_path:
        return jsonify({"error": "Path required"}), 400

    try:
        # Sauvegarde la nouvelle configuration
        with open(SERVER_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({"dbPath": new_path}, f, indent=2)

        DB_FILE = new_path

        # Vérifie que la DB existe au nouvel emplacement
        if not os.path.exists(DB_FILE):
            directory = os.path.dirname(DB_FILE)
            if directory and not os.path.exists(directory):
                os.makedirs(directory, exist_ok=True)

            # Crée une DB vierge avec l'admin par défaut
            initial_data = {
                "users": [{ "id": "u1", "uid": "Admin", "firstName": "System", "lastName": "Admin", "functionTitle": "Administrator", "role": "Admin", "password": "admin" }],
                "teams": [], "meetings": [], "weeklyReports": [], "workingGroups": [],
                "smartTodos": [], "oneOffQueries": [], "pmGantData": [], "pmReportData": [],
                "smtpConfig": default_smtp_config(), "pmEmailJobs": [],
                "notifications": [],
                "dismissedAlerts": {}, "systemMessage": { "active": False, "content": "", "level": "info" },
                "notes": [], "lastUpdated": int(time.time() * 1000)
            }
            with open(DB_FILE, 'w', encoding='utf-8') as f:
                json.dump(initial_data, f, indent=2)
                
        print(f"[Config] Chemin DB mis à jour vers: {DB_FILE}")
        return jsonify({"success": True, "path": DB_FILE})
    except Exception as e:
        print(f"Erreur lors de la mise à jour de la config: {e}")
        return jsonify({"error": "Échec de la mise à jour du chemin DB"}), 500


# --- ENDPOINTS SMTP / EMAIL ---
@app.route('/api/mail/test', methods=['POST'])
def test_mail_connection():
    """Teste la connectivité SMTP avec la configuration passée ou celle stockée côté serveur."""
    try:
        payload = request.json or {}
        smtp_config = payload.get('smtpConfig')

        if smtp_config is None:
            with DB_LOCK:
                if os.path.exists(DB_FILE):
                    with open(DB_FILE, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                else:
                    data = {}
            smtp_config = data.get('smtpConfig') or default_smtp_config()

        smtp_config = normalize_smtp_config(smtp_config)
        server = smtp_connect_and_auth(smtp_config)
        try:
            pass
        finally:
            try:
                server.quit()
            except Exception:
                pass
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


@app.route('/api/mail/send', methods=['POST'])
def send_mail_now():
    """Envoie immédiatement un email HTML en utilisant la configuration SMTP stockée."""
    try:
        payload = request.json or {}
        subject = str(payload.get('subject', '')).strip()
        html_body = str(payload.get('html', '')).strip()
        recipients = recipients_from_payload(payload)

        with DB_LOCK:
            if os.path.exists(DB_FILE):
                with open(DB_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            else:
                data = {}
        smtp_config = normalize_smtp_config(data.get('smtpConfig') or {})
        send_email_via_smtp(smtp_config, subject, html_body, recipients)

        return jsonify({"success": True})
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        print(f"[Mailer] Immediate send failed: {e}")
        return jsonify({"error": str(e)}), 500

# --- SERVICE DE L'APPLICATION FRONTEND (REACT) ---
# Sert l'interface React en production et gère le React Router

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    """Sert les fichiers statiques et retourne index.html pour React Router."""
    if path != "" and os.path.exists(os.path.join(app.static_folder, path)):
        return send_from_directory(app.static_folder, path)
    else:
        # Retourne index.html pour que React Router gère le routage côté client
        if os.path.exists(os.path.join(app.static_folder, 'index.html')):
            return send_from_directory(app.static_folder, 'index.html')
        else:
            return "Serveur API en cours d'exécution. Frontend non construit (vérifiez le dossier 'dist'). Utilisez npm run dev pour le développement.", 200

if __name__ == '__main__':
    debug_mode = True
    should_start_scheduler = os.environ.get('WERKZEUG_RUN_MAIN') == 'true' or not debug_mode
    if should_start_scheduler:
        threading.Thread(target=scheduler_loop, daemon=True, name='pm-email-scheduler').start()
    print(f"\n📡 SERVEUR API PYTHON (FLASK) LANCÉ !\n-------------------------------------\nPort API        : {PORT}\nFichier Données : {DB_FILE}\n-------------------------------------\n")
    app.run(port=PORT, debug=debug_mode, use_reloader=True)
