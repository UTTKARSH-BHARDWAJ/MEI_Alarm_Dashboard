import csv
import json
import re
import os
import io
import sys
import signal
import logging
import platform
import sqlite3
import tempfile
import threading
import openpyxl
import xlrd
from datetime import datetime
from contextlib import closing
from logging.handlers import RotatingFileHandler
from flask import Flask, render_template, request, jsonify

# ---------------------------------------------------------------------------
# Logging — console + rotating file
# ---------------------------------------------------------------------------
STORAGE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'storage')
LOG_DIR = os.path.join(STORAGE_DIR, 'logs')
os.makedirs(LOG_DIR, exist_ok=True)

log_formatter = logging.Formatter(
    '%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S'
)

# Console handler
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(log_formatter)

# Rotating file handler — 5 MB per file, keep 5 backups
file_handler = RotatingFileHandler(
    os.path.join(LOG_DIR, 'server.log'),
    maxBytes=5 * 1024 * 1024,
    backupCount=5,
    encoding='utf-8',
)
file_handler.setFormatter(log_formatter)

logging.basicConfig(level=logging.INFO, handlers=[console_handler, file_handler])
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------
app = Flask(__name__, template_folder='.')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB upload cap

DB_PATH = os.path.join(STORAGE_DIR, 'alarms.db')
os.makedirs(STORAGE_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Database helpers — WAL mode for concurrent reads/writes, busy timeout
# ---------------------------------------------------------------------------
_db_lock = threading.Lock()

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=5000")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS alarms (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            machine_name TEXT NOT NULL,
            message TEXT,
            hour INTEGER,
            category TEXT,
            ingested_at TEXT NOT NULL,
            UNIQUE(timestamp, machine_name, message, category)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_timestamp ON alarms(timestamp)")
    conn.commit()
    return conn

def init_db():
    with closing(get_db()):
        pass

init_db()

# ---------------------------------------------------------------------------
# Security headers middleware
# ---------------------------------------------------------------------------
@app.after_request
def set_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Permissions-Policy'] = 'geolocation=(), camera=(), microphone=()'
    return response

def insert_records(conn, records):
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    with _db_lock:
        conn.executemany(
            """INSERT OR IGNORE INTO alarms
               (timestamp, machine_name, message, hour, category, ingested_at)
               VALUES (?, ?, ?, ?, ?, ?)""",
            [
                (r.get("Timestamp"), r.get("Machine Name"), r.get("Message"),
                 r.get("Hour"), r.get("Category"), now)
                for r in records
            ],
        )
        conn.commit()

def fetch_all_records(conn):
    rows = conn.execute(
        "SELECT timestamp, machine_name, message, hour, category "
        "FROM alarms ORDER BY timestamp"
    ).fetchall()
    return [
        {
            "Timestamp": r["timestamp"],
            "Machine Name": r["machine_name"],
            "Message": r["message"],
            "Hour": r["hour"],
            "Category": r["category"]
        }
        for r in rows
    ]

def parse_spreadsheet_rows(rows):
    """Parse rows (list of dicts with keys like 'Machine Name', 'DateTime', 'Message') into records."""
    return _process_rows(rows)

def parse_csv_content(file_content_str):
    """Parse CSV text content into rows, then process."""
    f = io.StringIO(file_content_str)
    reader = csv.DictReader(f)
    return _process_rows(reader)

def parse_xlsx(file_bytes):
    """Parse .xlsx file bytes into records."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    headers = [str(h).strip() if h else '' for h in next(rows_iter)]
    row_dicts = []
    for row in rows_iter:
        row_dicts.append({headers[i]: (str(row[i]) if row[i] is not None else '') for i in range(len(headers))})
    wb.close()
    return _process_rows(row_dicts)

def parse_xls(file_bytes):
    """Parse .xls file bytes into records."""
    wb = xlrd.open_workbook(file_contents=file_bytes)
    ws = wb.sheet_by_index(0)
    headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]
    row_dicts = []
    for r in range(1, ws.nrows):
        row_dicts.append({headers[c]: str(ws.cell_value(r, c)) for c in range(ws.ncols)})
    return _process_rows(row_dicts)

def parse_ods(file_bytes):
    """Parse .ods file bytes into records using odfpy."""
    from odf.opendocument import load as odf_load
    from odf.table import Table, TableRow, TableCell
    from odf.text import P
    doc = odf_load(io.BytesIO(file_bytes))
    sheets = doc.getElementsByType(Table)
    if not sheets:
        return []
    sheet = sheets[0]
    all_rows = sheet.getElementsByType(TableRow)
    if not all_rows:
        return []
    # get headers
    header_cells = all_rows[0].getElementsByType(TableCell)
    headers = []
    for cell in header_cells:
        ps = cell.getElementsByType(P)
        text = ''.join([p.firstChild.data if p.firstChild else '' for p in ps]).strip()
        repeat = int(cell.getAttribute('numbercolumnsrepeated') or 1)
        headers.extend([text] * repeat)
    # drop empty headers
    while headers and headers[-1] == '':
        headers.pop()
    row_dicts = []
    for row in all_rows[1:]:
        cells = row.getElementsByType(TableCell)
        values = []
        for cell in cells:
            ps = cell.getElementsByType(P)
            text = ''.join([p.firstChild.data if p.firstChild else '' for p in ps]).strip()
            repeat = int(cell.getAttribute('numbercolumnsrepeated') or 1)
            values.extend([text] * repeat)
        values = values[:len(headers)]
        if len(values) < len(headers):
            values.extend([''] * (len(headers) - len(values)))
        if any(v for v in values):
            row_dicts.append({headers[i]: values[i] for i in range(len(headers))})
    return _process_rows(row_dicts)

def parse_file(file_storage):
    """Dispatch to the correct parser based on file extension."""
    filename = file_storage.filename.lower()
    if filename.endswith('.csv'):
        content = file_storage.read().decode('utf-8', errors='replace')
        return parse_csv_content(content)
    elif filename.endswith('.xlsx'):
        return parse_xlsx(file_storage.read())
    elif filename.endswith('.xls'):
        return parse_xls(file_storage.read())
    elif filename.endswith('.ods'):
        return parse_ods(file_storage.read())
    else:
        return []

def _process_rows(rows):
    """Core logic: transform row dicts into categorized records."""
    records = []
    for row in rows:
        machine_name = row.get('Machine Name', 'Unknown')
        if not machine_name.startswith('Racer'):
            continue
            
        dt_str = row.get('DateTime', '')
        if not dt_str:
            continue
            
        try:
            dt_obj = datetime.strptime(dt_str, '%m/%d/%Y %I:%M:%S %p')
        except ValueError:
            continue
            
        msg = row.get('Message', '')
        category = msg
        
        # parse category from message string
        dash_idx = msg.find(' - ')
        if dash_idx != -1:
            rest = msg[dash_idx + 3:]
            colon_idx = rest.find(':')
            if colon_idx != -1:
                category = rest[:colon_idx].strip()
            else:
                category = rest.strip()

        # normalize tray pos
        if category.startswith('Tray pos.'):
            category = 'Tray pos.'

        # normalize analog input
        if category.lower().startswith('analog input terminal'):
            category = 'Analog Input terminal'

        # normalize L0I axis
        if category.lower().startswith('l0i axis'):
            category = 'L0I Axis'

        # normalize pickup variations
        if category.lower().startswith('pickup'):
            category = 'Pickup'

        # normalize R-Axis
        if category.lower().startswith('r axis') or category.lower().startswith('r-axis'):
            category = 'R-Axis'

        # normalize cutting variations
        if category.lower().startswith('cut'):
            category = 'Cut Area'

        # normalize spindle to cut area
        if category.lower().startswith('spindle'):
            category = 'Cut Area'

        # normalize inspection alarms
        if category.lower().startswith('lcd') or 'thickness' in category.lower() or 'inspection' in category.lower():
            category = 'Inspection Area'

        # normalize tray
        if 'tray' in category.lower():
            category = 'Tray Transport'

        # normalize C-Axis to cut area
        if 'axis' in category.lower() and re.search(r'\b[a-z]\d+c\b', category.lower()):
            category = 'Cut Area'

        # normalize I-Axis to inspection area
        if 'axis' in category.lower() and re.search(r'\b[a-z]+\d+i\b', category.lower()):
            category = 'Inspection Area'

        # normalize deposits
        if category.lower().startswith('deposit'):
            category = 'Deposits'

        # normalize door open
        if 'door open' in category.lower():
            category = 'Door Open'

        # normalize unexpected lens
        if category.lower().startswith('unexpectedlens'):
            category = 'UnexpectedLens'

        # normalize emergencies
        if 'please generate a log and send it to mei' in category.lower():
            category = 'Unexpected behavior/ Emergency. Please generate a log and send it to MEI'

        # normalize unload
        if category.lower().startswith('unload'):
            category = 'Unload'

        # normalize load group
        if category.lower().startswith('load group') or category.lower().startswith('load gripper'):
            category = 'Load Group/ gripper'

        # normalize table
        if category.lower().startswith('table'):
            category = 'Table'
        
        records.append({
            'DateTime': dt_str,
            'Machine Name': machine_name,
            'Message': msg,
            'Timestamp': dt_obj.strftime('%Y-%m-%dT%H:%M:%S'),
            'Hour': dt_obj.hour,
            'Category': category
        })

    return records

def aggregate_records(records):
    if not records:
        return {"data": "[]", "min_ts": "", "max_ts": "", "unique_msgs": [], "count": 0}

    timestamps = [r['Timestamp'] for r in records]
    _min_ts = min(timestamps)[:16]
    _max_ts = max(timestamps)[:16]

    _unique_msgs = sorted(list(set(r['Category'] for r in records)))
    _json_data = json.dumps(records)

    return {
        "data": _json_data,
        "min_ts": _min_ts,
        "max_ts": _max_ts,
        "unique_msgs": _unique_msgs,
        "count": len(records)
    }
    
# ---------------------------------------------------------------------------
# Health check endpoint
# ---------------------------------------------------------------------------
@app.route('/health')
def health():
    """Lightweight health check for load balancers / monitoring."""
    try:
        with closing(get_db()) as conn:
            conn.execute("SELECT 1").fetchone()
        return jsonify({"status": "healthy", "db": "ok"}), 200
    except Exception as exc:
        log.error('Health check failed: %s', exc)
        return jsonify({"status": "unhealthy", "db": str(exc)}), 503

# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route('/')
def dashboard():
    try:
        with closing(get_db()) as conn:
            records = fetch_all_records(conn)
        res = aggregate_records(records)
        log.info('Dashboard loaded with %d records', len(records))
        return render_template('index.html', data=res['data'], min_ts=res['min_ts'], max_ts=res['max_ts'], unique_msgs=res['unique_msgs'])
    except Exception as exc:
        log.error('Failed to load stored data: %s', exc)
        return render_template('index.html', data="[]", min_ts="", max_ts="", unique_msgs=[])

ALLOWED_EXTENSIONS = {'.csv', '.xlsx', '.xls', '.ods'}

def _validate_files(files):
    """Reject empty uploads and unsupported extensions."""
    if not files or files[0].filename == '':
        return None, (jsonify({"error": "No selected files"}), 400)
    for f in files:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            return None, (jsonify({"error": f"Unsupported file type: {ext}"}), 400)
    return files[:100], None

@app.route('/analyze', methods=['POST'])
def analyze():
    files, err = _validate_files(request.files.getlist('files'))
    if err:
        return err

    all_records = []
    for file in files:
        all_records.extend(parse_file(file))

    log.info('Analyzed %d records from %d file(s)', len(all_records), len(files))
    res = aggregate_records(all_records)
    return jsonify(res)

@app.route('/store', methods=['POST'])
def store():
    files, err = _validate_files(request.files.getlist('files'))
    if err:
        return err

    all_records = []
    for file in files:
        all_records.extend(parse_file(file))

    try:
        with closing(get_db()) as conn:
            insert_records(conn, all_records)
        log.info('Stored %d records from %d file(s) into SQLite', len(all_records), len(files))
    except Exception as exc:
        log.exception('Failed to write stored data to SQLite: %s', exc)
        return jsonify({"error": "Storage failed"}), 500

    # Fetch updated records to return
    with closing(get_db()) as conn:
        updated_records = fetch_all_records(conn)
    res = aggregate_records(updated_records)
    return jsonify(res)

@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "File too large. Max upload size is 50MB."}), 413

@app.errorhandler(404)
def not_found(e):
    return jsonify({"error": "Not found"}), 404

@app.errorhandler(500)
def internal_error(e):
    log.exception('Unhandled server error')
    return jsonify({"error": "Internal server error"}), 500

# ---------------------------------------------------------------------------
# Cross-platform production server
# ---------------------------------------------------------------------------
def _run_production(host, port):
    """Pick the best available WSGI server for the current platform."""
    if platform.system() == 'Windows':
        # Gunicorn does not support Windows — use Waitress
        try:
            from waitress import serve
            log.info('Starting Waitress server on %s:%s', host, port)
            serve(app, host=host, port=port, threads=4,
                  channel_timeout=120, map_size=100000)
        except ImportError:
            log.warning('Waitress not installed — falling back to Flask dev server. '
                        'Install waitress: pip install waitress')
            app.run(host=host, port=port, threaded=True)
    else:
        # On Linux/macOS prefer Gunicorn, fall back gracefully
        try:
            # When run via `python server.py`, spawn gunicorn programmatically
            import subprocess
            log.info('Starting Gunicorn on %s:%s', host, port)
            subprocess.call([
                sys.executable, '-m', 'gunicorn',
                'server:app',
                '--bind', f'{host}:{port}',
                '--workers', os.environ.get('WEB_WORKERS', '2'),
                '--timeout', '120',
                '--access-logfile', '-',
            ])
        except Exception:
            log.warning('Gunicorn not available — falling back to Flask dev server. '
                        'Install gunicorn: pip install gunicorn')
            app.run(host=host, port=port, threaded=True)


if __name__ == '__main__':
    host = os.environ.get('HOST', '0.0.0.0')
    port = int(os.environ.get('PORT', 8080))
    debug = os.environ.get('FLASK_DEBUG', '0') == '1'

    if debug:
        log.info('Starting Flask DEV server on %s:%s (debug=True)', host, port)
        app.run(host=host, port=port, debug=True, threaded=True)
    else:
        _run_production(host, port)
