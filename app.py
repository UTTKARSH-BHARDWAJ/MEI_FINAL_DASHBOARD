import csv
import json
import re
import os
import io
import sys
import uuid
import signal
import logging
import platform
import sqlite3
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from contextlib import closing
from logging.handlers import RotatingFileHandler
import pandas as pd
import openpyxl
import xlrd

from flask import Flask, render_template, request, jsonify, send_from_directory, Response, redirect, url_for
from werkzeug.utils import secure_filename

# ---------------------------------------------------------------------------
# Setup Directories & Logging
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STORAGE_DIR = os.path.join(BASE_DIR, 'storage')
LOG_DIR = os.path.join(STORAGE_DIR, 'logs')
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

os.makedirs(STORAGE_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

MAX_LOG_LINES = 500
MAX_LOG_AGE_SECONDS = 86400
SERVER_LOG_PATH = os.path.join(LOG_DIR, 'server.log')
_log_file_lock = threading.Lock()

def _prune_old_log_entries(log_path=SERVER_LOG_PATH, max_age_seconds=MAX_LOG_AGE_SECONDS, max_lines=MAX_LOG_LINES):
    if not os.path.exists(log_path):
        return
    cutoff = datetime.now() - timedelta(seconds=max_age_seconds)
    with _log_file_lock:
        try:
            with open(log_path, 'r', encoding='utf-8', errors='replace') as f:
                lines = f.readlines()
            
            kept_lines = []
            for line in lines:
                if len(line) >= 19:
                    try:
                        log_dt = datetime.strptime(line[:19], '%Y-%m-%d %H:%M:%S')
                        if log_dt >= cutoff:
                            kept_lines.append(line)
                        continue
                    except ValueError:
                        pass
                kept_lines.append(line)
            
            if len(kept_lines) > max_lines:
                kept_lines = kept_lines[-max_lines:]
                
            if len(kept_lines) != len(lines):
                with open(log_path, 'w', encoding='utf-8', errors='replace') as f:
                    f.writelines(kept_lines)
        except Exception:
            pass

def _start_log_cleaner_daemon():
    def _cleaner_loop():
        import time
        while True:
            time.sleep(5)
            _prune_old_log_entries()
    t = threading.Thread(target=_cleaner_loop, daemon=True)
    t.start()

_start_log_cleaner_daemon()

class MaxLinesFileHandler(logging.Handler):
    """Custom logging handler that maintains logs within 1 minute age and at most max_lines."""
    def __init__(self, filename, max_lines=MAX_LOG_LINES, encoding='utf-8'):
        super().__init__()
        self.filename = os.path.abspath(filename)
        self.max_lines = max_lines
        self.encoding = encoding
        _prune_old_log_entries(self.filename, MAX_LOG_AGE_SECONDS, self.max_lines)

    def emit(self, record):
        try:
            msg = self.format(record) + '\n'
            with _log_file_lock:
                lines = []
                if os.path.exists(self.filename):
                    with open(self.filename, 'r', encoding=self.encoding, errors='replace') as f:
                        lines = f.readlines()
                lines.append(msg)
                
                cutoff = datetime.now() - timedelta(seconds=MAX_LOG_AGE_SECONDS)
                kept = []
                for l in lines:
                    if len(l) >= 19:
                        try:
                            dt = datetime.strptime(l[:19], '%Y-%m-%d %H:%M:%S')
                            if dt >= cutoff:
                                kept.append(l)
                            continue
                        except ValueError:
                            pass
                    kept.append(l)
                
                if len(kept) > self.max_lines:
                    kept = kept[-self.max_lines:]
                    
                with open(self.filename, 'w', encoding=self.encoding, errors='replace') as f:
                    f.writelines(kept)
        except Exception:
            self.handleError(record)

log_formatter = logging.Formatter(
    '%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S'
)

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(log_formatter)

file_handler = MaxLinesFileHandler(
    SERVER_LOG_PATH,
    max_lines=MAX_LOG_LINES,
    encoding='utf-8',
)
file_handler.setFormatter(log_formatter)

logging.basicConfig(level=logging.INFO, handlers=[console_handler, file_handler])
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# App Config
# ---------------------------------------------------------------------------
app = Flask(__name__, template_folder='templates', static_folder='static')
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB max upload cap

ALARM_DB_PATH = os.path.join(STORAGE_DIR, 'alarms.db')
REJECTION_DB_PATH = os.path.join(STORAGE_DIR, 'rejections.db')

_alarm_db_lock = threading.Lock()
_rejection_db_lock = threading.Lock()

MAX_FILES = 100
RETENTION_DAYS = 60

# ---------------------------------------------------------------------------
# Security Headers Middleware
# ---------------------------------------------------------------------------
@app.after_request
def set_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Permissions-Policy'] = 'geolocation=(), camera=(), microphone=()'
    return response

# ===========================================================================
# 1. ALARM DASHBOARD STORAGE & LOGIC (Preserved word-for-word)
# ===========================================================================

def get_alarm_db():
    conn = sqlite3.connect(ALARM_DB_PATH, timeout=30)
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

def init_alarm_db():
    with closing(get_alarm_db()):
        pass

init_alarm_db()

def insert_alarm_records(conn, records):
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    with _alarm_db_lock:
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

def fetch_all_alarm_records(conn):
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

def parse_datetime_flexible(dt_val):
    if dt_val is None or dt_val == '':
        return None
    if isinstance(dt_val, (datetime, pd.Timestamp)):
        return dt_val
    dt_str = str(dt_val).strip()
    if not dt_str:
        return None
    patterns = [
        '%d/%m/%Y %H:%M:%S',
        '%m/%d/%Y %I:%M:%S %p',
        '%d/%m/%Y %I:%M:%S %p',
        '%m/%d/%Y %H:%M:%S',
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%dT%H:%M:%S',
        '%d-%m-%Y %H:%M:%S',
        '%m-%d-%Y %H:%M:%S',
        '%d/%m/%Y %H:%M',
        '%m/%d/%Y %H:%M',
        '%d/%m/%Y %I:%M %p',
        '%m/%d/%Y %I:%M %p',
        '%Y-%m-%d %H:%M',
        '%d-%m-%Y',
        '%m/%d/%Y',
        '%d/%m/%Y',
        '%Y-%m-%d'
    ]
    for fmt in patterns:
        try:
            return datetime.strptime(dt_str, fmt)
        except ValueError:
            pass
    try:
        ts = pd.to_datetime(dt_str, errors='coerce', dayfirst=True)
        if pd.notna(ts):
            return ts.to_pydatetime()
    except Exception:
        pass
    return None

def _process_rows(rows):
    """Core logic: transform row dicts into categorized records."""
    records = []
    for raw_row in rows:
        row = {str(k).replace('\ufeff', '').strip(): v for k, v in raw_row.items() if k is not None}
        machine_name = row.get('Machine Name', 'Unknown')
        if not machine_name.startswith('Racer'):
            continue
            
        dt_str = row.get('DateTime', '')
        if not dt_str:
            continue
            
        dt_obj = parse_datetime_flexible(dt_str)
        if not dt_obj:
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

def parse_csv_content(file_content_str):
    f = io.StringIO(file_content_str)
    reader = csv.DictReader(f)
    return _process_rows(reader)

def parse_xlsx(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    try:
        header_row = next(rows_iter)
    except StopIteration:
        wb.close()
        return []
    headers = [str(h).strip() if h else '' for h in header_row]
    row_dicts = []
    for row in rows_iter:
        if not row:
            continue
        row_dicts.append({headers[i]: (str(row[i]) if i < len(row) and row[i] is not None else '') for i in range(len(headers))})
    wb.close()
    return _process_rows(row_dicts)

def parse_xls(file_bytes):
    wb = xlrd.open_workbook(file_contents=file_bytes)
    ws = wb.sheet_by_index(0)
    headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]
    row_dicts = []
    for r in range(1, ws.nrows):
        row_dicts.append({headers[c]: str(ws.cell_value(r, c)) for c in range(ws.ncols)})
    return _process_rows(row_dicts)

def parse_ods(file_bytes):
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
    header_cells = all_rows[0].getElementsByType(TableCell)
    headers = []
    for cell in header_cells:
        ps = cell.getElementsByType(P)
        text = ''.join([p.firstChild.data if p.firstChild else '' for p in ps]).strip()
        repeat = int(cell.getAttribute('numbercolumnsrepeated') or 1)
        headers.extend([text] * repeat)
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

def parse_alarm_file(file_storage):
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

def aggregate_alarm_records(records):
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

ALLOWED_EXTENSIONS = {'.csv', '.xlsx', '.xls', '.ods'}

def _validate_alarm_files(files):
    if not files or files[0].filename == '':
        return None, (jsonify({"error": "No selected files"}), 400)
    for f in files:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            return None, (jsonify({"error": f"Unsupported file type: {ext}"}), 400)
    return files[:100], None


# ===========================================================================
# 2. REJECTION DASHBOARD STORAGE & LOGIC (Preserved word-for-word)
# ===========================================================================

def get_rejection_db():
    conn = sqlite3.connect(REJECTION_DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("""
        CREATE TABLE IF NOT EXISTS rejections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            machine_name TEXT NOT NULL,
            reject_detail TEXT,
            message TEXT,
            hour INTEGER,
            job_name TEXT,
            ingested_at TEXT NOT NULL,
            UNIQUE(timestamp, machine_name, reject_detail, message, job_name)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_timestamp ON rejections(timestamp)")
    try:
        conn.execute("ALTER TABLE rejections ADD COLUMN nr_order TEXT")
    except sqlite3.OperationalError:
        pass
    conn.commit()
    return conn

def init_rejection_db():
    with closing(get_rejection_db()):
        pass

init_rejection_db()

def prune_old_rejection_records(conn):
    cutoff = (datetime.now() - timedelta(days=RETENTION_DAYS)).strftime("%Y-%m-%dT%H:%M:%S")
    conn.execute("DELETE FROM rejections WHERE timestamp < ?", (cutoff,))
    conn.commit()

def insert_rejection_records(conn, records):
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    with _rejection_db_lock:
        conn.executemany(
            """INSERT OR IGNORE INTO rejections
               (timestamp, machine_name, reject_detail, message, hour, job_name, nr_order, ingested_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            [
                (r.get("Timestamp"), r.get("Machine Name"), r.get("Reject Detail"),
                 r.get("Message"), r.get("Hour"), r.get("Job Name"), r.get("Nr Order"), now)
                for r in records
            ],
        )
        conn.commit()

def fetch_all_rejection_records(conn):
    rows = conn.execute(
        "SELECT timestamp, machine_name, reject_detail, message, hour, job_name, nr_order "
        "FROM rejections ORDER BY timestamp"
    ).fetchall()
    return [
        {
            "Timestamp": r["timestamp"],
            "Machine Name": r["machine_name"],
            "Reject Detail": r["reject_detail"],
            "Message": r["message"],
            "Hour": r["hour"],
            "Job Name": r["job_name"],
            "Nr Order": r["nr_order"],
        }
        for r in rows
    ]

def read_rejection_file_to_df(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        return pd.read_csv(filepath)
    elif ext in (".xlsx", ".xls"):
        return pd.read_excel(filepath, engine="openpyxl")
    elif ext == ".ods":
        return pd.read_excel(filepath, engine="odf")
    else:
        raise ValueError(f"Unsupported file format: {ext}")

def clean_rejection_data(df):
    df.columns = df.columns.str.strip()
    df["DateTime"] = pd.to_datetime(df["DateTime"], errors="coerce")

    df_r = df[df["Machine Name"].fillna("").str.startswith("Racer")].copy()
    df_r = df_r[df_r["Reject Detail"] != "Twin lens rejected"].copy()
    df_r = df_r.sort_values("DateTime").reset_index(drop=True)

    job_missing = (
        df_r["Job Name"].fillna("").astype(str).str.strip()
        .replace(["?", "---", "nan", "None"], "").eq("")
    )
    order_missing = (
        df_r["Nr Order"].isna()
        | df_r["Nr Order"].astype(str).str.strip().isin(["", "nan", "None"])
    )

    is_c1 = (~job_missing) & (~order_missing)
    is_c2 = (~job_missing) & (order_missing)
    is_c3 = (job_missing) & (order_missing)

    # Case 1
    df_c1_clean = df_r[is_c1].drop_duplicates(subset=["Job Name", "Nr Order"], keep="first")

    WINDOW_DAY = pd.Timedelta(hours=7.0)
    WINDOW_NIGHT = pd.Timedelta(hours=14)
    DAY_LIMIT = pd.Timedelta(hours=24)

    # Case 2
    accepted_rows_c2 = []
    c2_sub = df_r.loc[is_c2, ["DateTime", "Machine Name"]]
    job_names_c2 = df_r.loc[is_c2, "Job Name"]
    for job_name, grp in c2_sub.groupby(job_names_c2, sort=False):
        grp = grp.sort_values("DateTime")
        last_kept_machine = None
        last_kept_time = None
        instance_times = []
        for idx, dt, machine in zip(grp.index, grp["DateTime"], grp["Machine Name"]):
            if last_kept_time is None:
                accepted_rows_c2.append(idx)
                instance_times = [dt]
                last_kept_machine, last_kept_time = machine, dt
                continue

            window = WINDOW_NIGHT if last_kept_time.hour >= 20 else WINDOW_DAY
            gap = dt - last_kept_time

            if machine == last_kept_machine and gap <= window:
                continue

            if machine != last_kept_machine:
                accepted_rows_c2.append(idx)
                last_kept_machine, last_kept_time = machine, dt
                continue

            instance_times = [t for t in instance_times if dt - t < DAY_LIMIT]
            if len(instance_times) >= 2:
                continue
            accepted_rows_c2.append(idx)
            instance_times.append(dt)
            last_kept_machine, last_kept_time = machine, dt

    df_c2_clean = df_r.loc[accepted_rows_c2]

    # Case 3
    accepted_rows_c3 = []
    c3_sub = df_r.loc[is_c3, ["DateTime"]]
    group_keys_c3 = df_r.loc[is_c3, ["Machine Name", "Reject Detail", "Message"]]
    for key, grp in c3_sub.groupby(
        [group_keys_c3["Machine Name"], group_keys_c3["Reject Detail"], group_keys_c3["Message"]],
        sort=False,
    ):
        grp = grp.sort_values("DateTime")
        last_kept_time = None
        for idx, dt in zip(grp.index, grp["DateTime"]):
            if last_kept_time is None:
                accepted_rows_c3.append(idx)
                last_kept_time = dt
                continue
            window = WINDOW_NIGHT if last_kept_time.hour >= 20 else WINDOW_DAY
            if dt - last_kept_time <= window:
                continue
            accepted_rows_c3.append(idx)
            last_kept_time = dt

    df_c3_clean = df_r.loc[accepted_rows_c3]

    final_df = pd.concat([df_c1_clean, df_c2_clean, df_c3_clean]).sort_values("DateTime")
    final_df["Timestamp"] = final_df["DateTime"].dt.strftime("%Y-%m-%dT%H:%M:%S")
    final_df["Hour"] = final_df["DateTime"].dt.hour
    final_df = final_df.fillna("")

    return final_df

def _save_and_read_rejection(file_tuple):
    file_obj, upload_folder = file_tuple
    raw_name = secure_filename(file_obj.filename)
    if not raw_name:
        raise ValueError("Invalid filename")
    filename = f"{uuid.uuid4().hex}_{raw_name}"
    filepath = os.path.join(upload_folder, filename)
    file_obj.save(filepath)
    try:
        return read_rejection_file_to_df(filepath)
    finally:
        if os.path.exists(filepath):
            os.remove(filepath)

def fast_json_response(obj):
    return Response(json.dumps(obj, default=str), mimetype="application/json")

def _rejection_dataset_response():
    with closing(get_rejection_db()) as conn:
        prune_old_rejection_records(conn)
        records = fetch_all_rejection_records(conn)

    if not records:
        return fast_json_response({"data": [], "min_ts": "", "max_ts": ""})

    timestamps = [r["Timestamp"] for r in records if r["Timestamp"]]
    min_ts = min(timestamps)[:16]
    max_dt = pd.to_datetime(max(timestamps)) + pd.Timedelta(minutes=1)
    max_ts = max_dt.strftime("%Y-%m-%dT%H:%M")
    return fast_json_response({"data": records, "min_ts": min_ts, "max_ts": max_ts})

# ===========================================================================
# 3. FLASK ROUTES
# ===========================================================================

@app.route('/')
def root():
    return redirect(url_for('alarm_dashboard'))

@app.route('/alarm')
def alarm_dashboard():
    try:
        with closing(get_alarm_db()) as conn:
            records = fetch_all_alarm_records(conn)
        res = aggregate_alarm_records(records)
        log.info('Alarm Dashboard loaded with %d records', len(records))
        return render_template('alarm.html', data=res['data'], min_ts=res['min_ts'], max_ts=res['max_ts'], unique_msgs=res['unique_msgs'])
    except Exception as exc:
        log.error('Failed to load stored alarm data: %s', exc)
        return render_template('alarm.html', data="[]", min_ts="", max_ts="", unique_msgs=[])

@app.route('/rejection')
def rejection_dashboard():
    return render_template('rejection.html')

@app.route('/health')
def health():
    try:
        with closing(get_alarm_db()) as conn_a:
            conn_a.execute("SELECT 1").fetchone()
        with closing(get_rejection_db()) as conn_r:
            conn_r.execute("SELECT 1").fetchone()
        return jsonify({"status": "healthy", "alarm_db": "ok", "rejection_db": "ok"}), 200
    except Exception as exc:
        log.error('Health check failed: %s', exc)
        return jsonify({"status": "unhealthy", "db": str(exc)}), 503

# --- Alarm Endpoints ---

@app.route('/analyze', methods=['POST'])
@app.route('/api/alarm/analyze', methods=['POST'])
def analyze_alarms():
    files, err = _validate_alarm_files(request.files.getlist('files'))
    if err:
        return err

    all_records = []
    for file in files:
        all_records.extend(parse_alarm_file(file))

    log.info('Analyzed %d alarm records from %d file(s)', len(all_records), len(files))
    res = aggregate_alarm_records(all_records)
    return jsonify(res)

@app.route('/store', methods=['POST'])
@app.route('/api/alarm/store', methods=['POST'])
def store_alarms():
    files, err = _validate_alarm_files(request.files.getlist('files'))
    if err:
        return err

    all_records = []
    for file in files:
        all_records.extend(parse_alarm_file(file))

    try:
        with closing(get_alarm_db()) as conn:
            insert_alarm_records(conn, all_records)
        log.info('Stored %d alarm records from %d file(s) into SQLite', len(all_records), len(files))
    except Exception as exc:
        log.exception('Failed to write stored alarm data to SQLite: %s', exc)
        return jsonify({"error": "Storage failed"}), 500

    with closing(get_alarm_db()) as conn:
        updated_records = fetch_all_alarm_records(conn)
    res = aggregate_alarm_records(updated_records)
    return jsonify(res)

# --- Rejection Endpoints ---

@app.route('/api/data', methods=['GET'])
@app.route('/api/rejection/data', methods=['GET'])
def get_rejection_data():
    try:
        return _rejection_dataset_response()
    except Exception:
        log.exception("Fetch stored rejection data error")
        return jsonify({"error": "Could not load stored rejection data"}), 500

@app.route('/api/analyze', methods=['POST'])
@app.route('/api/rejection/analyze', methods=['POST'])
def analyze_rejection_file():
    files = request.files.getlist("file")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No file provided"}), 400

    files = [f for f in files if f.filename != ""]
    if len(files) > MAX_FILES:
        return jsonify({"error": f"Too many files. Maximum is {MAX_FILES}."}), 400

    try:
        if len(files) == 1:
            df = _save_and_read_rejection((files[0], UPLOAD_FOLDER))
        else:
            with ThreadPoolExecutor(max_workers=min(len(files), 8)) as pool:
                dfs = list(pool.map(_save_and_read_rejection, [(f, UPLOAD_FOLDER) for f in files]))
            df = pd.concat(dfs, ignore_index=True)

        final_df = clean_rejection_data(df)

        records = (
            final_df[["Timestamp", "Machine Name", "Reject Detail", "Message", "Hour", "Job Name", "Nr Order"]]
            .dropna(subset=["Timestamp", "Machine Name"])
            .to_dict(orient="records")
        )

        if not records:
            return fast_json_response({"data": [], "min_ts": "", "max_ts": ""})

        timestamps = [r["Timestamp"] for r in records if r["Timestamp"]]
        min_ts = min(timestamps)[:16]
        max_dt = pd.to_datetime(max(timestamps)) + pd.Timedelta(minutes=1)
        max_ts = max_dt.strftime("%Y-%m-%dT%H:%M")
        return fast_json_response({"data": records, "min_ts": min_ts, "max_ts": max_ts})

    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except KeyError as e:
        return jsonify({"error": f"Missing required column: {e}"}), 400
    except Exception as e:
        log.exception("Analyze processing error")
        return jsonify({"error": f"Processing error: {str(e)}"}), 500

@app.route('/api/upload', methods=['POST'])
@app.route('/api/rejection/upload', methods=['POST'])
def upload_rejection_file():
    files = request.files.getlist("file")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No file provided"}), 400

    files = [f for f in files if f.filename != ""]
    if len(files) > MAX_FILES:
        return jsonify({"error": f"Too many files. Maximum is {MAX_FILES}."}), 400

    try:
        if len(files) == 1:
            df = _save_and_read_rejection((files[0], UPLOAD_FOLDER))
        else:
            with ThreadPoolExecutor(max_workers=min(len(files), 8)) as pool:
                dfs = list(pool.map(_save_and_read_rejection, [(f, UPLOAD_FOLDER) for f in files]))
            df = pd.concat(dfs, ignore_index=True)

        final_df = clean_rejection_data(df)

        new_records = (
            final_df[["Timestamp", "Machine Name", "Reject Detail", "Message", "Hour", "Job Name", "Nr Order"]]
            .dropna(subset=["Timestamp", "Machine Name"])
            .to_dict(orient="records")
        )

        with closing(get_rejection_db()) as conn:
            insert_rejection_records(conn, new_records)
            prune_old_rejection_records(conn)

        return _rejection_dataset_response()

    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except KeyError as e:
        return jsonify({"error": f"Missing required column: {e}"}), 400
    except Exception as e:
        log.exception("Upload processing error")
        return jsonify({"error": f"Processing error: {str(e)}"}), 500

# Error Handlers
@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "File too large. Max upload size is 500MB."}), 413

@app.errorhandler(404)
def not_found(e):
    return jsonify({"error": "Not found"}), 404

@app.errorhandler(500)
def internal_error(e):
    log.exception('Unhandled server error')
    return jsonify({"error": "Internal server error"}), 500

# ---------------------------------------------------------------------------
# Server Startup Helper
# ---------------------------------------------------------------------------
def _run_production(host, port):
    """Cross-platform production WSGI server runner (Waitress for Windows/macOS, Gunicorn for Linux)."""
    system = platform.system()
    
    # On Windows and macOS, Waitress is preferred because it uses multi-threading rather than process fork()
    if system in ('Windows', 'Darwin'):
        try:
            from waitress import serve
            log.info('Starting production Waitress WSGI server on %s:%s', host, port)
            serve(app, host=host, port=port, threads=8, channel_timeout=120)
            return
        except ImportError:
            log.warning('Waitress not installed — falling back to Gunicorn or Flask dev server.')

    # Try Gunicorn (for Linux or if Waitress is not available)
    try:
        import subprocess
        log.info('Starting production Gunicorn WSGI server on %s:%s', host, port)
        env = os.environ.copy()
        env['OBJC_DISABLE_INITIALIZE_FORK_SAFETY'] = 'YES'
        subprocess.call([
            sys.executable, '-m', 'gunicorn',
            'app:app',
            '--bind', f'{host}:{port}',
            '--workers', env.get('WEB_WORKERS', '4'),
            '--timeout', '120',
            '--access-logfile', '-',
        ], env=env)
        return
    except Exception as e:
        log.warning('Failed to launch Gunicorn: %s', e)

    # Final fallback to Waitress / Flask
    try:
        from waitress import serve
        log.info('Starting production Waitress WSGI server on %s:%s', host, port)
        serve(app, host=host, port=port, threads=8, channel_timeout=120)
    except ImportError:
        log.warning('Neither Gunicorn nor Waitress installed — falling back to Flask dev server.')
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
