import os
import sys
import atexit
import subprocess
from dotenv import load_dotenv
import time
import threading
import hashlib
import smtplib
import ssl
import re
import serial
import pyttsx3
import tempfile
from copy import deepcopy

try:
    import edge_tts
    import asyncio
    EDGE_TTS_AVAILABLE = True
except Exception:
    edge_tts = None
    asyncio = None
    EDGE_TTS_AVAILABLE = False

try:
    import pygame
    PYGAME_AVAILABLE = True
except Exception:
    pygame = None
    PYGAME_AVAILABLE = False

PYTTSX3_AVAILABLE = True
from email.message import EmailMessage
from datetime import datetime, date, timedelta

from reportlab.graphics import renderPM, renderSVG
try:
    from reportlab.graphics.barcode import createBarcodeDrawing
except ImportError:
    from reportlab.graphics.barcode.widgets import BarcodeCode128

    def createBarcodeDrawing(
        barcode_type,
        value,
        barHeight=72,
        barWidth=0.42,
        humanReadable=False,
        quiet=True,
        **kwargs,
    ):
        """
        Compatibility fallback for ReportLab versions where
        reportlab.graphics.barcode.createBarcodeDrawing is not available.
        """
        if str(barcode_type).lower() != "code128":
            raise ValueError("Only Code128 barcode is supported by this fallback.")

        barcode = BarcodeCode128()
        barcode.value = str(value)
        barcode.barHeight = barHeight
        barcode.barWidth = barWidth
        barcode.humanReadable = humanReadable
        barcode.quiet = quiet

        drawing = Drawing(barcode.width, barcode.height)
        drawing.add(barcode)
        return drawing
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.lib import colors

from flask import Flask, render_template, Response, request, jsonify, redirect, url_for, session
from flask_socketio import SocketIO

MYSQL_MODE = None
try:
    import mysql.connector as mysql_connector  # type: ignore
    MYSQL_MODE = "mysql_connector"
except Exception:
    mysql_connector = None

try:
    import pymysql  # type: ignore
    from pymysql.cursors import DictCursor  # type: ignore
    if MYSQL_MODE is None:
        MYSQL_MODE = "pymysql"
except Exception:
    pymysql = None
    DictCursor = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

# Load both .env and email.env if present
load_dotenv(os.path.join(BASE_DIR, ".env"))
load_dotenv(os.path.join(BASE_DIR, "email.env"))

app = Flask(__name__, template_folder=TEMPLATE_DIR)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "libratrack-secret-key")
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

socketio = SocketIO(app, cors_allowed_origins="*", async_mode="threading")

DB_HOST = os.environ.get("DB_HOST", "127.0.0.1")
DB_PORT = int(os.environ.get("DB_PORT", "3306"))
DB_USER = os.environ.get("DB_USER", "root")
DB_PASSWORD = os.environ.get("DB_PASSWORD", "")
DB_NAME = os.environ.get("DB_NAME", "libratrack_db")

# --------------------------
# SMTP / EMAIL CONFIG
# --------------------------

SMTP_ENV_FILES = [
    os.path.join(BASE_DIR, ".env"),
    os.path.join(BASE_DIR, "email.env"),
]


def reload_smtp_env():
    """Reload email settings so Python uses the latest .env/email.env after restart."""
    for env_file in SMTP_ENV_FILES:
        if os.path.isfile(env_file):
            load_dotenv(env_file, override=True)


def clean_app_password(value: str) -> str:
    """Gmail App Passwords are often copied with spaces. Gmail needs the 16 characters only."""
    return str(value or "").replace(" ", "").strip()


def get_smtp_settings() -> dict:
    """
    Return current SMTP settings.

    Gmail modes:
    - Port 465 uses SMTP_SSL immediately. Do NOT call starttls().
    - Port 587 uses SMTP then STARTTLS.
    """
    reload_smtp_env()

    host = os.getenv("SMTP_HOST", "smtp.gmail.com").strip()

    try:
        port = int(str(os.getenv("SMTP_PORT", "465")).strip())
    except Exception:
        port = 465

    username = os.getenv("SMTP_USERNAME", "cdmlibratrack@gmail.com").strip()
    password = clean_app_password(os.getenv("SMTP_PASSWORD", ""))
    from_email = os.getenv("SMTP_FROM_EMAIL", username or "cdmlibratrack@gmail.com").strip()
    from_name = os.getenv("SMTP_FROM_NAME", "LibraryTracker").strip()

    # Force the correct Gmail TLS mode from the port.
    # This prevents the common mistake: port 465 + STARTTLS = Gmail closes the connection.
    if port == 465:
        use_tls = False
    elif port == 587:
        use_tls = True
    else:
        use_tls = str(os.getenv("SMTP_USE_TLS", "false")).strip().lower() in {"1", "true", "yes", "on"}

    return {
        "host": host,
        "port": port,
        "username": username,
        "password": password,
        "from_email": from_email,
        "from_name": from_name,
        "use_tls": use_tls,
    }


# Backward-compatible globals used by the rest of the app.
_settings = get_smtp_settings()
SMTP_HOST = _settings["host"]
SMTP_PORT = _settings["port"]
SMTP_USERNAME = _settings["username"]
SMTP_PASSWORD = _settings["password"]
SMTP_FROM_EMAIL = _settings["from_email"]
SMTP_FROM_NAME = _settings["from_name"]
SMTP_USE_TLS = _settings["use_tls"]


def refresh_smtp_globals():
    """Keep old global names updated for routes/functions that read them."""
    global SMTP_HOST, SMTP_PORT, SMTP_USERNAME, SMTP_PASSWORD, SMTP_FROM_EMAIL, SMTP_FROM_NAME, SMTP_USE_TLS
    settings = get_smtp_settings()
    SMTP_HOST = settings["host"]
    SMTP_PORT = settings["port"]
    SMTP_USERNAME = settings["username"]
    SMTP_PASSWORD = settings["password"]
    SMTP_FROM_EMAIL = settings["from_email"]
    SMTP_FROM_NAME = settings["from_name"]
    SMTP_USE_TLS = settings["use_tls"]
    return settings


def smtp_configured() -> bool:
    settings = refresh_smtp_globals()
    if not settings["host"]:
        return False
    if settings["port"] not in {465, 587}:
        return False
    if not settings["from_email"]:
        return False
    if not settings["username"]:
        return False
    if not settings["password"]:
        return False
    return True


def send_email_message(message: EmailMessage, success_message: str, fail_prefix: str = "Email failed") -> dict:
    """
    Sends an EmailMessage using Gmail SMTP.

    Correct Gmail modes:
    - Port 465: SMTP_SSL. Do not call starttls().
    - Port 587: SMTP + STARTTLS.
    """
    settings = refresh_smtp_globals()

    if not smtp_configured():
        return {
            "sent": False,
            "message": (
                "SMTP is not configured. Use Gmail SMTP_PORT=465 and SMTP_USE_TLS=false, "
                "then put your NEW Gmail App Password in .env without spaces."
            )
        }

    host = settings["host"]
    port = settings["port"]
    username = settings["username"]
    password = settings["password"]
    from_email = settings["from_email"]
    from_name = settings["from_name"]
    use_tls = settings["use_tls"]

    # Always apply the current From address from .env/email.env.
    if "From" in message:
        del message["From"]
    message["From"] = f"{from_name} <{from_email}>" if from_name else from_email

    try:
        context = ssl.create_default_context()

        print("[SMTP] Host:", host)
        print("[SMTP] Port:", port)
        print("[SMTP] Username:", username)
        print("[SMTP] From:", from_email)
        print("[SMTP] TLS:", use_tls)
        print("[SMTP] Password set:", bool(password))

        if port == 465:
            # Gmail port 465 uses SSL immediately. Never use starttls() here.
            with smtplib.SMTP_SSL(host, port, timeout=60, context=context) as server:
                server.ehlo()
                server.login(username, password)
                server.send_message(message)

        elif port == 587:
            # Gmail port 587 connects first, then upgrades using STARTTLS.
            with smtplib.SMTP(host, port, timeout=60) as server:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
                server.login(username, password)
                server.send_message(message)

        else:
            return {
                "sent": False,
                "message": f"{fail_prefix}: Unsupported SMTP_PORT {port}. Use 465 or 587 for Gmail."
            }

        return {"sent": True, "message": success_message}

    except smtplib.SMTPAuthenticationError:
        return {
            "sent": False,
            "message": (
                f"{fail_prefix}: Gmail rejected the username or app password. "
                "Generate a NEW Gmail App Password, paste it in .env without spaces, "
                "and make sure 2-Step Verification is enabled."
            )
        }

    except smtplib.SMTPServerDisconnected as exc:
        return {
            "sent": False,
            "message": (
                f"{fail_prefix}: Gmail closed the connection. "
                "Use SMTP_PORT=465 with SMTP_USE_TLS=false, make sure Python uses SMTP_SSL, "
                "close the old terminal, then restart Python. "
                f"Current settings: host={host}, port={port}, tls={use_tls}. Details: {exc}"
            )
        }

    except (smtplib.SMTPConnectError, ConnectionRefusedError, OSError) as exc:
        return {
            "sent": False,
            "message": (
                f"{fail_prefix}: Could not connect to Gmail SMTP. "
                "Check internet, firewall/antivirus, school network restrictions, or try mobile hotspot. "
                f"Details: {type(exc).__name__}: {exc}"
            )
        }

    except TimeoutError as exc:
        return {
            "sent": False,
            "message": (
                f"{fail_prefix}: SMTP connection timed out. "
                "Check your internet, firewall, antivirus, or try another network. "
                f"Details: {exc}"
            )
        }

    except Exception as exc:
        return {
            "sent": False,
            "message": f"{fail_prefix}: {type(exc).__name__}: {exc}"
        }


# Default seeded admin account
DEV_ADMIN_USERNAME = os.getenv("DEV_ADMIN_USERNAME", "admin")
DEV_ADMIN_PASSWORD = os.getenv("DEV_ADMIN_PASSWORD", "admin123")

COOLDOWN_SECONDS = 20
barcode_cooldowns = {}
EMAIL_REGEX = re.compile(r"^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$", re.IGNORECASE)
EMAIL_REMINDER_CHECK_SECONDS = int(os.getenv("EMAIL_REMINDER_CHECK_SECONDS", "60"))
reminder_thread_started = False
tracker_process = None
AUTO_START_TRACKER = os.getenv("AUTO_START_TRACKER", "true").lower() in {"1", "true", "yes", "on"}

# FORCE BACKEND TO RUN THE CAMERA-UPDATED TRACKER
TRACKER_SCRIPT_CANDIDATES = [
    os.path.join(BASE_DIR, "face_tracker_hidden.py"),
    os.path.join(BASE_DIR, "Pasted code (2).py"),
]

def resolve_tracker_script():
    env_path = str(os.getenv("TRACKER_SCRIPT", "")).strip()
    candidates = []

    if env_path:
        candidates.append(env_path if os.path.isabs(env_path) else os.path.join(BASE_DIR, env_path))

    candidates.extend(TRACKER_SCRIPT_CANDIDATES)

    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            return candidate

    return TRACKER_SCRIPT_CANDIDATES[0]

SERIAL_ENABLED = os.getenv("SERIAL_ENABLED", "false").lower() in {"1", "true", "yes", "on"}
ESP32_PORT = os.getenv("ESP32_PORT", "COM5")
ESP32_BAUD = int(os.getenv("ESP32_BAUD", "115200"))

serial_conn = None
serial_thread_started = False
serial_lock = threading.Lock()

audio_lock = threading.Lock()
tts_engine = None
audio_busy = False

TIME_IN_AUDIO_ENABLED = True
TIME_IN_AUDIO_MESSAGE = "Welcome to the library"
TIME_OUT_AUDIO_ENABLED = False
TIME_OUT_AUDIO_MESSAGE = "Goodbye"

# Set this to False to avoid Windows asyncio RuntimeError:
# RuntimeError: Event loop is closed
PREFER_EDGE_TTS = False
EDGE_TTS_VOICE = "en-US-GuyNeural"
EDGE_TTS_RATE = "-5%"
EDGE_TTS_PITCH = "-1Hz"
EDGE_TTS_VOLUME = "+0%"
FACE_AUDIO_RATE = 150

def sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()

def now_dt() -> datetime:
    return datetime.now()

def format_datetime(dt_value):
    if not dt_value:
        return "--"
    if isinstance(dt_value, str):
        return dt_value
    return dt_value.strftime("%Y-%m-%d %I:%M:%S %p")

def format_date(dt_value):
    if not dt_value:
        return "--"
    if isinstance(dt_value, str):
        return dt_value
    return dt_value.strftime("%Y-%m-%d")

def format_time(dt_value):
    if not dt_value:
        return ""
    if isinstance(dt_value, str):
        return dt_value
    return dt_value.strftime("%I:%M:%S %p")

def require_db_driver():
    if MYSQL_MODE is None:
        raise RuntimeError("No MySQL Python driver installed. Install mysql-connector-python or pymysql.")

def get_server_connection():
    require_db_driver()
    if MYSQL_MODE == "mysql_connector":
        return mysql_connector.connect(
            host=DB_HOST,
            port=DB_PORT,
            user=DB_USER,
            password=DB_PASSWORD,
            autocommit=False,
        )
    return pymysql.connect(
        host=DB_HOST,
        port=DB_PORT,
        user=DB_USER,
        password=DB_PASSWORD,
        autocommit=False,
    )

def get_db_connection():
    require_db_driver()
    if MYSQL_MODE == "mysql_connector":
        return mysql_connector.connect(
            host=DB_HOST,
            port=DB_PORT,
            user=DB_USER,
            password=DB_PASSWORD,
            database=DB_NAME,
            autocommit=False,
        )
    return pymysql.connect(
        host=DB_HOST,
        port=DB_PORT,
        user=DB_USER,
        password=DB_PASSWORD,
        database=DB_NAME,
        cursorclass=DictCursor,
        autocommit=False,
    )

def dict_cursor(conn):
    if MYSQL_MODE == "mysql_connector":
        return conn.cursor(dictionary=True)
    return conn.cursor()

def ensure_database_exists():
    conn = None
    cursor = None
    try:
        conn = get_server_connection()
        cursor = conn.cursor()
        cursor.execute(
            f"CREATE DATABASE IF NOT EXISTS `{DB_NAME}` "
            "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
        )
        conn.commit()
        print(f"[DB] Database ensured: {DB_NAME}")
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def ensure_users_table_and_admin():
    conn = None
    cursor = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                username VARCHAR(100) NOT NULL UNIQUE,
                full_name VARCHAR(150) DEFAULT NULL,
                password_hash VARCHAR(255) NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        admin_username = DEV_ADMIN_USERNAME.strip()
        admin_password_hash = sha256_text(DEV_ADMIN_PASSWORD)

        cursor.execute(
            "SELECT id, username, password_hash FROM users WHERE username = %s LIMIT 1",
            (admin_username,)
        )
        user = cursor.fetchone()

        if user:
            if user["password_hash"] != admin_password_hash:
                cursor.execute(
                    """
                    UPDATE users
                    SET password_hash = %s,
                        full_name = %s
                    WHERE username = %s
                    """,
                    (admin_password_hash, "Administrator", admin_username)
                )
                print(f"[DB] Updated admin account: {admin_username}")
        else:
            cursor.execute(
                """
                INSERT INTO users (username, full_name, password_hash)
                VALUES (%s, %s, %s)
                """,
                (admin_username, "Administrator", admin_password_hash)
            )
            print(f"[DB] Seeded admin account: {admin_username}")

        conn.commit()
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def ensure_students_table():
    conn = None
    cursor = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS students (
                student_id VARCHAR(50) NOT NULL PRIMARY KEY,
                barcode_value VARCHAR(100) NOT NULL UNIQUE,
                full_name VARCHAR(150) NOT NULL,
                email VARCHAR(150) NOT NULL UNIQUE,
                age INT NOT NULL,
                year_level VARCHAR(50) DEFAULT '',
                course VARCHAR(100) DEFAULT '',
                address VARCHAR(255) DEFAULT '',
                last_attendance DATETIME NULL,
                status VARCHAR(50) DEFAULT 'Registered',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.commit()
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def ensure_attendance_table():
    conn = None
    cursor = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS attendance (
                id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                student_id VARCHAR(50) NOT NULL,
                attendance_date DATE NOT NULL,
                time_in DATETIME NOT NULL,
                time_out DATETIME NULL,
                last_action VARCHAR(50) DEFAULT 'TIME IN',
                status VARCHAR(50) DEFAULT 'OPEN',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                INDEX idx_attendance_student_day (student_id, attendance_date),
                CONSTRAINT fk_attendance_student
                    FOREIGN KEY (student_id) REFERENCES students(student_id)
                    ON DELETE CASCADE
            )
            """
        )
        conn.commit()
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def ensure_books_table():
    conn = None
    cursor = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS books (
                id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                book_id VARCHAR(50) NOT NULL UNIQUE,
                barcode_value VARCHAR(100) NOT NULL UNIQUE,
                title VARCHAR(255) NOT NULL,
                author VARCHAR(255) DEFAULT '',
                category VARCHAR(100) DEFAULT '',
                status VARCHAR(50) DEFAULT 'AVAILABLE',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.commit()
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def ensure_teachers_table():
    conn = None
    cursor = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS teachers (
                teacher_id VARCHAR(50) NOT NULL PRIMARY KEY,
                barcode_value VARCHAR(100) NOT NULL UNIQUE,
                full_name VARCHAR(150) NOT NULL,
                email VARCHAR(150) NOT NULL UNIQUE,
                department VARCHAR(100) DEFAULT '',
                address VARCHAR(255) DEFAULT '',
                status VARCHAR(50) DEFAULT 'Active',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.commit()
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def table_exists(cursor, table_name: str) -> bool:
    cursor.execute("SHOW TABLES LIKE %s", (table_name,))
    return bool(cursor.fetchone())

def get_table_columns(cursor, table_name: str):
    cursor.execute(f"SHOW COLUMNS FROM {table_name}")
    rows = cursor.fetchall()
    columns = {}
    for row in rows:
        if isinstance(row, dict):
            columns[row["Field"]] = row
        else:
            columns[row[0]] = {
                "Field": row[0],
                "Type": row[1],
                "Null": row[2],
                "Key": row[3],
                "Default": row[4],
                "Extra": row[5],
            }
    return columns

def ensure_borrow_transactions_table():
    conn = None
    cursor = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)

        if not table_exists(cursor, "borrow_transactions"):
            cursor.execute(
                """
                CREATE TABLE borrow_transactions (
                    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                    borrower_type VARCHAR(20) NOT NULL DEFAULT 'STUDENT',
                    student_id VARCHAR(50) NULL,
                    teacher_id VARCHAR(50) NULL,
                    book_id INT NOT NULL,
                    borrow_date DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    due_date DATETIME NOT NULL,
                    return_date DATETIME NULL,
                    status VARCHAR(50) DEFAULT 'BORROWED',
                    reminder_sent TINYINT(1) DEFAULT 0,
                    overdue_email_sent TINYINT(1) DEFAULT 0,
                    overdue_email_sent_at DATETIME NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    INDEX idx_borrow_book_status (book_id, status),
                    INDEX idx_borrow_student_status (student_id, status),
                    INDEX idx_borrow_teacher_status (teacher_id, status),
                    INDEX idx_borrow_due_date (due_date),
                    INDEX idx_borrower_type (borrower_type),
                    CONSTRAINT fk_borrow_student
                        FOREIGN KEY (student_id) REFERENCES students(student_id)
                        ON DELETE CASCADE,
                    CONSTRAINT fk_borrow_teacher
                        FOREIGN KEY (teacher_id) REFERENCES teachers(teacher_id)
                        ON DELETE CASCADE,
                    CONSTRAINT fk_borrow_book
                        FOREIGN KEY (book_id) REFERENCES books(id)
                        ON DELETE CASCADE
                )
                """
            )
            conn.commit()
            return

        columns = get_table_columns(cursor, "borrow_transactions")
        needs_migration = (
            "borrower_type" not in columns
            or "teacher_id" not in columns
            or columns.get("student_id", {}).get("Null", "").upper() != "YES"
        )

        if needs_migration:
            cursor.execute("DROP TABLE IF EXISTS borrow_transactions_new")
            cursor.execute(
                """
                CREATE TABLE borrow_transactions_new (
                    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                    borrower_type VARCHAR(20) NOT NULL DEFAULT 'STUDENT',
                    student_id VARCHAR(50) NULL,
                    teacher_id VARCHAR(50) NULL,
                    book_id INT NOT NULL,
                    borrow_date DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    due_date DATETIME NOT NULL,
                    return_date DATETIME NULL,
                    status VARCHAR(50) DEFAULT 'BORROWED',
                    reminder_sent TINYINT(1) DEFAULT 0,
                    overdue_email_sent TINYINT(1) DEFAULT 0,
                    overdue_email_sent_at DATETIME NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    INDEX idx_borrow_book_status (book_id, status),
                    INDEX idx_borrow_student_status (student_id, status),
                    INDEX idx_borrow_teacher_status (teacher_id, status),
                    INDEX idx_borrow_due_date (due_date),
                    INDEX idx_borrower_type (borrower_type),
                    CONSTRAINT fk_borrow_student_new
                        FOREIGN KEY (student_id) REFERENCES students(student_id)
                        ON DELETE CASCADE,
                    CONSTRAINT fk_borrow_teacher_new
                        FOREIGN KEY (teacher_id) REFERENCES teachers(teacher_id)
                        ON DELETE CASCADE,
                    CONSTRAINT fk_borrow_book_new
                        FOREIGN KEY (book_id) REFERENCES books(id)
                        ON DELETE CASCADE
                )
                """
            )

            select_parts = [
                "id",
                "'STUDENT' AS borrower_type",
                "student_id",
                "NULL AS teacher_id",
                "book_id",
                "borrow_date",
                "due_date",
                "return_date",
                "status",
                "COALESCE(reminder_sent, 0) AS reminder_sent",
                "0 AS overdue_email_sent",
                "NULL AS overdue_email_sent_at",
                "created_at",
                "CURRENT_TIMESTAMP AS updated_at",
            ]
            if "overdue_email_sent" in columns:
                select_parts[10] = "COALESCE(overdue_email_sent, 0) AS overdue_email_sent"
            if "overdue_email_sent_at" in columns:
                select_parts[11] = "overdue_email_sent_at"
            if "updated_at" in columns:
                select_parts[13] = "updated_at"

            cursor.execute(
                f"""
                INSERT INTO borrow_transactions_new (
                    id, borrower_type, student_id, teacher_id, book_id,
                    borrow_date, due_date, return_date, status,
                    reminder_sent, overdue_email_sent, overdue_email_sent_at,
                    created_at, updated_at
                )
                SELECT {", ".join(select_parts)}
                FROM borrow_transactions
                """
            )

            cursor.execute("SET FOREIGN_KEY_CHECKS = 0")
            cursor.execute("DROP TABLE borrow_transactions")
            cursor.execute("RENAME TABLE borrow_transactions_new TO borrow_transactions")
            cursor.execute("SET FOREIGN_KEY_CHECKS = 1")
            conn.commit()
            return

        # forward-compatible columns/indexes
        if "borrower_type" not in columns:
            cursor.execute("ALTER TABLE borrow_transactions ADD COLUMN borrower_type VARCHAR(20) NOT NULL DEFAULT 'STUDENT' AFTER id")
        if "teacher_id" not in columns:
            cursor.execute("ALTER TABLE borrow_transactions ADD COLUMN teacher_id VARCHAR(50) NULL AFTER student_id")
        if "overdue_email_sent" not in columns:
            cursor.execute("ALTER TABLE borrow_transactions ADD COLUMN overdue_email_sent TINYINT(1) DEFAULT 0 AFTER reminder_sent")
        if "overdue_email_sent_at" not in columns:
            cursor.execute("ALTER TABLE borrow_transactions ADD COLUMN overdue_email_sent_at DATETIME NULL AFTER overdue_email_sent")

        cursor.execute("SHOW INDEX FROM borrow_transactions")
        index_rows = cursor.fetchall()
        existing_indexes = set(row["Key_name"] if isinstance(row, dict) else row[2] for row in index_rows)

        if "idx_borrow_teacher_status" not in existing_indexes:
            cursor.execute("CREATE INDEX idx_borrow_teacher_status ON borrow_transactions (teacher_id, status)")
        if "idx_borrower_type" not in existing_indexes:
            cursor.execute("CREATE INDEX idx_borrower_type ON borrow_transactions (borrower_type)")

        conn.commit()
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def initialize_database():
    ensure_database_exists()
    ensure_users_table_and_admin()
    ensure_students_table()
    ensure_attendance_table()
    ensure_books_table()
    ensure_teachers_table()
    ensure_borrow_transactions_table()

def emit_status(message):
    socketio.emit("systemStatus", {"message": message})
    print(message)

def open_serial_connection():
    global serial_conn

    if not SERIAL_ENABLED:
        print("[SERIAL] Serial sync disabled.")
        return None

    if serial_conn and getattr(serial_conn, "is_open", False):
        return serial_conn

    try:
        serial_conn = serial.Serial(ESP32_PORT, ESP32_BAUD, timeout=1)
        time.sleep(2)
        print(f"[SERIAL] Connected to ESP32 on {ESP32_PORT} @ {ESP32_BAUD}")
        return serial_conn
    except Exception as exc:
        serial_conn = None
        print(f"[SERIAL ERROR] Could not open {ESP32_PORT}: {exc}")
        return None

def close_serial_connection():
    global serial_conn
    try:
        if serial_conn and getattr(serial_conn, "is_open", False):
            serial_conn.close()
            print("[SERIAL] Connection closed.")
    except Exception:
        pass
    serial_conn = None

def send_led_command(command: str):
    conn = open_serial_connection()
    if not conn:
        return False

    try:
        with serial_lock:
            conn.write((command.strip() + "\n").encode("utf-8"))
            conn.flush()
        print(f"[SERIAL] SENT -> {command}")
        return True
    except Exception as exc:
        print(f"[SERIAL ERROR] Failed to send '{command}': {exc}")
        close_serial_connection()
        return False

def init_tts_engine():
    global tts_engine
    if tts_engine is not None:
        return tts_engine

    if not PYTTSX3_AVAILABLE:
        return None

    tts_engine = pyttsx3.init()
    tts_engine.setProperty("rate", FACE_AUDIO_RATE)

    try:
        voices = tts_engine.getProperty("voices")
        preferred = None
        for voice in voices:
            name = (getattr(voice, "name", "") or "").lower()
            vid = (getattr(voice, "id", "") or "").lower()
            if any(keyword in name or keyword in vid for keyword in ["david", "mark", "guy", "male"]):
                preferred = voice.id
                break
        if preferred:
            tts_engine.setProperty("voice", preferred)
    except Exception:
        pass

    return tts_engine


def init_hidden_audio_player():
    if not PYGAME_AVAILABLE:
        return False
    try:
        if not pygame.mixer.get_init():
            pygame.mixer.pre_init(frequency=24000, size=-16, channels=1, buffer=4096)
            pygame.mixer.init()
        return True
    except Exception:
        return False


async def save_edge_tts_to_file(text, output_file):
    communicate = edge_tts.Communicate(
        text=text,
        voice=EDGE_TTS_VOICE,
        rate=EDGE_TTS_RATE,
        pitch=EDGE_TTS_PITCH,
        volume=EDGE_TTS_VOLUME,
    )
    await communicate.save(output_file)


def speak_with_edge_tts(text):
    if not EDGE_TTS_AVAILABLE:
        return False

    if not init_hidden_audio_player():
        return False

    temp_file = os.path.join(tempfile.gettempdir(), "backend_voice_output.mp3")

    try:
        if pygame.mixer.music.get_busy():
            pygame.mixer.music.stop()
            time.sleep(0.03)

        try:
            if hasattr(pygame.mixer.music, "unload"):
                pygame.mixer.music.unload()
        except Exception:
            pass

        asyncio.run(save_edge_tts_to_file(text, temp_file))

        pygame.mixer.music.load(temp_file)
        pygame.mixer.music.set_volume(1.0)
        pygame.mixer.music.play()

        while pygame.mixer.music.get_busy():
            time.sleep(0.02)

        try:
            pygame.mixer.music.stop()
        except Exception:
            pass

        try:
            if hasattr(pygame.mixer.music, "unload"):
                pygame.mixer.music.unload()
        except Exception:
            pass

        try:
            if os.path.exists(temp_file):
                os.remove(temp_file)
        except Exception:
            pass

        return True

    except Exception as exc:
        print(f"[AUDIO ERROR] EDGE TTS failed: {exc}")
        return False


def speak_with_pyttsx3(text):
    try:
        engine = init_tts_engine()
        if engine is None:
            return False

        try:
            engine.stop()
        except Exception:
            pass

        engine.say(text)
        engine.runAndWait()
        return True

    except Exception as exc:
        print(f"[AUDIO ERROR] pyttsx3 failed: {exc}")
        return False


def speak_text(text):
    global audio_busy

    if not text:
        return False

    with audio_lock:
        if audio_busy:
            return False
        audio_busy = True

    try:
        spoken = False

        if PREFER_EDGE_TTS and EDGE_TTS_AVAILABLE:
            spoken = speak_with_edge_tts(text)

        if not spoken:
            spoken = speak_with_pyttsx3(text)

        return spoken

    finally:
        with audio_lock:
            audio_busy = False


def speak_async(text: str):
    if not text:
        return

    def runner():
        try:
            speak_text(text)
        except Exception as exc:
            print(f"[AUDIO ERROR] {exc}")

    threading.Thread(target=runner, daemon=True).start()


def handle_serial_scan(barcode: str):
    payload, status_code = process_barcode_scan(barcode, source="serial")

    if status_code == 200:
        action = str(payload.get("action", "")).upper()

        if action == "TIME IN":
            send_led_command("LED:GREEN")
            if TIME_IN_AUDIO_ENABLED:
                speak_async(TIME_IN_AUDIO_MESSAGE)
        elif action == "TIME OUT":
            send_led_command("LED:RED")
            if TIME_OUT_AUDIO_ENABLED:
                speak_async(TIME_OUT_AUDIO_MESSAGE)
        else:
            send_led_command("LED:ERROR")

        socketio.emit("attendanceUpdated", payload)
        socketio.emit("attendanceTableChanged", {
            "ok": True,
            "message": payload.get("message", "Attendance updated."),
            "record": payload.get("record"),
            "attendance": payload.get("attendance"),
            "student": payload.get("student"),
        })
    else:
        send_led_command("LED:ERROR")
        socketio.emit("barcodeRejected", payload)

    return payload, status_code

def serial_listener_loop():
    emit_status(f"[SERIAL] Listening for ESP32 scans on {ESP32_PORT}...")

    while True:
        conn = open_serial_connection()
        if not conn:
            time.sleep(3)
            continue

        try:
            if conn.in_waiting:
                raw = conn.readline().decode("utf-8", errors="ignore").strip()
                if not raw:
                    continue

                print(f"[ESP32] {raw}")

                if raw.startswith("SCAN:"):
                    barcode = raw[5:].strip()
                    if barcode:
                        handle_serial_scan(barcode)
            else:
                time.sleep(0.02)

        except Exception as exc:
            print(f"[SERIAL ERROR] Listener crashed: {exc}")
            close_serial_connection()
            time.sleep(2)

def start_serial_listener_thread():
    global serial_thread_started
    if serial_thread_started:
        return

    if not SERIAL_ENABLED:
        print("[SERIAL] Listener disabled.")
        return

    thread = threading.Thread(target=serial_listener_loop, daemon=True)
    thread.start()
    serial_thread_started = True

def normalize_email(email_value: str) -> str:
    return str(email_value or "").strip().lower()

def is_valid_email(email_value: str) -> bool:
    return bool(EMAIL_REGEX.match(normalize_email(email_value)))


def is_logged_in() -> bool:
    return bool(session.get("username"))

def set_logged_in_user(user_id, username, full_name=None):
    session["user_id"] = user_id
    session["username"] = username
    session["full_name"] = full_name or username

def build_barcode_drawing(barcode_value: str):
    barcode_value = str(barcode_value or "").strip()
    if not barcode_value:
        raise ValueError("Barcode value is required.")

    barcode = createBarcodeDrawing(
        "Code128",
        value=barcode_value,
        barHeight=72,
        barWidth=0.42,
        humanReadable=False,
        quiet=True,
    )

    margin = 18
    label_height = 20
    total_width = barcode.width + (margin * 2)
    total_height = barcode.height + (margin * 2) + label_height

    drawing = Drawing(total_width, total_height)
    drawing.add(Rect(0, 0, total_width, total_height, fillColor=colors.white, strokeColor=colors.white))

    barcode_group = deepcopy(barcode)
    barcode_group.translate(margin, margin + label_height)
    drawing.add(barcode_group)
    drawing.add(
        String(
            total_width / 2,
            8,
            barcode_value,
            textAnchor="middle",
            fontName="Helvetica-Bold",
            fontSize=10,
            fillColor=colors.black,
        )
    )
    return drawing

def generate_barcode_png(barcode_value: str) -> bytes:
    drawing = build_barcode_drawing(barcode_value)
    return renderPM.drawToString(drawing, fmt="PNG", dpi=300, bg=0xFFFFFF)

def generate_barcode_svg(barcode_value: str) -> bytes:
    drawing = build_barcode_drawing(barcode_value)
    return renderSVG.drawToString(drawing).encode("utf-8")


def send_person_barcode_email(person_row, role_label: str, id_field: str) -> dict:
    email_address = normalize_email((person_row or {}).get("email") or "")
    barcode_value = str((person_row or {}).get("barcode_value") or (person_row or {}).get(id_field) or "").strip()
    person_name = str((person_row or {}).get("full_name") or role_label).strip()
    person_id = str((person_row or {}).get(id_field) or "").strip()

    if not email_address:
        return {"sent": False, "message": f"{role_label} email is blank. Barcode email was not sent."}

    if not is_valid_email(email_address):
        return {"sent": False, "message": f"{role_label} email format is invalid. Please save a real email address first."}

    if not smtp_configured():
        return {"sent": False, "message": f"SMTP is not configured yet. {role_label} was saved, but barcode email was not sent."}

    barcode_png = generate_barcode_png(barcode_value)
    barcode_svg = generate_barcode_svg(barcode_value)

    message = EmailMessage()
    message["Subject"] = f"Your Libratrack Barcode - {barcode_value}"
    message["From"] = f"{SMTP_FROM_NAME} <{SMTP_FROM_EMAIL}>" if SMTP_FROM_NAME else SMTP_FROM_EMAIL
    message["To"] = email_address

    message.set_content(
        f"Hello {person_name},\n\n"
        f"Your Libratrack barcode is attached in two formats for easier scanning.\n\n"
        f"{role_label} ID: {person_id}\n"
        f"Barcode: {barcode_value}\n\n"
        "Recommended use:\n"
        "- Use the SVG file for printing\n"
        "- Use the PNG file for phone display\n"
        "- Do not crop the white border\n"
        "- Set your phone brightness high when scanning\n\n"
        "Please keep this email as your copy.\n\n"
        "Libratrack Attendance System"
    )

    message.add_alternative(
        f"""
        <html>
          <body style="font-family:Segoe UI,Arial,sans-serif;color:#2c3e50;line-height:1.6;background:#f7f9fc;padding:24px;">
            <div style="max-width:640px;margin:auto;background:#ffffff;border:1px solid #e3e8ef;border-radius:14px;padding:28px;">
              <h2 style="margin-top:0;color:#1f2937;">Your Libratrack Barcode</h2>
              <p>Hello {person_name},</p>
              <p>Your {role_label.lower()} barcode for <strong>Libratrack</strong> is attached in <strong>PNG</strong> and <strong>SVG</strong> formats for clearer scanning.</p>

              <div style="background:#f8fafc;border:1px solid #dbe4ee;border-radius:10px;padding:16px;margin:18px 0;">
                <p style="margin:0 0 8px 0;"><strong>{role_label} ID:</strong> {person_id}</p>
                <p style="margin:0;"><strong>Barcode:</strong> {barcode_value}</p>
              </div>

              <p style="margin-bottom:8px;"><strong>For best scanning:</strong></p>
              <ul style="margin-top:0;padding-left:20px;">
                <li>Use the SVG file for printing</li>
                <li>Use the PNG file on phones</li>
                <li>Do not crop the barcode image</li>
                <li>Keep the white border visible</li>
                <li>Increase screen brightness if scanning from phone</li>
              </ul>

              <p>Please keep this email as your copy of the barcode.</p>
              <p style="margin-top:24px;">Libratrack Attendance System</p>
            </div>
          </body>
        </html>
        """,
        subtype="html",
    )

    message.add_attachment(barcode_png, maintype="image", subtype="png", filename=f"barcode_{barcode_value}.png")
    message.add_attachment(barcode_svg, maintype="image", subtype="svg+xml", filename=f"barcode_{barcode_value}.svg")

    return send_email_message(
        message,
        success_message=f"Barcode email sent to {email_address}.",
        fail_prefix=f"{role_label} saved, but barcode email failed"
    )

def send_barcode_email(student_row) -> dict:
    return send_person_barcode_email(student_row, "Student", "student_id")

def send_teacher_barcode_email(teacher_row) -> dict:
    return send_person_barcode_email(teacher_row, "Teacher", "teacher_id")

def send_plain_email(to_email: str, subject: str, body: str) -> dict:
    email_address = normalize_email(to_email)
    if not email_address:
        return {"sent": False, "message": "Recipient email is blank."}
    if not is_valid_email(email_address):
        return {"sent": False, "message": "Recipient email format is invalid."}
    if not smtp_configured():
        return {"sent": False, "message": "SMTP is not configured yet."}

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = f"{SMTP_FROM_NAME} <{SMTP_FROM_EMAIL}>" if SMTP_FROM_NAME else SMTP_FROM_EMAIL
    message["To"] = email_address
    message.set_content(body)

    return send_email_message(
        message,
        success_message=f"Email sent to {email_address}.",
        fail_prefix="Email failed"
    )


def send_book_borrowed_email(borrower_row, book_row, due_date: datetime, borrower_type: str = "STUDENT") -> dict:
    email_address = normalize_email((borrower_row or {}).get("email") or "")
    borrower_type = str(borrower_type or "STUDENT").upper()
    borrower_label = "Teacher" if borrower_type == "TEACHER" else "Student"
    borrower_id_field = "teacher_id" if borrower_type == "TEACHER" else "student_id"
    borrower_name = str((borrower_row or {}).get("full_name") or borrower_label).strip()
    borrower_id = str((borrower_row or {}).get(borrower_id_field) or "").strip()
    book_title = str((book_row or {}).get("title") or "Book").strip()
    book_code = str((book_row or {}).get("book_id") or (book_row or {}).get("barcode_value") or "").strip()
    deadline_text = format_datetime(due_date)

    if not email_address:
        return {"sent": False, "message": f"{borrower_label} email is blank. Borrow confirmation email was not sent."}

    if not is_valid_email(email_address):
        return {"sent": False, "message": f"{borrower_label} email format is invalid. Borrow confirmation email was not sent."}

    subject = f"Library Borrowed Book Confirmation: {book_title}"
    body = (
        f"Hello {borrower_name},\n\n"
        "This is a confirmation that you borrowed a book from the library.\n\n"
        f"{borrower_label} ID: {borrower_id}\n"
        f"Book Title: {book_title}\n"
        f"Book Code: {book_code}\n"
        f"Deadline: {deadline_text}\n\n"
        "Please return the book on or before the deadline to avoid an overdue reminder.\n\n"
        "Thank you.\n"
        "Libratrack Attendance System"
    )
    return send_plain_email(email_address, subject, body)

def parse_due_date(due_value: str) -> datetime:
    raw = str(due_value or "").strip()
    if not raw:
        raise ValueError("Due date is required.")

    supported_formats = [
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%dT%H:%M:%S",
    ]

    for fmt in supported_formats:
        try:
            parsed = datetime.strptime(raw, fmt)
            if fmt == "%Y-%m-%d":
                return parsed.replace(hour=23, minute=59, second=59)
            return parsed
        except ValueError:
            continue

    try:
        parsed = datetime.fromisoformat(raw)
        if parsed.hour == 0 and parsed.minute == 0 and parsed.second == 0 and "T" not in raw and " " not in raw:
            parsed = parsed.replace(hour=23, minute=59, second=59)
        return parsed
    except ValueError:
        pass

    raise ValueError("Due date must be YYYY-MM-DD or a valid datetime string.")

def serialize_student(row):
    if not row:
        return None
    student_id = row["student_id"]
    barcode_value = row.get("barcode_value") or student_id
    full_name = row["full_name"]
    has_overdue_books = bool(row.get("has_overdue_books"))
    return {
        "id": student_id,
        "studentId": student_id,
        "barcode": barcode_value,
        "name": full_name,
        "fullName": full_name,
        "email": row.get("email") or "",
        "age": row["age"],
        "year": row.get("year_level") or "",
        "course": row.get("course") or "",
        "address": row.get("address") or "",
        "lastAttendance": format_datetime(row.get("last_attendance")),
        "status": row.get("status") or "Registered",
        "hasOverdueBooks": has_overdue_books,
        "has_overdue_books": has_overdue_books,
        "createdAt": format_datetime(row.get("created_at")),
    }


def serialize_teacher(row):
    if not row:
        return None
    teacher_id = row["teacher_id"]
    barcode_value = row.get("barcode_value") or teacher_id
    full_name = row["full_name"]
    return {
        "id": teacher_id,
        "teacherId": teacher_id,
        "barcode": barcode_value,
        "name": full_name,
        "fullName": full_name,
        "email": row.get("email") or "",
        "department": row.get("department") or "",
        "address": row.get("address") or "",
        "status": row.get("status") or "Active",
        "createdAt": format_datetime(row.get("created_at")),
    }

def serialize_attendance(row):
    if not row:
        return None
    student_id = row["student_id"]
    barcode_value = row.get("barcode_value") or student_id
    full_name = row["full_name"]
    return {
        "id": row["id"],
        "studentId": student_id,
        "student_id": student_id,
        "barcode": barcode_value,
        "name": full_name,
        "fullName": full_name,
        "date": format_date(row.get("attendance_date")),
        "timeIn": format_time(row.get("time_in")) or "--",
        "timeOut": format_time(row.get("time_out")) or "--",
        "lastAction": row.get("last_action") or "--",
        "status": row.get("status") or "--",
        "createdAt": format_datetime(row.get("created_at")),
    }

def next_student_id(conn):
    current_year = datetime.now().year
    cursor = dict_cursor(conn)
    try:
        cursor.execute(
            """
            SELECT student_id
            FROM students
            WHERE student_id LIKE %s
            ORDER BY student_id DESC
            LIMIT 1
            """,
            (f"{current_year}-%",)
        )
        row = cursor.fetchone()
    finally:
        cursor.close()

    if row:
        current = row["student_id"] if isinstance(row, dict) else row[0]
        try:
            seq = int(str(current).split("-")[1]) + 1
        except Exception:
            seq = 1
    else:
        seq = 1

    return f"{current_year}-{seq:04d}"


def next_teacher_id(conn):
    cursor = dict_cursor(conn)
    try:
        cursor.execute(
            """
            SELECT teacher_id
            FROM teachers
            WHERE teacher_id LIKE 'TCH-%'
            ORDER BY teacher_id DESC
            LIMIT 1
            """
        )
        row = cursor.fetchone()
    finally:
        cursor.close()

    if row:
        current = row["teacher_id"] if isinstance(row, dict) else row[0]
        try:
            seq = int(str(current).split("-")[1]) + 1
        except Exception:
            seq = 1
    else:
        seq = 1

    return f"TCH-{seq:04d}"

def fetch_teacher_by_barcode(conn, barcode):
    cursor = dict_cursor(conn)
    try:
        cursor.execute(
            """
            SELECT * FROM teachers
            WHERE teacher_id = %s OR barcode_value = %s
            LIMIT 1
            """,
            (barcode, barcode)
        )
        return cursor.fetchone()
    finally:
        cursor.close()

def fetch_teacher_by_id(conn, teacher_id):
    cursor = dict_cursor(conn)
    try:
        cursor.execute("SELECT * FROM teachers WHERE teacher_id = %s LIMIT 1", (teacher_id,))
        return cursor.fetchone()
    finally:
        cursor.close()

def fetch_student_by_barcode(conn, barcode):
    cursor = dict_cursor(conn)
    try:
        cursor.execute(
            """
            SELECT * FROM students
            WHERE student_id = %s OR barcode_value = %s
            LIMIT 1
            """,
            (barcode, barcode)
        )
        return cursor.fetchone()
    finally:
        cursor.close()

def fetch_student_by_id(conn, student_id):
    cursor = dict_cursor(conn)
    try:
        cursor.execute("SELECT * FROM students WHERE student_id = %s LIMIT 1", (student_id,))
        return cursor.fetchone()
    finally:
        cursor.close()

def fetch_latest_open_record(conn, student_id, attendance_day):
    cursor = dict_cursor(conn)
    try:
        cursor.execute(
            """
            SELECT a.id, a.student_id, s.barcode_value, s.full_name, a.attendance_date, a.time_in, a.time_out,
                   a.last_action, a.status, a.created_at
            FROM attendance a
            INNER JOIN students s ON s.student_id = a.student_id
            WHERE a.student_id = %s
              AND a.attendance_date = %s
              AND a.time_out IS NULL
            ORDER BY a.id DESC
            LIMIT 1
            """,
            (student_id, attendance_day)
        )
        return cursor.fetchone()
    finally:
        cursor.close()

def fetch_attendance_by_id(conn, attendance_id):
    cursor = dict_cursor(conn)
    try:
        cursor.execute(
            """
            SELECT a.id, a.student_id, s.barcode_value, s.full_name, a.attendance_date, a.time_in, a.time_out,
                   a.last_action, a.status, a.created_at
            FROM attendance a
            INNER JOIN students s ON s.student_id = a.student_id
            WHERE a.id = %s
            LIMIT 1
            """,
            (attendance_id,)
        )
        return cursor.fetchone()
    finally:
        cursor.close()

def process_barcode_scan(barcode: str, source: str = "api"):
    barcode = str(barcode or "").strip()
    if not barcode:
        return {
            "ok": False,
            "message": "Barcode or Student ID is required.",
            "reason": "missing_barcode",
            "barcode": barcode,
            "source": source,
        }, 400

    current_time = time.time()
    last_scan = barcode_cooldowns.get(barcode)
    if last_scan and (current_time - last_scan) < COOLDOWN_SECONDS:
        remaining = max(1, int(COOLDOWN_SECONDS - (current_time - last_scan)))
        return {
            "ok": False,
            "message": f"Barcode or Student ID {barcode} is on cooldown. Wait {remaining}s.",
            "reason": "cooldown",
            "barcode": barcode,
            "remainingSeconds": remaining,
            "source": source,
        }, 429

    conn = None
    try:
        conn = get_db_connection()
        student = fetch_student_by_barcode(conn, barcode)
        if not student:
            return {
                "ok": False,
                "message": f"Student not found for barcode or student ID: {barcode}",
                "reason": "not_found",
                "barcode": barcode,
                "source": source,
            }, 404

        barcode_cooldowns[barcode] = current_time
        scan_time = now_dt()
        attendance_day = scan_time.date()
        open_record = fetch_latest_open_record(conn, student["student_id"], attendance_day)

        cursor = dict_cursor(conn)
        try:
            if open_record:
                cursor.execute(
                    """
                    UPDATE attendance
                    SET time_out = %s,
                        last_action = 'TIME OUT',
                        status = 'COMPLETED'
                    WHERE id = %s
                    """,
                    (scan_time, open_record["id"])
                )
                attendance_id = open_record["id"]
                action = "TIME OUT"
            else:
                cursor.execute(
                    """
                    INSERT INTO attendance (student_id, attendance_date, time_in, time_out, last_action, status)
                    VALUES (%s, %s, %s, NULL, 'TIME IN', 'OPEN')
                    """,
                    (student["student_id"], attendance_day, scan_time)
                )
                attendance_id = cursor.lastrowid
                action = "TIME IN"

            cursor.execute(
                """
                UPDATE students
                SET last_attendance = %s,
                    status = %s
                WHERE student_id = %s
                """,
                (scan_time, action, student["student_id"])
            )
            conn.commit()
        finally:
            cursor.close()

        updated_student = fetch_student_by_id(conn, student["student_id"])
        updated_attendance = fetch_attendance_by_id(conn, attendance_id)

        serialized_student = serialize_student(updated_student)
        serialized_attendance = serialize_attendance(updated_attendance)

        borrow_cursor = dict_cursor(conn)
        try:
            unreturned_books = get_student_unreturned_books(borrow_cursor, student["student_id"])
        finally:
            borrow_cursor.close()

        has_unreturned_books = len(unreturned_books) > 0

        return {
            "ok": True,
            "message": f"{action} recorded for {updated_student['full_name']}",
            "source": source,
            "barcode": barcode,
            "action": action,
            "scannedAt": format_datetime(scan_time),
            "student": serialized_student,
            "record": serialized_attendance,
            "attendance": serialized_attendance,
            "unreturnedBooks": unreturned_books,
            "unreturned_books": unreturned_books,
            "hasUnreturnedBooks": has_unreturned_books,
            "has_unreturned_books": has_unreturned_books,
        }, 200
    except Exception as exc:
        if conn:
            conn.rollback()
        return {
            "ok": False,
            "message": f"Database error: {exc}",
            "reason": "db_error",
            "barcode": barcode,
            "source": source,
        }, 500
    finally:
        if conn:
            conn.close()

# Initialize DB and seed admin on startup
try:
    initialize_database()
except Exception as exc:
    print(f"[DB INIT ERROR] {exc}")


def serialize_book(row):
    if not row:
        return None

    book_id = row.get("book_id") or ""
    barcode_value = row.get("barcode_value") or book_id
    due_date_formatted = format_datetime(row.get("due_date"))
    created_at_formatted = format_datetime(row.get("created_at"))
    borrowed_by_id = (
        row.get("borrowed_by_id")
        or row.get("borrower_id")
        or row.get("borrowed_by_student_id")
        or row.get("student_id")
        or row.get("teacher_id")
        or ""
    )
    borrowed_by_name = row.get("borrowed_by_name") or row.get("borrower_name") or row.get("student_name") or row.get("teacher_name") or ""
    borrowed_by_email = row.get("borrowed_by_email") or row.get("borrower_email") or row.get("student_email") or row.get("teacher_email") or ""
    borrower_type = row.get("borrower_type") or ("TEACHER" if row.get("teacher_id") else "STUDENT" if row.get("student_id") else "")

    return {
        "id": row.get("id"),
        "bookId": book_id,
        "book_id": book_id,
        "barcode": barcode_value,
        "barcodeValue": barcode_value,
        "barcode_value": barcode_value,
        "title": row.get("title") or "",
        "author": row.get("author") or "",
        "category": row.get("category") or "",
        "status": row.get("status") or "AVAILABLE",
        "borrowerType": borrower_type,
        "borrower_type": borrower_type,
        "borrowedBy": borrowed_by_id,
        "borrowed_by": borrowed_by_id,
        "borrowedById": borrowed_by_id,
        "borrowed_by_id": borrowed_by_id,
        "borrowedByName": borrowed_by_name,
        "borrowed_by_name": borrowed_by_name,
        "borrowedByEmail": borrowed_by_email,
        "borrowed_by_email": borrowed_by_email,
        "dueDate": due_date_formatted,
        "due_date": due_date_formatted,
        "createdAt": created_at_formatted,
        "created_at": created_at_formatted,
    }


def serialize_borrowing(row):
    if not row:
        return None

    due_value = row.get("due_date")
    borrow_value = row.get("borrow_date")
    return_value = row.get("return_date")
    book_barcode = row.get("book_barcode") or row.get("barcode_value") or ""
    book_code = row.get("book_code") or row.get("book_id") or ""
    borrower_id = row.get("borrower_id") or row.get("student_id") or row.get("teacher_id") or ""
    borrower_name = row.get("borrower_name") or row.get("student_name") or row.get("teacher_name") or row.get("full_name") or ""
    borrower_email = row.get("borrower_email") or row.get("student_email") or row.get("teacher_email") or row.get("email") or ""
    borrower_barcode = row.get("borrower_barcode") or row.get("student_barcode") or row.get("teacher_barcode") or ""
    borrower_type = row.get("borrower_type") or ("TEACHER" if row.get("teacher_id") else "STUDENT")
    borrow_date_formatted = format_datetime(borrow_value)
    due_date_formatted = format_datetime(due_value)
    return_date_formatted = format_datetime(return_value)

    return {
        "id": row.get("id"),
        "borrowerType": borrower_type,
        "borrower_type": borrower_type,
        "borrowerId": borrower_id,
        "borrower_id": borrower_id,
        "borrowerName": borrower_name,
        "borrower_name": borrower_name,
        "borrowerEmail": borrower_email,
        "borrower_email": borrower_email,
        "borrowerBarcode": borrower_barcode,
        "borrower_barcode": borrower_barcode,
        "studentId": row.get("student_id") or "",
        "student_id": row.get("student_id") or "",
        "teacherId": row.get("teacher_id") or "",
        "teacher_id": row.get("teacher_id") or "",
        "bookPk": row.get("book_pk"),
        "book_pk": row.get("book_pk"),
        "bookId": book_code,
        "book_id": book_code,
        "bookCode": book_code,
        "book_code": book_code,
        "title": row.get("title") or "",
        "author": row.get("author") or "",
        "category": row.get("category") or "",
        "barcode": book_barcode,
        "barcode_value": book_barcode,
        "bookBarcode": book_barcode,
        "book_barcode": book_barcode,
        "borrowDate": borrow_date_formatted,
        "borrow_date": borrow_date_formatted,
        "dueDate": due_date_formatted,
        "due_date": due_date_formatted,
        "returnDate": return_date_formatted,
        "return_date": return_date_formatted,
        "status": row.get("status") or "",
    }

def serialize_unreturned_book(row):
    if not row:
        return None

    book_code = row.get("book_code") or row.get("book_id") or ""
    book_barcode = row.get("book_barcode") or row.get("barcode_value") or ""
    borrow_date_formatted = format_datetime(row.get("borrow_date"))
    due_date_formatted = format_datetime(row.get("due_date"))

    return {
        "bookPk": row.get("book_pk"),
        "book_pk": row.get("book_pk"),
        "bookId": book_code,
        "book_id": book_code,
        "bookCode": book_code,
        "book_code": book_code,
        "title": row.get("title") or "",
        "author": row.get("author") or "",
        "category": row.get("category") or "",
        "barcode": book_barcode,
        "barcode_value": book_barcode,
        "bookBarcode": book_barcode,
        "book_barcode": book_barcode,
        "borrowDate": borrow_date_formatted,
        "borrow_date": borrow_date_formatted,
        "dueDate": due_date_formatted,
        "due_date": due_date_formatted,
        "status": row.get("status") or "",
    }

def get_student_unreturned_books(cursor, student_id):
    cursor.execute(
        """
        SELECT
            b.id AS book_pk,
            b.book_id,
            b.book_id AS book_code,
            b.title,
            b.author,
            b.category,
            b.barcode_value AS book_barcode,
            bt.borrow_date,
            bt.due_date,
            CASE
                WHEN bt.return_date IS NOT NULL THEN 'RETURNED'
                WHEN bt.due_date < NOW() THEN 'OVERDUE'
                ELSE 'BORROWED'
            END AS status
        FROM borrow_transactions bt
        INNER JOIN books b
            ON b.id = bt.book_id
        WHERE bt.student_id = %s
          AND bt.return_date IS NULL
        ORDER BY
            CASE
                WHEN bt.due_date < NOW() THEN 0
                ELSE 1
            END,
            bt.due_date ASC,
            bt.id DESC
        """,
        (student_id,),
    )
    rows = cursor.fetchall()
    return [serialize_unreturned_book(row) for row in rows]

def generate_next_book_id(cursor):
    cursor.execute("SELECT COUNT(*) AS total FROM books")
    row = cursor.fetchone()
    total = row["total"] if isinstance(row, dict) else row[0]
    return f"BOOK-{int(total) + 1:05d}"

def get_student_by_scan(cursor, code):
    cursor.execute(
        """
        SELECT student_id, barcode_value, full_name, email, age, year_level, course, address, status, created_at
        FROM students
        WHERE student_id = %s OR barcode_value = %s
        LIMIT 1
        """,
        (code, code),
    )
    return cursor.fetchone()

def get_book_by_scan(cursor, code):
    cursor.execute(
        """
        SELECT id, book_id, barcode_value, title, author, category, status, created_at
        FROM books
        WHERE book_id = %s OR barcode_value = %s
        LIMIT 1
        """,
        (code, code),
    )
    return cursor.fetchone()

def get_active_borrowing_for_book(cursor, book_pk):
    cursor.execute(
        """
        SELECT
            bt.id,
            bt.student_id,
            s.full_name AS student_name,
            s.email AS student_email,
            s.barcode_value AS student_barcode,
            b.id AS book_pk,
            b.book_id,
            b.book_id AS book_code,
            b.title,
            b.author,
            b.category,
            b.barcode_value AS book_barcode,
            bt.borrow_date,
            bt.due_date,
            bt.return_date,
            CASE
                WHEN bt.return_date IS NOT NULL THEN 'RETURNED'
                WHEN bt.due_date < NOW() THEN 'OVERDUE'
                ELSE 'BORROWED'
            END AS status
        FROM borrow_transactions bt
        INNER JOIN students s
            ON s.student_id = bt.student_id
        INNER JOIN books b
            ON b.id = bt.book_id
        WHERE bt.book_id = %s
          AND bt.status = 'BORROWED'
        ORDER BY bt.id DESC
        LIMIT 1
        """,
        (book_pk,),
    )
    return cursor.fetchone()

def send_due_soon_emails():
    if not smtp_configured():
        return {
            "ok": False,
            "message": "SMTP is not configured yet.",
            "sentCount": 0,
            "failedCount": 0,
            "skippedCount": 0,
        }

    conn = None
    cursor = None
    sent_count = 0
    failed_count = 0
    skipped_count = 0

    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        cursor.execute(
            """
            SELECT
                bt.id,
                bt.borrower_type,
                bt.due_date,
                COALESCE(s.email, t.email) AS email,
                COALESCE(s.full_name, t.full_name) AS full_name,
                COALESCE(s.student_id, t.teacher_id) AS borrower_id,
                b.title
            FROM borrow_transactions bt
            LEFT JOIN students s ON s.student_id = bt.student_id
            LEFT JOIN teachers t ON t.teacher_id = bt.teacher_id
            JOIN books b ON b.id = bt.book_id
            WHERE bt.status = 'BORROWED'
              AND bt.return_date IS NULL
              AND bt.reminder_sent = 0
              AND bt.due_date BETWEEN NOW() AND DATE_ADD(NOW(), INTERVAL 1 DAY)
            """
        )
        rows = cursor.fetchall()

        for row in rows:
            email_address = normalize_email(row.get("email") or "")
            if not email_address or not is_valid_email(email_address):
                skipped_count += 1
                continue

            subject = "Library book due soon"
            borrower_label = "Teacher" if str(row.get("borrower_type") or "").upper() == "TEACHER" else "Student"
            body = (
                f"Hello {row['full_name']},\n\n"
                f"The borrowing time for '{row['title']}' is ending on {format_datetime(row['due_date'])}.\n"
                f"{borrower_label} ID: {row.get('borrower_id') or '--'}\n"
                "Please return it to the library on time.\n\n"
                "Thank you."
            )
            result = send_plain_email(email_address, subject, body)
            if result.get("sent"):
                sent_count += 1
                cursor.execute(
                    "UPDATE borrow_transactions SET reminder_sent = 1 WHERE id = %s",
                    (row["id"],),
                )
            else:
                failed_count += 1

        conn.commit()
        return {
            "ok": True,
            "message": f"Due-soon reminder run finished. Sent: {sent_count}, failed: {failed_count}, skipped: {skipped_count}.",
            "sentCount": sent_count,
            "failedCount": failed_count,
            "skippedCount": skipped_count,
        }
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def send_overdue_emails():
    if not smtp_configured():
        return {
            "ok": False,
            "message": "SMTP is not configured yet.",
            "sentCount": 0,
            "failedCount": 0,
            "skippedCount": 0,
        }

    conn = None
    cursor = None
    sent_count = 0
    failed_count = 0
    skipped_count = 0

    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        cursor.execute(
            """
            SELECT
                bt.id,
                bt.borrower_type,
                bt.due_date,
                COALESCE(s.email, t.email) AS email,
                COALESCE(s.full_name, t.full_name) AS full_name,
                COALESCE(s.student_id, t.teacher_id) AS borrower_id,
                b.title
            FROM borrow_transactions bt
            LEFT JOIN students s ON s.student_id = bt.student_id
            LEFT JOIN teachers t ON t.teacher_id = bt.teacher_id
            JOIN books b ON b.id = bt.book_id
            WHERE bt.status = 'BORROWED'
              AND bt.return_date IS NULL
              AND bt.due_date < NOW()
              AND COALESCE(bt.overdue_email_sent, 0) = 0
            ORDER BY bt.due_date ASC, bt.id ASC
            """
        )
        rows = cursor.fetchall()

        for row in rows:
            email_address = normalize_email(row.get("email") or "")
            if not email_address or not is_valid_email(email_address):
                skipped_count += 1
                continue

            subject = f"Overdue Reminder: {row['title']}"
            borrower_label = "Teacher" if str(row.get("borrower_type") or "").upper() == "TEACHER" else "Student"
            body = (
                f"Hello {row['full_name']},\n\n"
                "This is a reminder from the library that the deadline for the book below has already been reached and the book is still not returned.\n\n"
                f"{borrower_label} ID: {row.get('borrower_id') or '--'}\n"
                f"Book Title: {row['title']}\n"
                f"Deadline: {format_datetime(row['due_date'])}\n\n"
                "Please return the book to the library as soon as possible.\n\n"
                "Thank you.\n"
                "Libratrack Attendance System"
            )
            result = send_plain_email(email_address, subject, body)
            if result.get("sent"):
                sent_count += 1
                cursor.execute(
                    """
                    UPDATE borrow_transactions
                    SET overdue_email_sent = 1,
                        overdue_email_sent_at = NOW()
                    WHERE id = %s
                    """,
                    (row["id"],),
                )
            else:
                failed_count += 1

        conn.commit()
        return {
            "ok": True,
            "message": f"Overdue reminder run finished. Sent: {sent_count}, failed: {failed_count}, skipped: {skipped_count}.",
            "sentCount": sent_count,
            "failedCount": failed_count,
            "skippedCount": skipped_count,
        }
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def due_soon_email_loop():
    emit_status("[SYSTEM] Email reminder worker started.")
    while True:
        try:
            due_soon_result = send_due_soon_emails()
            overdue_result = send_overdue_emails()

            if due_soon_result.get("sentCount"):
                emit_status(f"[REMINDER] Sent {due_soon_result['sentCount']} due-soon email reminder(s).")
            if overdue_result.get("sentCount"):
                emit_status(f"[REMINDER] Sent {overdue_result['sentCount']} overdue email reminder(s).")
        except Exception as exc:
            emit_status(f"[REMINDER ERROR] {exc}")
        time.sleep(EMAIL_REMINDER_CHECK_SECONDS)

def start_due_soon_email_thread():
    global reminder_thread_started
    if reminder_thread_started:
        return
    thread = threading.Thread(target=due_soon_email_loop, daemon=True)
    thread.start()
    reminder_thread_started = True

def start_tracker_process():
    global tracker_process

    if not AUTO_START_TRACKER:
        print("[TRACKER] Auto-start disabled.")
        return

    if tracker_process and tracker_process.poll() is None:
        print("[TRACKER] Already running.")
        return

    tracker_script = resolve_tracker_script()

    if not os.path.isfile(tracker_script):
        print(f"[TRACKER ERROR] Script not found: {tracker_script}")
        return

    try:
        python_exe = sys.executable
        kwargs = {
            "args": [python_exe, tracker_script],
            "cwd": BASE_DIR,
        }

        if os.name == "nt":
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        tracker_process = subprocess.Popen(**kwargs)
        print(f"[TRACKER] Started hidden tracker: {tracker_script}")
    except Exception as exc:
        tracker_process = None
        print(f"[TRACKER ERROR] Failed to start tracker: {exc}")


def stop_tracker_process():
    global tracker_process

    if tracker_process and tracker_process.poll() is None:
        try:
            tracker_process.terminate()
            tracker_process.wait(timeout=5)
            print("[TRACKER] Tracker stopped.")
        except Exception:
            try:
                tracker_process.kill()
            except Exception:
                pass
    tracker_process = None


atexit.register(stop_tracker_process)
atexit.register(close_serial_connection)

# --------------------------
# PAGE ROUTES
# --------------------------

@app.route("/")
@app.route("/login")
@app.route("/login.html")
def login_page():
    if is_logged_in():
        return redirect(url_for("dashboard"))
    return render_template("login.html")

@app.route("/dashboard")
@app.route("/index.html")
def dashboard():
    if not is_logged_in():
        return redirect(url_for("login_page"))
    return render_template("index.html", username=session.get("username"))

@app.route("/attendance")
@app.route("/attendance.html")
def attendance_page():
    if not is_logged_in():
        return redirect(url_for("login_page"))
    return render_template("attendance.html", username=session.get("username"))

@app.route("/books")
@app.route("/books.html")
@app.route("/book")
@app.route("/book.html")
def books_page():
    if not is_logged_in():
        return redirect(url_for("login_page"))
    return render_template("book.html", username=session.get("username"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login_page"))

# --------------------------
# AUTH API
# --------------------------

@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json(silent=True) or {}
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", "")).strip()

    if not username or not password:
        return jsonify({"ok": False, "message": "Username and password are required."}), 400

    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            cursor.execute(
                "SELECT id, username, full_name, password_hash FROM users WHERE username = %s LIMIT 1",
                (username,)
            )
            user = cursor.fetchone()
        finally:
            cursor.close()

        if not user or user["password_hash"] != sha256_text(password):
            return jsonify({"ok": False, "message": "Invalid username or password."}), 401

        set_logged_in_user(user["id"], user["username"], user.get("full_name") or user["username"])
        return jsonify({
            "ok": True,
            "message": f"Welcome, {user['username']}!",
            "username": user["username"]
        })
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500

@app.route("/api/session", methods=["GET"])
def api_session():
    return jsonify({
        "ok": True,
        "loggedIn": bool(session.get("username")),
        "username": session.get("username"),
        "fullName": session.get("full_name"),
    })

# --------------------------
# STUDENT API
# --------------------------

@app.route("/api/student-id/next", methods=["GET"])
def api_next_student_id():
    conn = None
    try:
        conn = get_db_connection()
        return jsonify({"ok": True, "studentId": next_student_id(conn)})
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/students", methods=["GET"])
def api_students_list():
    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            cursor.execute(
                """
                SELECT
                    s.student_id,
                    s.barcode_value,
                    s.full_name,
                    s.email,
                    s.age,
                    s.year_level,
                    s.course,
                    s.address,
                    s.last_attendance,
                    s.status,
                    s.created_at,
                    EXISTS(
                        SELECT 1
                        FROM borrow_transactions bt
                        WHERE bt.student_id = s.student_id
                          AND bt.return_date IS NULL
                          AND bt.due_date < NOW()
                    ) AS has_overdue_books
                FROM students s
                ORDER BY has_overdue_books DESC, s.created_at DESC, s.student_id DESC
                """
            )
            rows = cursor.fetchall()
        finally:
            cursor.close()
        return jsonify({"ok": True, "students": [serialize_student(row) for row in rows]})
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/students", methods=["POST"])
def api_students_create():
    data = request.get_json(silent=True) or {}
    student_id = str(data.get("id", "")).strip()
    name = str(data.get("name", "")).strip()
    email = normalize_email(data.get("email", ""))
    age = data.get("age")
    year_level = str(data.get("year", "")).strip()
    course = str(data.get("course", "")).strip()
    address = str(data.get("address", "")).strip()

    if not name or not email or age in (None, ""):
        return jsonify({"ok": False, "message": "Name, email, and age are required."}), 400

    if not is_valid_email(email):
        return jsonify({"ok": False, "message": "Please enter a valid student email address."}), 400

    try:
        age = int(age)
    except Exception:
        return jsonify({"ok": False, "message": "Age must be a valid number."}), 400

    conn = None
    try:
        conn = get_db_connection()
        if not student_id:
            student_id = next_student_id(conn)

        cursor = dict_cursor(conn)
        try:
            cursor.execute("SELECT student_id FROM students WHERE student_id = %s LIMIT 1", (student_id,))
            existing = cursor.fetchone()
            if existing:
                return jsonify({"ok": False, "message": "Student ID already exists."}), 409

            cursor.execute("SELECT student_id FROM students WHERE email = %s LIMIT 1", (email,))
            existing_email = cursor.fetchone()
            if existing_email:
                return jsonify({"ok": False, "message": "Email is already registered to another student."}), 409

            cursor.execute(
                """
                INSERT INTO students (
                    student_id, barcode_value, full_name, email, age, year_level, course,
                    address, last_attendance, status
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NULL, 'Registered')
                """,
                (student_id, student_id, name, email, age, year_level, course, address)
            )
            conn.commit()
        finally:
            cursor.close()

        created = fetch_student_by_id(conn, student_id)
        email_result = send_barcode_email(created)
        base_message = "Student registered successfully."
        full_message = f"{base_message} {email_result['message']}".strip()

        return jsonify({
            "ok": True,
            "message": full_message,
            "student": serialize_student(created),
            "emailSent": email_result.get("sent", False),
            "emailMessage": email_result.get("message", ""),
        }), 201
    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/students/<student_id>/email", methods=["PUT"])
def api_students_update_email(student_id):
    data = request.get_json(silent=True) or {}
    email = normalize_email(data.get("email", ""))

    if not email:
        return jsonify({"ok": False, "message": "Email is required."}), 400

    if not is_valid_email(email):
        return jsonify({"ok": False, "message": "Please enter a valid student email address."}), 400

    conn = None
    try:
        conn = get_db_connection()
        student = fetch_student_by_id(conn, student_id)
        if not student:
            return jsonify({"ok": False, "message": "Student not found."}), 404

        cursor = dict_cursor(conn)
        try:
            cursor.execute(
                "SELECT student_id FROM students WHERE email = %s AND student_id <> %s LIMIT 1",
                (email, student_id)
            )
            existing_email = cursor.fetchone()
            if existing_email:
                return jsonify({"ok": False, "message": "Email is already registered to another student."}), 409

            cursor.execute("UPDATE students SET email = %s WHERE student_id = %s", (email, student_id))
            conn.commit()
        finally:
            cursor.close()

        updated = fetch_student_by_id(conn, student_id)
        return jsonify({
            "ok": True,
            "message": "Student email updated successfully.",
            "student": serialize_student(updated)
        })
    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/students/<student_id>/email-barcode", methods=["POST"])
def api_students_email_barcode(student_id):
    conn = None
    try:
        conn = get_db_connection()
        student = fetch_student_by_id(conn, student_id)
        if not student:
            return jsonify({"ok": False, "message": "Student not found."}), 404

        email_result = send_barcode_email(student)
        status_code = 200 if email_result.get("sent") else 500

        if "blank" in email_result.get("message", "").lower() or "not configured" in email_result.get("message", "").lower():
            status_code = 400

        return jsonify({
            "ok": bool(email_result.get("sent")),
            "message": email_result.get("message", "Barcode email processed."),
            "student": serialize_student(student),
        }), status_code
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Email error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/students/<student_id>", methods=["DELETE"])
def api_students_delete(student_id):
    conn = None
    try:
        conn = get_db_connection()
        student = fetch_student_by_id(conn, student_id)
        if not student:
            return jsonify({"ok": False, "message": "Student not found."}), 404

        cursor = dict_cursor(conn)
        try:
            cursor.execute("DELETE FROM students WHERE student_id = %s", (student_id,))
            conn.commit()
        finally:
            cursor.close()

        return jsonify({"ok": True, "message": f"Removed {student['full_name']} and related records."})
    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()


# --------------------------
# TEACHER API
# --------------------------

@app.route("/api/teacher-id/next", methods=["GET"])
def api_next_teacher_id():
    conn = None
    try:
        conn = get_db_connection()
        return jsonify({"ok": True, "teacherId": next_teacher_id(conn)})
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/teachers", methods=["GET"])
def api_teachers_list():
    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            cursor.execute(
                """
                SELECT teacher_id, barcode_value, full_name, email, department, address, status, created_at
                FROM teachers
                ORDER BY created_at DESC, teacher_id DESC
                """
            )
            rows = cursor.fetchall()
        finally:
            cursor.close()
        return jsonify({"ok": True, "teachers": [serialize_teacher(row) for row in rows]})
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/teachers", methods=["POST"])
def api_teachers_create():
    data = request.get_json(silent=True) or {}
    teacher_id = str(data.get("id", "")).strip()
    name = str(data.get("name", "")).strip()
    email = normalize_email(data.get("email", ""))
    department = str(data.get("department", "")).strip()
    address = str(data.get("address", "")).strip()

    if not name or not email:
        return jsonify({"ok": False, "message": "Teacher name and email are required."}), 400

    if not is_valid_email(email):
        return jsonify({"ok": False, "message": "Please enter a valid teacher email address."}), 400

    conn = None
    try:
        conn = get_db_connection()
        if not teacher_id:
            teacher_id = next_teacher_id(conn)

        cursor = dict_cursor(conn)
        try:
            cursor.execute("SELECT teacher_id FROM teachers WHERE teacher_id = %s LIMIT 1", (teacher_id,))
            existing = cursor.fetchone()
            if existing:
                return jsonify({"ok": False, "message": "Teacher ID already exists."}), 409

            cursor.execute("SELECT teacher_id FROM teachers WHERE email = %s LIMIT 1", (email,))
            existing_email = cursor.fetchone()
            if existing_email:
                return jsonify({"ok": False, "message": "Email is already registered to another teacher."}), 409

            cursor.execute(
                """
                INSERT INTO teachers (
                    teacher_id, barcode_value, full_name, email, department, address, status
                ) VALUES (%s, %s, %s, %s, %s, %s, 'Active')
                """,
                (teacher_id, teacher_id, name, email, department, address)
            )
            conn.commit()
        finally:
            cursor.close()

        created = fetch_teacher_by_id(conn, teacher_id)
        email_result = send_teacher_barcode_email(created)
        base_message = "Teacher registered successfully."
        full_message = f"{base_message} {email_result['message']}".strip()

        return jsonify({
            "ok": True,
            "message": full_message,
            "teacher": serialize_teacher(created),
            "emailSent": email_result.get("sent", False),
            "emailMessage": email_result.get("message", ""),
        }), 201
    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()


@app.route("/api/teachers/<teacher_id>", methods=["DELETE"])
def api_teachers_delete(teacher_id):
    conn = None
    try:
        conn = get_db_connection()
        teacher = fetch_teacher_by_id(conn, teacher_id)
        if not teacher:
            return jsonify({"ok": False, "message": "Teacher not found."}), 404

        cursor = dict_cursor(conn)
        try:
            cursor.execute("DELETE FROM teachers WHERE teacher_id = %s", (teacher_id,))
            conn.commit()
        finally:
            cursor.close()

        return jsonify({"ok": True, "message": f"Removed {teacher['full_name']} and related records."})
    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

# --------------------------
# LIBRARY API
# --------------------------


@app.route("/api/books", methods=["GET"])
def api_books():
    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            cursor.execute(
                """
                SELECT
                    b.id,
                    b.book_id,
                    b.barcode_value,
                    b.title,
                    b.author,
                    b.category,
                    CASE
                        WHEN bt.id IS NOT NULL AND bt.return_date IS NULL AND bt.due_date < NOW() THEN 'OVERDUE'
                        WHEN bt.id IS NOT NULL AND bt.return_date IS NULL THEN 'BORROWED'
                        ELSE 'AVAILABLE'
                    END AS status,
                    b.created_at,
                    bt.borrower_type,
                    COALESCE(s.student_id, t.teacher_id) AS borrowed_by_id,
                    COALESCE(s.full_name, t.full_name) AS borrowed_by_name,
                    COALESCE(s.email, t.email) AS borrowed_by_email,
                    bt.due_date
                FROM books b
                LEFT JOIN borrow_transactions bt
                    ON bt.book_id = b.id
                   AND bt.return_date IS NULL
                LEFT JOIN students s
                    ON s.student_id = bt.student_id
                LEFT JOIN teachers t
                    ON t.teacher_id = bt.teacher_id
                ORDER BY b.created_at DESC, b.id DESC
                """
            )
            rows = cursor.fetchall()
        finally:
            cursor.close()

        return jsonify({"ok": True, "books": [serialize_book(row) for row in rows]})
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/books/lookup", methods=["GET"])
def api_books_lookup():
    code = str(request.args.get("code", "")).strip()

    if not code:
        return jsonify({"ok": False, "message": "Book code is required."}), 400

    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            book = get_book_by_scan(cursor, code)
            if not book:
                return jsonify({"ok": False, "message": "Book not found."}), 404

            current_borrowing = get_active_borrowing_for_book(cursor, book["id"])
        finally:
            cursor.close()

        return jsonify({
            "ok": True,
            "message": "Book information loaded.",
            "book": serialize_book(book),
            "currentBorrowing": serialize_borrowing(current_borrowing) if current_borrowing else None,
            "current_borrowing": serialize_borrowing(current_borrowing) if current_borrowing else None,
        })
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/book-id/next", methods=["GET"])
def api_next_book_id():
    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            next_id = generate_next_book_id(cursor)
        finally:
            cursor.close()

        return jsonify({"ok": True, "bookId": next_id})
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/books", methods=["POST"])
def api_register_book():
    payload = request.get_json(silent=True) or {}
    title = str(payload.get("title", "")).strip()
    author = str(payload.get("author", "")).strip()
    category = str(payload.get("category", "")).strip()
    book_id = str(payload.get("book_id") or payload.get("bookId") or "").strip()

    if not title:
        return jsonify({"ok": False, "message": "Book title is required."}), 400

    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            if not book_id:
                book_id = generate_next_book_id(cursor)

            cursor.execute("SELECT id FROM books WHERE book_id = %s LIMIT 1", (book_id,))
            existing_id = cursor.fetchone()
            if existing_id:
                return jsonify({"ok": False, "message": "Book ID already exists."}), 409

            barcode_value = book_id
            cursor.execute("SELECT id FROM books WHERE barcode_value = %s LIMIT 1", (barcode_value,))
            existing_barcode = cursor.fetchone()
            if existing_barcode:
                return jsonify({"ok": False, "message": "Barcode value already exists."}), 409

            cursor.execute(
                """
                INSERT INTO books (book_id, barcode_value, title, author, category, status)
                VALUES (%s, %s, %s, %s, %s, 'AVAILABLE')
                """,
                (book_id, barcode_value, title, author, category),
            )
            conn.commit()

            cursor.execute(
                """
                SELECT id, book_id, barcode_value, title, author, category, status, created_at
                FROM books
                WHERE book_id = %s
                LIMIT 1
                """,
                (book_id,),
            )
            book = cursor.fetchone()
        finally:
            cursor.close()

        return jsonify({
            "ok": True,
            "message": "Book registered successfully. Print the barcode and stick it to the book.",
            "book": serialize_book(book),
            "barcodeValue": book_id,
            "barcodeSvgUrl": f"/api/barcode/{book_id}.svg",
            "barcodePngUrl": f"/api/barcode/{book_id}.png",
        }), 201
    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()


@app.route("/api/books/<book_id>", methods=["DELETE"])
def api_delete_book(book_id):
    book_id = str(book_id or "").strip()
    if not book_id:
        return jsonify({"ok": False, "message": "Book ID is required."}), 400

    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            cursor.execute(
                """
                SELECT id, book_id, barcode_value, title, author, category, status, created_at
                FROM books
                WHERE book_id = %s
                LIMIT 1
                """,
                (book_id,),
            )
            book = cursor.fetchone()
            if not book:
                return jsonify({"ok": False, "message": "Book not found."}), 404

            cursor.execute(
                """
                SELECT id
                FROM borrow_transactions
                WHERE book_id = %s
                  AND return_date IS NULL
                LIMIT 1
                """,
                (book["id"],),
            )
            active_borrowing = cursor.fetchone()
            if active_borrowing:
                return jsonify({
                    "ok": False,
                    "message": "This book cannot be deleted because it is still borrowed. Return the book first."
                }), 409

            cursor.execute("DELETE FROM books WHERE id = %s", (book["id"],))
            conn.commit()
        finally:
            cursor.close()

        socketio.emit("bookDeleted", {
            "ok": True,
            "bookId": book_id,
            "message": f"{book['title']} was deleted from the inventory."
        })

        return jsonify({
            "ok": True,
            "message": f"{book['title']} was deleted from the inventory.",
            "book": serialize_book(book)
        })
    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()


@app.route("/api/borrowings/issue", methods=["POST"])
def api_issue_book():
    payload = request.get_json(silent=True) or {}
    book_code = str(payload.get("book_code") or payload.get("bookCode") or "").strip()
    borrower_type = str(payload.get("borrower_type") or payload.get("borrowerType") or "STUDENT").strip().upper()
    borrower_code = str(
        payload.get("borrower_code")
        or payload.get("borrowerCode")
        or payload.get("student_code")
        or payload.get("studentCode")
        or ""
    ).strip()
    due_date_raw = str(payload.get("due_date") or payload.get("dueDate") or "").strip()

    if not book_code or not borrower_code or not due_date_raw:
        return jsonify({"ok": False, "message": "Book scan, borrower scan, and due date are required."}), 400

    if borrower_type not in {"STUDENT", "TEACHER"}:
        return jsonify({"ok": False, "message": "Borrower type must be STUDENT or TEACHER."}), 400

    try:
        due_date = parse_due_date(due_date_raw)
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400

    if due_date <= now_dt():
        return jsonify({"ok": False, "message": "Due date must be in the future."}), 400

    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            if borrower_type == "TEACHER":
                cursor.execute(
                    """
                    SELECT teacher_id, barcode_value, full_name, email, department, address, status, created_at
                    FROM teachers
                    WHERE teacher_id = %s OR barcode_value = %s
                    LIMIT 1
                    """,
                    (borrower_code, borrower_code),
                )
            else:
                cursor.execute(
                    """
                    SELECT student_id, barcode_value, full_name, email, age, year_level, course, address, status, created_at
                    FROM students
                    WHERE student_id = %s OR barcode_value = %s
                    LIMIT 1
                    """,
                    (borrower_code, borrower_code),
                )
            borrower = cursor.fetchone()

            if not borrower:
                return jsonify({"ok": False, "message": f"{'Teacher' if borrower_type == 'TEACHER' else 'Student'} not found."}), 404

            book = get_book_by_scan(cursor, book_code)
            if not book:
                return jsonify({"ok": False, "message": "Book not registered yet. Register it first."}), 404

            cursor.execute(
                "SELECT id FROM borrow_transactions WHERE book_id = %s AND return_date IS NULL LIMIT 1",
                (book["id"],),
            )
            active = cursor.fetchone()
            if active:
                return jsonify({"ok": False, "message": "This book is already borrowed and not yet returned."}), 409

            student_id = borrower["student_id"] if borrower_type == "STUDENT" else None
            teacher_id = borrower["teacher_id"] if borrower_type == "TEACHER" else None

            cursor.execute(
                """
                INSERT INTO borrow_transactions (
                    borrower_type, student_id, teacher_id, book_id, due_date, status, reminder_sent
                )
                VALUES (%s, %s, %s, %s, %s, 'BORROWED', 0)
                """,
                (borrower_type, student_id, teacher_id, book["id"], due_date),
            )
            borrow_id = cursor.lastrowid

            cursor.execute(
                "UPDATE books SET status = 'BORROWED' WHERE id = %s",
                (book["id"],),
            )
            conn.commit()

            cursor.execute(
                """
                SELECT
                    bt.id,
                    bt.borrower_type,
                    bt.student_id,
                    bt.teacher_id,
                    COALESCE(s.student_id, t.teacher_id) AS borrower_id,
                    COALESCE(s.full_name, t.full_name) AS borrower_name,
                    COALESCE(s.email, t.email) AS borrower_email,
                    COALESCE(s.barcode_value, t.barcode_value) AS borrower_barcode,
                    b.id AS book_pk,
                    b.book_id,
                    b.book_id AS book_code,
                    b.title,
                    b.author,
                    b.category,
                    b.barcode_value AS book_barcode,
                    bt.borrow_date,
                    bt.due_date,
                    bt.return_date,
                    CASE
                        WHEN bt.return_date IS NOT NULL THEN 'RETURNED'
                        WHEN bt.due_date < NOW() THEN 'OVERDUE'
                        ELSE 'BORROWED'
                    END AS status
                FROM borrow_transactions bt
                LEFT JOIN students s ON s.student_id = bt.student_id
                LEFT JOIN teachers t ON t.teacher_id = bt.teacher_id
                JOIN books b ON b.id = bt.book_id
                WHERE bt.id = %s
                LIMIT 1
                """,
                (borrow_id,),
            )
            borrowing = cursor.fetchone()
        finally:
            cursor.close()

        borrowed_book_payload = {
            **book,
            "status": "BORROWED",
            "borrower_type": borrower_type,
            "borrowed_by_id": borrower.get("teacher_id") if borrower_type == "TEACHER" else borrower.get("student_id"),
            "borrowed_by_name": borrower.get("full_name") or "",
            "borrowed_by_email": borrower.get("email") or "",
            "due_date": due_date,
        }
        borrow_email_result = send_book_borrowed_email(borrower, book, due_date, borrower_type=borrower_type)

        socketio.emit("bookBorrowed", {
            "ok": True,
            "borrowing": serialize_borrowing(borrowing),
            "book": serialize_book(borrowed_book_payload),
            "borrower": serialize_teacher(borrower) if borrower_type == "TEACHER" else serialize_student(borrower),
        })

        base_message = f"{book['title']} borrowed by {borrower['full_name']} until {format_datetime(due_date)}."
        full_message = f"{base_message} {borrow_email_result.get('message', '')}".strip()

        response = {
            "ok": True,
            "message": full_message,
            "book": serialize_book(borrowed_book_payload),
            "borrowing": serialize_borrowing(borrowing),
            "borrowerType": borrower_type,
            "borrower_type": borrower_type,
            "borrowEmailSent": borrow_email_result.get("sent", False),
            "borrowEmailMessage": borrow_email_result.get("message", ""),
        }
        if borrower_type == "TEACHER":
            response["teacher"] = serialize_teacher(borrower)
        else:
            response["student"] = serialize_student(borrower)

        return jsonify(response), 201
    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/borrowings/return", methods=["POST"])
def api_return_book():
    payload = request.get_json(silent=True) or {}
    book_code = str(payload.get("book_code") or payload.get("bookCode") or "").strip()

    if not book_code:
        return jsonify({"ok": False, "message": "Book scan is required."}), 400

    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            book = get_book_by_scan(cursor, book_code)
            if not book:
                return jsonify({"ok": False, "message": "Book not found."}), 404

            cursor.execute(
                """
                SELECT
                    bt.id,
                    bt.borrower_type,
                    bt.student_id,
                    bt.teacher_id,
                    COALESCE(s.student_id, t.teacher_id) AS borrower_id,
                    COALESCE(s.full_name, t.full_name) AS borrower_name,
                    COALESCE(s.email, t.email) AS borrower_email,
                    COALESCE(s.barcode_value, t.barcode_value) AS borrower_barcode,
                    b.id AS book_pk,
                    b.book_id,
                    b.book_id AS book_code,
                    b.title,
                    b.author,
                    b.category,
                    b.barcode_value AS book_barcode,
                    bt.borrow_date,
                    bt.due_date,
                    bt.return_date,
                    CASE
                        WHEN bt.return_date IS NOT NULL THEN 'RETURNED'
                        WHEN bt.due_date < NOW() THEN 'OVERDUE'
                        ELSE 'BORROWED'
                    END AS status
                FROM borrow_transactions bt
                LEFT JOIN students s ON s.student_id = bt.student_id
                LEFT JOIN teachers t ON t.teacher_id = bt.teacher_id
                JOIN books b ON b.id = bt.book_id
                WHERE bt.book_id = %s AND bt.return_date IS NULL
                ORDER BY bt.id DESC
                LIMIT 1
                """,
                (book["id"],),
            )
            active = cursor.fetchone()

            if not active:
                return jsonify({"ok": False, "message": "This book already appears returned."}), 409

            returned_at = now_dt()
            cursor.execute(
                """
                UPDATE borrow_transactions
                SET return_date = %s, status = 'RETURNED'
                WHERE id = %s
                """,
                (returned_at, active["id"]),
            )
            cursor.execute("UPDATE books SET status = 'AVAILABLE' WHERE id = %s", (book["id"],))
            conn.commit()

            cursor.execute(
                """
                SELECT
                    bt.id,
                    bt.borrower_type,
                    bt.student_id,
                    bt.teacher_id,
                    COALESCE(s.student_id, t.teacher_id) AS borrower_id,
                    COALESCE(s.full_name, t.full_name) AS borrower_name,
                    COALESCE(s.email, t.email) AS borrower_email,
                    COALESCE(s.barcode_value, t.barcode_value) AS borrower_barcode,
                    b.id AS book_pk,
                    b.book_id,
                    b.book_id AS book_code,
                    b.title,
                    b.author,
                    b.category,
                    b.barcode_value AS book_barcode,
                    bt.borrow_date,
                    bt.due_date,
                    bt.return_date,
                    CASE
                        WHEN bt.return_date IS NOT NULL THEN 'RETURNED'
                        WHEN bt.due_date < NOW() THEN 'OVERDUE'
                        ELSE 'BORROWED'
                    END AS status
                FROM borrow_transactions bt
                LEFT JOIN students s ON s.student_id = bt.student_id
                LEFT JOIN teachers t ON t.teacher_id = bt.teacher_id
                JOIN books b ON b.id = bt.book_id
                WHERE bt.id = %s
                LIMIT 1
                """,
                (active["id"],),
            )
            completed = cursor.fetchone()
        finally:
            cursor.close()

        socketio.emit("bookReturned", {
            "ok": True,
            "borrowing": serialize_borrowing(completed),
            "book": serialize_book({
                **book,
                "status": "AVAILABLE",
                "due_date": None,
                "borrowed_by_student_id": "",
                "borrowed_by_name": "",
                "borrowed_by_email": "",
            }),
        })

        return jsonify({
            "ok": True,
            "message": f"{active['title']} returned successfully by {active.get('borrower_name') or 'the borrower'}.",
            "borrowing": serialize_borrowing(completed),
            "book": serialize_book({
                **book,
                "status": "AVAILABLE",
                "due_date": None,
                "borrowed_by_student_id": "",
                "borrowed_by_name": "",
                "borrowed_by_email": "",
            }),
        })
    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()


@app.route("/api/borrowings", methods=["GET"])
def api_borrowings():
    conn = None
    try:
        send_overdue_emails()
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            cursor.execute(
                """
                SELECT
                    bt.id,
                    bt.borrower_type,
                    bt.student_id,
                    bt.teacher_id,
                    COALESCE(s.student_id, t.teacher_id) AS borrower_id,
                    COALESCE(s.full_name, t.full_name) AS borrower_name,
                    COALESCE(s.email, t.email) AS borrower_email,
                    COALESCE(s.barcode_value, t.barcode_value) AS borrower_barcode,
                    b.id AS book_pk,
                    b.book_id,
                    b.book_id AS book_code,
                    b.title,
                    b.author,
                    b.category,
                    b.barcode_value AS book_barcode,
                    bt.borrow_date,
                    bt.due_date,
                    bt.return_date,
                    CASE
                        WHEN bt.return_date IS NOT NULL THEN 'RETURNED'
                        WHEN bt.due_date < NOW() THEN 'OVERDUE'
                        ELSE 'BORROWED'
                    END AS status
                FROM borrow_transactions bt
                LEFT JOIN students s
                    ON s.student_id = bt.student_id
                LEFT JOIN teachers t
                    ON t.teacher_id = bt.teacher_id
                INNER JOIN books b
                    ON b.id = bt.book_id
                ORDER BY
                    CASE
                        WHEN bt.return_date IS NULL AND bt.due_date < NOW() THEN 0
                        WHEN bt.return_date IS NULL THEN 1
                        ELSE 2
                    END,
                    bt.borrow_date DESC,
                    bt.id DESC
                """
            )
            rows = cursor.fetchall()
        finally:
            cursor.close()

        return jsonify({"ok": True, "borrowings": [serialize_borrowing(row) for row in rows]})
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/borrowings/reminders/send-due-soon", methods=["POST"])
def api_send_due_soon_reminders():
    try:
        result = send_due_soon_emails()
        status_code = 200 if result.get("ok") else 500
        return jsonify(result), status_code
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Reminder error: {exc}"}), 500

@app.route("/api/borrowings/reminders/send-overdue", methods=["POST"])
def api_send_overdue_reminders():
    try:
        result = send_overdue_emails()
        status_code = 200 if result.get("ok") else 500
        return jsonify(result), status_code
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Reminder error: {exc}"}), 500

@app.route("/api/borrowings/clear", methods=["DELETE"])
def api_clear_borrowings():
    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            cursor.execute("UPDATE books SET status = 'AVAILABLE'")
            cursor.execute("DELETE FROM borrow_transactions")
            deleted = cursor.rowcount
            conn.commit()
        finally:
            cursor.close()

        socketio.emit("borrowingsCleared", {
            "ok": True,
            "message": f"Cleared {deleted} borrowing record(s)."
        })

        return jsonify({
            "ok": True,
            "message": f"Cleared {deleted} borrowing record(s).",
            "deleted": deleted
        })
    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

# --------------------------
# BARCODE API
# --------------------------

@app.route("/api/barcode/<path:barcode_value>.png", methods=["GET"])
def api_barcode_png(barcode_value):
    try:
        png_bytes = generate_barcode_png(barcode_value)
        return Response(
            png_bytes,
            mimetype="image/png",
            headers={"Content-Disposition": f'inline; filename=\"barcode_{barcode_value}.png\"'}
        )
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Barcode error: {exc}"}), 400

@app.route("/api/barcode/<path:barcode_value>.svg", methods=["GET"])
def api_barcode_svg(barcode_value):
    try:
        svg_bytes = generate_barcode_svg(barcode_value)
        return Response(
            svg_bytes,
            mimetype="image/svg+xml",
            headers={"Content-Disposition": f'inline; filename=\"barcode_{barcode_value}.svg\"'}
        )
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Barcode error: {exc}"}), 400

# --------------------------
# ATTENDANCE API
# --------------------------

@app.route("/api/attendance", methods=["GET"])
def api_attendance_list():
    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            cursor.execute(
                """
                SELECT a.id, a.student_id, s.barcode_value, s.full_name, a.attendance_date, a.time_in,
                       a.time_out, a.last_action, a.status, a.created_at
                FROM attendance a
                INNER JOIN students s ON s.student_id = a.student_id
                ORDER BY a.id DESC
                """
            )
            rows = cursor.fetchall()
        finally:
            cursor.close()
        return jsonify({"ok": True, "attendance": [serialize_attendance(row) for row in rows]})
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/attendance/scan", methods=["POST"])
@app.route("/api/scan-attendance", methods=["POST"])
def api_attendance_scan():
    data = request.get_json(silent=True) or {}
    barcode = str(data.get("barcode") or data.get("code") or data.get("studentId") or "").strip()
    payload, status_code = process_barcode_scan(barcode, source="api")

    if status_code == 200:
        action = str(payload.get("action", "")).upper()

        if action == "TIME IN":
            send_led_command("LED:GREEN")
            if TIME_IN_AUDIO_ENABLED:
                speak_async(TIME_IN_AUDIO_MESSAGE)
        elif action == "TIME OUT":
            send_led_command("LED:RED")
            if TIME_OUT_AUDIO_ENABLED:
                speak_async(TIME_OUT_AUDIO_MESSAGE)
        else:
            send_led_command("LED:ERROR")

        socketio.emit("attendanceUpdated", payload)
        socketio.emit("attendanceTableChanged", {
            "ok": True,
            "message": payload.get("message", "Attendance updated."),
            "record": payload.get("record"),
            "attendance": payload.get("attendance"),
            "student": payload.get("student"),
        })
    else:
        send_led_command("LED:ERROR")
        socketio.emit("barcodeRejected", payload)

    return jsonify(payload), status_code

@app.route("/api/attendance/today", methods=["DELETE"])
def api_attendance_clear_today():
    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            cursor.execute("DELETE FROM attendance WHERE attendance_date = %s", (date.today(),))
            deleted = cursor.rowcount
            conn.commit()
        finally:
            cursor.close()

        socketio.emit("attendanceTableChanged", {
            "ok": True,
            "message": f"Cleared {deleted} attendance records for today."
        })
        return jsonify({
            "ok": True,
            "message": f"Cleared {deleted} attendance records for today.",
            "deleted": deleted
        })
    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500
    finally:
        if conn:
            conn.close()


@app.route("/api/attendance/clear-all", methods=["DELETE"])
def api_attendance_clear_all():
    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            cursor.execute("DELETE FROM attendance")
            deleted = cursor.rowcount
            conn.commit()
        finally:
            cursor.close()

        socketio.emit("attendanceTableChanged", {
            "ok": True,
            "message": f"Cleared {deleted} attendance record(s)."
        })

        return jsonify({
            "ok": True,
            "message": f"Cleared {deleted} attendance record(s).",
            "deleted": deleted
        })

    except Exception as exc:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Database error: {exc}"}), 500

    finally:
        if conn:
            conn.close()

@app.route("/simulate_scan", methods=["POST"])
def simulate_scan():
    data = request.get_json(silent=True) or {}
    barcode = str(data.get("barcode", "")).strip()
    payload, status_code = process_barcode_scan(barcode, source="simulate_scan")

    if status_code == 200:
        socketio.emit("attendanceUpdated", payload)
    else:
        socketio.emit("barcodeRejected", payload)

    return jsonify(payload), status_code

@app.route("/api/db_status", methods=["GET"])
def api_db_status():
    conn = None
    try:
        conn = get_db_connection()
        cursor = dict_cursor(conn)
        try:
            cursor.execute("SELECT NOW() AS server_time")
            row = cursor.fetchone()
        finally:
            cursor.close()

        return jsonify({
            "ok": True,
            "driver": MYSQL_MODE,
            "database": DB_NAME,
            "serverTime": format_datetime(row.get("server_time") if isinstance(row, dict) else None),
        })
    except Exception as exc:
        return jsonify({"ok": False, "message": str(exc), "driver": MYSQL_MODE}), 500
    finally:
        if conn:
            conn.close()

# --------------------------
# SOCKET EVENTS
# --------------------------

@socketio.on("connect")
def handle_connect():
    print("[SOCKET] Client connected")
    socketio.emit("systemStatus", {"message": "Connected to server"})

@socketio.on("disconnect")
def handle_disconnect():
    print("[SOCKET] Client disconnected")

if __name__ == "__main__":
    # Run the backend and ML tracker only once.
    # Flask's debug reloader can spawn duplicate processes, which makes the
    # face-tracker / ML output overlap or appear in multiple panels.
    start_due_soon_email_thread()
    start_serial_listener_thread()
    start_tracker_process()

    socketio.run(
        app,
        host="0.0.0.0",
        port=5000,
        debug=False,
        use_reloader=False,
        allow_unsafe_werkzeug=True
    )
