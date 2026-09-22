import csv
import os
import shutil
import sqlite3
import re
import subprocess
import tempfile
import time
from datetime import datetime, date, timedelta

import sys
import threading
import webbrowser

# -----------------------------
# 중복 실행 방지 (뮤텍스)
# -----------------------------
APP_MUTEX_NAME = "Global\\SchoolPrinterKioskMutex"

# 뮤텍스 핸들은 프로그램이 살아있는 동안 계속 유지되어야 한다.
# (지역 변수로 두고 닫아버리면 뮤텍스가 즉시 사라져서 중복 실행을 못 막는다)
_APP_MUTEX_HANDLE = None


def is_app_already_running():
    global _APP_MUTEX_HANDLE

    try:
        import win32event
        import win32api
        import winerror
    except ImportError:
        # pywin32가 없으면 중복 실행 검사를 건너뛴다
        return False

    try:
        _APP_MUTEX_HANDLE = win32event.CreateMutex(None, False, APP_MUTEX_NAME)
        return win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS
    except Exception:
        return False


if is_app_already_running():
    sys.exit(0)

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog, simpledialog

try:
    import customtkinter as ctk
except ImportError:
    ctk = None

try:
    import win32print
except ImportError:
    win32print = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


DB_NAME = "school_printer.db"

# -----------------------------
# 자동 업데이트 설정
#
# ── 새 버전 배포 방법 ──
#   1. 아래 APP_VERSION 을 올린다      (예: 1.0.0 → 1.1.0)
#   2. exe 를 새로 빌드한다
#   3. GitHub 저장소 > Releases > 새 릴리스 작성
#        · 태그 이름 : v1.1.0        (앞의 v 를 반드시 포함)
#        · 빌드한 exe 파일을 첨부
#   4. 키오스크에서 관리자 모드 > 보호 기능 설정 > 업데이트 확인
#
# 프로그램이 스스로 업데이트를 확인하는 일은 없다.
# 관리자가 버튼을 눌렀을 때만 확인한다.
# -----------------------------
APP_VERSION = "1.5.3"
GITHUB_REPO = "sicgaonnury/printer"
DEFAULT_ADMIN_PASSWORD = "1234"

SUPPORTED_EXTENSIONS = [
    ".pdf",
    ".doc", ".docx",
    ".hwp", ".hwpx",
    ".jpg", ".jpeg", ".png",
    ".ppt", ".pptx",
    ".csv",
    ".xls", ".xlsx",
    ".txt"
]

LIBREOFFICE_EXTENSIONS = [
    ".doc", ".docx",
    ".hwp", ".hwpx",
    ".ppt", ".pptx",
    ".xls", ".xlsx",
    ".csv",
    ".txt"
]

IMAGE_EXTENSIONS = [".jpg", ".jpeg", ".png"]


# -----------------------------
# 한글 모드 바코드 입력 보정
# -----------------------------
CHOSUNG = [
    "ㄱ", "ㄲ", "ㄴ", "ㄷ", "ㄸ", "ㄹ", "ㅁ", "ㅂ", "ㅃ", "ㅅ",
    "ㅆ", "ㅇ", "ㅈ", "ㅉ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ"
]

JUNGSUNG = [
    "ㅏ", "ㅐ", "ㅑ", "ㅒ", "ㅓ", "ㅔ", "ㅕ", "ㅖ", "ㅗ", "ㅘ",
    "ㅙ", "ㅚ", "ㅛ", "ㅜ", "ㅝ", "ㅞ", "ㅟ", "ㅠ", "ㅡ", "ㅢ", "ㅣ"
]

JONGSUNG = [
    "", "ㄱ", "ㄲ", "ㄳ", "ㄴ", "ㄵ", "ㄶ", "ㄷ", "ㄹ", "ㄺ",
    "ㄻ", "ㄼ", "ㄽ", "ㄾ", "ㄿ", "ㅀ", "ㅁ", "ㅂ", "ㅄ", "ㅅ",
    "ㅆ", "ㅇ", "ㅈ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ"
]

KOREAN_TO_ENG_KEY = {
    "ㅂ": "q", "ㅃ": "Q",
    "ㅈ": "w", "ㅉ": "W",
    "ㄷ": "e", "ㄸ": "E",
    "ㄱ": "r", "ㄲ": "R",
    "ㅅ": "t", "ㅆ": "T",
    "ㅛ": "y",
    "ㅕ": "u",
    "ㅑ": "i",
    "ㅐ": "o", "ㅒ": "O",
    "ㅔ": "p", "ㅖ": "P",
    "ㅁ": "a",
    "ㄴ": "s",
    "ㅇ": "d",
    "ㄹ": "f",
    "ㅎ": "g",
    "ㅗ": "h",
    "ㅓ": "j",
    "ㅏ": "k",
    "ㅣ": "l",
    "ㅋ": "z",
    "ㅌ": "x",
    "ㅊ": "c",
    "ㅍ": "v",
    "ㅠ": "b",
    "ㅜ": "n",
    "ㅡ": "m",

    "ㅘ": "hk",
    "ㅙ": "ho",
    "ㅚ": "hl",
    "ㅝ": "nj",
    "ㅞ": "np",
    "ㅟ": "nl",
    "ㅢ": "ml",

    "ㄳ": "rt",
    "ㄵ": "sw",
    "ㄶ": "sg",
    "ㄺ": "fr",
    "ㄻ": "fa",
    "ㄼ": "fq",
    "ㄽ": "ft",
    "ㄾ": "fx",
    "ㄿ": "fv",
    "ㅀ": "fg",
    "ㅄ": "qt",
}


def hangul_syllable_to_eng_keys(ch):
    code = ord(ch)

    if not (0xAC00 <= code <= 0xD7A3):
        return KOREAN_TO_ENG_KEY.get(ch, ch)

    base = code - 0xAC00
    cho = base // 588
    jung = (base % 588) // 28
    jong = base % 28

    result = ""
    result += KOREAN_TO_ENG_KEY.get(CHOSUNG[cho], CHOSUNG[cho])
    result += KOREAN_TO_ENG_KEY.get(JUNGSUNG[jung], JUNGSUNG[jung])

    if JONGSUNG[jong]:
        result += KOREAN_TO_ENG_KEY.get(JONGSUNG[jong], JONGSUNG[jong])

    return result


def normalize_card_code(text):
    if text is None:
        return ""

    text = str(text).strip()
    converted = "".join(hangul_syllable_to_eng_keys(ch) for ch in text)
    converted = converted.replace(" ", "")
    return converted.upper()


# -----------------------------
# DB 관련 함수
# -----------------------------
# 업로드 서버 스레드, 자동 삭제 타이머, 화면 스레드가 각각 DB 를 열기 때문에
# 서로 겹치면 "database is locked" 가 날 수 있다.
# WAL 모드로 읽기와 쓰기가 서로를 막지 않게 하고, 잠기면 잠시 기다리게 한다.
_db_ready = {"wal": False}


def connect_db():
    conn = sqlite3.connect(DB_NAME, timeout=10)

    try:
        conn.execute("PRAGMA busy_timeout=10000")

        if not _db_ready["wal"]:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA synchronous=NORMAL")
            _db_ready["wal"] = True
    except Exception:
        pass

    return conn


def hash_password(password, salt=None):
    """관리자 비밀번호를 해시로 바꾼다. 형식: sha256$소금$해시값"""
    import hashlib
    import secrets

    salt = salt or secrets.token_hex(8)
    digest = hashlib.sha256((salt + str(password)).encode("utf-8")).hexdigest()
    return f"sha256${salt}${digest}"


def verify_password(password, stored):
    """
    입력한 비밀번호가 맞는지 확인한다.

    예전 버전은 비밀번호를 그대로 저장했으므로, 저장된 값이 해시 형식이 아니면
    평문끼리 비교한다. (로그인에 성공하면 그 자리에서 해시로 바꿔 저장한다)
    """
    stored = str(stored or "")

    if stored.startswith("sha256$"):
        try:
            _, salt, _ = stored.split("$", 2)
        except ValueError:
            return False
        return hash_password(password, salt) == stored

    return str(password) == stored


def is_legacy_password(stored):
    """저장된 비밀번호가 아직 평문인지."""
    return not str(stored or "").startswith("sha256$")


def printer_exists(printer_name):
    """이름이 같은 프린터가 실제로 있는지 확인한다."""
    if not printer_name:
        return False

    if win32print is None:
        return True   # 확인할 방법이 없으면 막지 않는다

    try:
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        names = [p[2] for p in win32print.EnumPrinters(flags)]
    except Exception:
        return True

    return printer_name in names


# ── 프린터 상태 확인용 값 (winspool 정의) ──
# SumatraPDF 는 "인쇄 대기열에 넣는 데 성공"만 알려주기 때문에,
# 케이블이 빠져 있어도 성공으로 돌아온다. 스풀러 쪽 상태를 따로 확인해야
# 실제로 프린터가 받았는지 알 수 있다.
PRINTER_ATTRIBUTE_WORK_OFFLINE = 0x00000400

PRINTER_PROBLEM_FLAGS = [
    (0x00000080, "프린터가 오프라인 상태입니다. 케이블과 전원을 확인해 주세요."),
    (0x00001000, "프린터를 사용할 수 없는 상태입니다. 연결을 확인해 주세요."),
    (0x00000002, "프린터에 오류가 발생했습니다."),
    (0x00000008, "용지가 걸렸습니다."),
    (0x00000010, "용지가 없습니다."),
    (0x00000040, "용지에 문제가 있습니다."),
    (0x00400000, "프린터 덮개가 열려 있습니다."),
    (0x00040000, "토너가 없습니다."),
    (0x00100000, "프린터에서 확인이 필요합니다."),
    (0x00200000, "프린터 메모리가 부족합니다."),
    (0x00000001, "프린터가 일시 중지되어 있습니다."),
]

JOB_PROBLEM_FLAGS = [
    (0x00000002, "인쇄 작업에 오류가 발생했습니다."),
    (0x00000020, "프린터가 오프라인 상태입니다. 케이블과 전원을 확인해 주세요."),
    (0x00000040, "용지가 없습니다."),
    (0x00000200, "프린터가 작업을 받지 못하고 있습니다. 연결을 확인해 주세요."),
    (0x00000400, "프린터에서 확인이 필요합니다."),
]

JOB_STATUS_DONE = 0x00000080 | 0x00001000   # PRINTED, COMPLETE


def get_printer_problem(printer_name):
    """
    프린터가 지금 인쇄할 수 없는 상태인지 확인한다.
    문제가 없으면 None, 있으면 안내 문구를 돌려준다.
    """
    if win32print is None or not printer_name:
        return None

    try:
        handle = win32print.OpenPrinter(printer_name)
    except Exception:
        return "프린터에 연결할 수 없습니다. 케이블과 전원을 확인해 주세요."

    try:
        info = win32print.GetPrinter(handle, 2)
    except Exception:
        return None
    finally:
        try:
            win32print.ClosePrinter(handle)
        except Exception:
            pass

    if not isinstance(info, dict):
        return None

    if info.get("Attributes", 0) & PRINTER_ATTRIBUTE_WORK_OFFLINE:
        return "프린터가 오프라인으로 설정되어 있습니다. 케이블과 전원을 확인해 주세요."

    status = info.get("Status", 0)

    for flag, message in PRINTER_PROBLEM_FLAGS:
        if status & flag:
            return message

    return None


def list_print_jobs(printer_name):
    """대기열에 있는 인쇄 작업 목록. 확인할 수 없으면 None."""
    if win32print is None or not printer_name:
        return None

    try:
        handle = win32print.OpenPrinter(printer_name)
    except Exception:
        return None

    try:
        return win32print.EnumJobs(handle, 0, 999, 1)
    except Exception:
        return None
    finally:
        try:
            win32print.ClosePrinter(handle)
        except Exception:
            pass


def snapshot_job_ids(printer_name):
    """지금 대기열에 있는 작업 번호들. 새로 들어온 작업을 가려내는 데 쓴다."""
    jobs = list_print_jobs(printer_name)

    if jobs is None:
        return None

    return {job.get("JobId") for job in jobs}


def check_print_job_result(printer_name, before_ids, timeout=8):
    """
    보낸 인쇄 작업이 대기열에서 제대로 처리되는지 잠시 지켜본다.

    돌려주는 값은 (상태, 안내문).
      "done"    — 작업이 대기열에서 빠졌거나 인쇄 완료로 바뀌었다 (프린터가 받아갔다)
      "unknown" — 확인하지 못했다 (아직 인쇄 중이거나 확인할 방법이 없음)
      "error"   — 프린터나 작업에 문제가 있다
    """
    if win32print is None or before_ids is None:
        return "unknown", None

    started = time.time()
    deadline = started + timeout
    seen_job = False

    # 작은 파일은 우리가 확인하기도 전에 인쇄가 끝나 대기열에서 사라진다.
    # 이런 경우까지 끝까지 기다리면 출력이 느려지므로, 잠깐만 보고 넘어간다.
    no_job_grace = 2.0

    while time.time() < deadline:
        problem = get_printer_problem(printer_name)

        if problem:
            return "error", problem

        jobs = list_print_jobs(printer_name)

        if jobs is None:
            return "unknown", None

        new_jobs = [j for j in jobs if j.get("JobId") not in before_ids]

        if new_jobs:
            seen_job = True

            for job in new_jobs:
                status = job.get("Status", 0)

                for flag, message in JOB_PROBLEM_FLAGS:
                    if status & flag:
                        return "error", message

                if status & JOB_STATUS_DONE:
                    return "done", None

        elif seen_job:
            # 대기열에서 빠졌다는 것은 프린터가 받아갔다는 뜻이다
            return "done", None

        elif time.time() - started > no_job_grace:
            # 작업이 보이지도 않고 프린터도 멀쩡하면 이미 끝난 것으로 본다
            return "unknown", None

        time.sleep(0.4)

    # 시간 안에 끝나지 않았어도, 프린터에 문제가 없으면 정상으로 본다
    # (쪽수가 많은 문서는 원래 오래 걸린다)
    problem = get_printer_problem(printer_name)

    return ("error", problem) if problem else ("unknown", None)


def get_default_printer():
    if win32print is None:
        return ""

    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


def add_column_if_missing(cur, table_name, column_name, column_sql):
    cur.execute(f"PRAGMA table_info({table_name})")
    columns = [row[1] for row in cur.fetchall()]

    if column_name not in columns:
        cur.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_sql}")


def init_db():
    conn = connect_db()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS students (
        student_id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        active INTEGER NOT NULL DEFAULT 1,
        unlimited INTEGER NOT NULL DEFAULT 0,
        created_date TEXT NOT NULL
    )
    """)

    add_column_if_missing(cur, "students", "unlimited", "unlimited INTEGER NOT NULL DEFAULT 0")
    add_column_if_missing(cur, "students", "student_number", "student_number TEXT")

    cur.execute("""
    CREATE INDEX IF NOT EXISTS idx_students_student_number
    ON students (student_number)
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS settings (
        setting_name TEXT PRIMARY KEY,
        setting_value TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS print_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id TEXT NOT NULL,
        student_name TEXT NOT NULL,
        file_name TEXT NOT NULL,
        print_date TEXT NOT NULL,
        print_time TEXT NOT NULL,
        printer_name TEXT NOT NULL,
        result TEXT NOT NULL
    )
    """)

    # 기존 DB에는 없을 수 있으므로 안전하게 컬럼을 추가한다.
    # sheet_count: 이 출력 건이 실제로 소모한 용지 장수 (양면이면 페이지수의 절반, 올림)
    add_column_if_missing(cur, "print_logs", "sheet_count", "sheet_count INTEGER NOT NULL DEFAULT 1")
    add_column_if_missing(cur, "print_logs", "duplex", "duplex INTEGER NOT NULL DEFAULT 0")

    # 관리자별 계정 (관리자 명단 CSV 에서 읽어 온다. 비밀번호와 코드는 해시로만 보관)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS admins (
        name TEXT PRIMARY KEY,
        password_hash TEXT NOT NULL DEFAULT '',
        code_hash TEXT NOT NULL DEFAULT '',
        note TEXT NOT NULL DEFAULT '',
        updated TEXT NOT NULL DEFAULT ''
    )
    """)

    # 관리자 화면에 '학번 이름 (직함)' 으로 보여주기 위한 칸
    add_column_if_missing(cur, "admins", "student_number", "student_number TEXT NOT NULL DEFAULT ''")
    add_column_if_missing(cur, "admins", "title", "title TEXT NOT NULL DEFAULT ''")

    # 관리자 작업 기록 (누가 · 언제 · 무엇을)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS admin_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        log_date TEXT NOT NULL,
        log_time TEXT NOT NULL,
        actor TEXT NOT NULL,
        action TEXT NOT NULL,
        detail TEXT NOT NULL DEFAULT ''
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS shutdown_schedules (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kind TEXT NOT NULL,
        days TEXT,
        run_date TEXT,
        run_time TEXT NOT NULL,
        created_date TEXT NOT NULL
    )
    """)

    default_settings = {
        "daily_limit": "3",
        "max_files_per_job": "3",
        "admin_password": hash_password(DEFAULT_ADMIN_PASSWORD),
        "admin_mode": "single",
        "idle_minutes": "5",
        "delete_mode": "print",
        "print_delete_minutes": "5",
        "selected_printer": get_default_printer(),
        "fullscreen": "0"
    }

    for key, value in default_settings.items():
        cur.execute(
            "INSERT OR IGNORE INTO settings(setting_name, setting_value) VALUES (?, ?)",
            (key, value)
        )

    conn.commit()
    conn.close()


def get_setting(name, default=""):
    conn = connect_db()
    cur = conn.cursor()
    cur.execute("SELECT setting_value FROM settings WHERE setting_name=?", (name,))
    row = cur.fetchone()
    conn.close()
    return row[0] if row else default


def set_setting(name, value):
    # 관리자가 설정을 바꾸면 무엇이 어떻게 바뀌었는지 기록한다
    before = get_setting(name, None) if admin_session_active() else None

    conn = connect_db()
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO settings(setting_name, setting_value)
    VALUES(?, ?)
    ON CONFLICT(setting_name) DO UPDATE SET setting_value=excluded.setting_value
    """, (name, str(value)))
    conn.commit()
    conn.close()

    if admin_session_active() and name in SETTING_LABELS and str(before) != str(value):
        label = SETTING_LABELS[name]

        if name in SECRET_SETTINGS:
            admin_log("설정 변경", f"{label} 변경")
        else:
            old_text = "(없음)" if before in (None, "") else describe_setting(name, before)
            admin_log("설정 변경", f"{label}: {old_text} → {describe_setting(name, value)}")


# -----------------------------
# 관리자 작업 기록
#
# 관리자 화면에 들어온 사람을 기억해 두고, 그 사이의 작업을 모두 남긴다.
#   · 단일 방식 : 이름 대신 '관리자' 로 남는다
#   · 개별 방식 : 관리자 명단의 이름으로 남는다
# -----------------------------
_admin_session = {"actor": None, "via": ""}

SETTING_LABELS = {
    "daily_limit": "하루 출력 횟수",
    "daily_page_limit": "하루 용지 장수",
    "limit_mode": "출력 제한 방식",
    "max_files_per_job": "한 번에 고를 파일 수",
    "max_copies": "파일당 최대 부수",
    "selected_printer": "프린터",
    "file_open_dir": "파일 선택 기본 폴더",
    "download_dir": "받은 파일 저장 폴더",
    "auto_delete_minutes": "받은 파일 보관 시간(분)",
    "delete_mode": "출력 후 삭제 방식",
    "print_delete_minutes": "출력 후 삭제 유예(분)",
    "fullscreen": "전체화면",
    "block_system_keys": "단축키 차단",
    "admin_password": "메인 비밀번호",
    "admin_barcode": "메인 관리자 코드",
    "admin_mode": "관리자 방식",
    "idle_minutes": "자리 비움 복귀(분)",
}

SECRET_SETTINGS = {"admin_password", "admin_barcode"}

_SETTING_VALUE_NAMES = {
    "limit_mode": {"count": "횟수", "pages": "용지 장수"},
    "delete_mode": {"time": "지우지 않음", "print": "바로 지움", "print_delay": "유예 후 지움"},
    "admin_mode": {"single": "단일 비밀번호", "individual": "관리자별 비밀번호"},
    "fullscreen": {"0": "끔", "1": "켬"},
    "block_system_keys": {"0": "끔", "1": "켬"},
}


def describe_setting(name, value):
    """기록에 남길 때 알아보기 쉬운 말로 바꾼다."""
    return _SETTING_VALUE_NAMES.get(name, {}).get(str(value), str(value))


def admin_session_active():
    return bool(_admin_session.get("actor"))


def current_admin():
    return _admin_session.get("actor") or ""


def admin_log_as(actor, action, detail=""):
    """누가 했는지 직접 정해서 기록한다. (로그인 실패처럼 아직 누구인지 모를 때)"""
    now = datetime.now()

    try:
        conn = connect_db()
        conn.execute(
            "INSERT INTO admin_logs(log_date, log_time, actor, action, detail) VALUES (?, ?, ?, ?, ?)",
            (now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S"), str(actor), str(action), str(detail))
        )
        conn.commit()
        conn.close()
    except Exception:
        # 기록이 실패해도 본래 작업은 계속되어야 한다
        pass


def admin_log(action, detail=""):
    """지금 관리자 화면에 들어와 있는 사람 이름으로 기록한다."""
    if not admin_session_active():
        return

    admin_log_as(current_admin(), action, detail)


def begin_admin_session(actor, via):
    _admin_session["actor"] = actor
    _admin_session["via"] = via

    shown = describe_admin(actor)
    admin_log("로그인", via + (f" · {shown}" if shown and shown != actor else ""))


def end_admin_session(reason="나감"):
    if not admin_session_active():
        return

    admin_log(reason)
    _admin_session["actor"] = None
    _admin_session["via"] = ""


def fetch_admin_logs(actor=None, limit=500):
    conn = connect_db()
    cur = conn.cursor()

    if actor:
        cur.execute(
            "SELECT log_date, log_time, actor, action, detail FROM admin_logs "
            "WHERE actor=? ORDER BY id DESC LIMIT ?", (actor, limit)
        )
    else:
        cur.execute(
            "SELECT log_date, log_time, actor, action, detail FROM admin_logs "
            "ORDER BY id DESC LIMIT ?", (limit,)
        )

    rows = cur.fetchall()
    conn.close()
    return rows


def list_admin_log_actors():
    conn = connect_db()
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT actor FROM admin_logs ORDER BY actor")
    rows = [r[0] for r in cur.fetchall()]
    conn.close()
    return rows


def export_admin_logs_to_csv(path, actor=None):
    rows = fetch_admin_logs(actor=actor, limit=1_000_000)

    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["날짜", "시각", "관리자", "작업", "내용"])
        for row in reversed(rows):
            writer.writerow(row)

    return len(rows)


# -----------------------------
# 관리자 명단 (관리자별 비밀번호 방식)
#
# 프로그램 폴더의 admin_list.csv 에 관리자를 적는다.
#   이름 · 학번 · 직함 · 관리자코드 · 비밀번호 · 비고
#   (학번 · 직함은 관리자 화면 위쪽에 '학번 이름 (직함)' 으로 표시된다. 비워도 된다)
#
# 비밀번호와 관리자코드는 파일에 오래 두면 안 되므로, 프로그램이 읽는 즉시
# 해시로 바꿔 DB 에 보관하고 파일의 그 칸은 '(등록됨)' 으로 바꿔 쓴다.
#   · 새로 정하거나 바꾸려면 → 그 칸에 새 값을 적고 저장
#   · 그대로 두려면        → '(등록됨)' 그대로 둔다
#   · 없애려면             → 칸을 비운다
# 파일에서 줄을 지우면 그 관리자는 더 이상 들어올 수 없다.
# -----------------------------
ADMIN_CSV_NAME = "admin_list.csv"
ADMIN_CSV_HEADER = ["이름", "학번", "직함", "관리자코드", "비밀번호", "비고"]
REGISTERED_MARK = "(등록됨)"

_admin_csv_state = {"stamp": None, "problems": []}


def get_admin_mode():
    mode = get_setting("admin_mode", "single")
    return mode if mode in ("single", "individual") else "single"


def get_admin_csv_path():
    return os.path.join(get_program_dir(), ADMIN_CSV_NAME)


def ensure_admin_csv():
    """관리자 명단 파일이 없으면 제목 줄만 있는 파일을 만든다."""
    path = get_admin_csv_path()

    if os.path.exists(path):
        return path

    try:
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerow(ADMIN_CSV_HEADER)
    except OSError:
        pass

    return path


def list_admins():
    conn = connect_db()
    cur = conn.cursor()
    cur.execute("SELECT name, password_hash, code_hash, note FROM admins ORDER BY name")
    rows = cur.fetchall()
    conn.close()
    return rows


def list_admin_profiles():
    """{이름: (학번, 직함)}"""
    conn = connect_db()
    cur = conn.cursor()
    cur.execute("SELECT name, student_number, title FROM admins")
    rows = {name: (number or "", title or "") for name, number, title in cur.fetchall()}
    conn.close()
    return rows


def describe_admin(name):
    """
    관리자를 '학번 이름 (직함)' 으로 나타낸다.
    학번이나 직함이 비어 있으면 그 부분은 뺀다.
    명단에 없는 이름(단일 방식의 '관리자' 등)은 그대로 돌려준다.
    """
    if not name:
        return ""

    number, title = list_admin_profiles().get(name, ("", ""))

    text = f"{number} {name}".strip()

    if title:
        text += f" ({title})"

    return text


def find_admin_by_password(password):
    if not password:
        return None

    for name, pw_hash, _, _ in list_admins():
        if pw_hash and verify_password(password, pw_hash):
            return name

    return None


def find_admin_by_code(code):
    code = normalize_card_code(code)

    if not code:
        return None

    for name, _, code_hash, _ in list_admins():
        if code_hash and verify_password(code, code_hash):
            return name

    return None


def _read_admin_csv_rows(path):
    for enc in ("utf-8-sig", "cp949"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                return list(csv.reader(f))
        except UnicodeDecodeError:
            continue
    raise ValueError("관리자 명단 파일의 글자 형식을 읽을 수 없습니다.")


def sync_admin_csv(force=False):
    """
    관리자 명단 파일을 읽어 DB 에 반영한다.
    파일이 바뀌지 않았으면 아무것도 하지 않는다. (3초마다 불려도 가볍다)
    돌려주는 값은 반영 여부.
    """
    path = ensure_admin_csv()
    stamp = _csv_stamp(path)

    if not force and stamp is not None and stamp == _admin_csv_state["stamp"]:
        return False

    try:
        rows = _read_admin_csv_rows(path)
    except Exception as e:
        _admin_csv_state["problems"] = [str(e)]
        _admin_csv_state["stamp"] = stamp
        return False

    if not rows:
        rows = [ADMIN_CSV_HEADER]

    header = [h.strip() for h in rows[0]]

    def col(label):
        return header.index(label) if label in header else -1

    i_name, i_code, i_pw, i_note = col("이름"), col("관리자코드"), col("비밀번호"), col("비고")
    i_number, i_title = col("학번"), col("직함")

    if i_name < 0:
        _admin_csv_state["problems"] = ["첫 줄에 '이름' 칸이 없습니다. 제목 줄을 확인하세요."]
        _admin_csv_state["stamp"] = stamp
        return False

    existing = {name: (pw, code, note) for name, pw, code, note in list_admins()}
    old_profiles = list_admin_profiles()
    profiles = {}
    result = {}
    problems = []
    rewritten = [ADMIN_CSV_HEADER]
    new_codes = {}
    new_pws = {}

    def cell(row, idx):
        return row[idx].strip() if 0 <= idx < len(row) else ""

    for row in rows[1:]:
        name = cell(row, i_name)

        if not name:
            continue

        if name in result:
            problems.append(f"'{name}' 이(가) 두 번 적혀 있어 뒤의 줄은 무시했습니다.")
            continue

        old_pw, old_code, _ = existing.get(name, ("", "", ""))
        pw_cell, code_cell = cell(row, i_pw), cell(row, i_code)
        note = cell(row, i_note)
        profiles[name] = (cell(row, i_number), cell(row, i_title))

        # 비밀번호
        if pw_cell == REGISTERED_MARK:
            pw_hash = old_pw
        elif pw_cell == "":
            pw_hash = ""
        else:
            pw_hash = old_pw if (old_pw and verify_password(pw_cell, old_pw)) else hash_password(pw_cell)
            new_pws[name] = pw_cell

        # 관리자코드
        if code_cell == REGISTERED_MARK:
            code_hash = old_code
        elif code_cell == "":
            code_hash = ""
        else:
            plain = normalize_card_code(code_cell)

            if get_student(plain):
                problems.append(f"'{name}' 의 관리자코드가 학생증 코드·학번과 겹쳐 등록하지 않았습니다.")
                code_hash = old_code
            else:
                code_hash = old_code if (old_code and verify_password(plain, old_code)) else hash_password(plain)
                new_codes[name] = plain

        result[name] = [pw_hash, code_hash, note]
        rewritten.append([
            name,
            profiles[name][0],
            profiles[name][1],
            REGISTERED_MARK if code_hash else "",
            REGISTERED_MARK if pw_hash else "",
            note,
        ])

    # 같은 비밀번호 · 같은 코드를 두 사람이 쓰면 누구인지 가릴 수 없다.
    # 새로 적은 값이 다른 사람의 값과 같으면 새로 적은 쪽을 받지 않는다.
    # 둘 다 이번에 새로 적었으면 파일에서 먼저 나온 사람을 살린다.
    order = list(result.keys())

    def dedupe(plain_map, index, label):
        for pos, name in enumerate(order):
            if name not in plain_map:
                continue

            plain = plain_map[name]

            for other in order:
                if other == name:
                    continue

                other_hash = result[other][index]

                if not other_hash:
                    continue

                # 뒤에 나온 사람이 이번에 새로 적은 값이면, 그쪽이 양보한다
                if other in plain_map and order.index(other) > pos:
                    continue

                if verify_password(plain, other_hash):
                    result[name][index] = existing.get(name, ("", "", ""))[index]
                    problems.append(f"'{name}' 의 {label}이(가) '{other}' 와 같아 등록하지 않았습니다.")
                    break

    dedupe(new_pws, 0, "비밀번호")
    dedupe(new_codes, 1, "관리자코드")

    # 반영 (명단에서 빠진 관리자는 지운다 → 더 이상 들어올 수 없다)
    changes = []
    conn = connect_db()
    cur = conn.cursor()
    today = str(date.today())

    for name, (pw_hash, code_hash, note) in result.items():
        old = existing.get(name)
        if old is None:
            changes.append(f"추가: {name}")
        else:
            number, title = profiles.get(name, ("", ""))
            old_number, old_title = old_profiles.get(name, ("", ""))
            parts = []
            if old[0] != pw_hash: parts.append("비밀번호")
            if old[1] != code_hash: parts.append("관리자코드")
            if old_number != number: parts.append("학번")
            if old_title != title: parts.append("직함")
            if old[2] != note: parts.append("비고")
            if parts:
                changes.append(f"변경: {name} ({', '.join(parts)})")

        number, title = profiles.get(name, ("", ""))

        cur.execute("""
        INSERT INTO admins(name, password_hash, code_hash, note, updated, student_number, title)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(name) DO UPDATE SET password_hash=excluded.password_hash,
            code_hash=excluded.code_hash, note=excluded.note, updated=excluded.updated,
            student_number=excluded.student_number, title=excluded.title
        """, (name, pw_hash, code_hash, note, today, number, title))

    for name in existing:
        if name not in result:
            cur.execute("DELETE FROM admins WHERE name=?", (name,))
            changes.append(f"삭제: {name}")

    conn.commit()
    conn.close()

    # 파일의 비밀번호 · 코드 칸을 '(등록됨)' 으로 바꿔 쓴다 (평문이 파일에 남지 않도록)
    for r in rewritten[1:]:
        name = r[0]
        r[3] = REGISTERED_MARK if result[name][1] else ""
        r[4] = REGISTERED_MARK if result[name][0] else ""

    try:
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerows(rewritten)
    except OSError:
        problems.append("명단 파일이 다른 프로그램(엑셀 등)에서 열려 있어 비밀번호 칸을 가리지 못했습니다. 파일을 닫아 주세요.")

    _admin_csv_state["stamp"] = _csv_stamp(path)
    _admin_csv_state["problems"] = problems

    if changes:
        actor = current_admin() or "(명단 파일)"
        admin_log_as(actor, "관리자 명단 반영", " / ".join(changes))

    return True


# 제한값이 0 이면 "제한 없음"을 뜻한다.
UNLIMITED = 0


def get_daily_limit():
    """하루 출력 가능한 파일 개수. 0 이면 제한 없음."""
    try:
        return int(get_setting("daily_limit", "3"))
    except ValueError:
        return 3


def get_daily_page_limit():
    """하루 사용 가능한 용지 장수. 0 이면 제한 없음."""
    try:
        return int(get_setting("daily_page_limit", "20"))
    except ValueError:
        return 20


def get_limit_mode():
    """
    하루 출력 제한을 무엇으로 셀지.
      "files" - 파일(출력 건) 개수로 제한 (기존 방식)
      "pages" - 용지 장수로 제한
    """
    mode = get_setting("limit_mode", "files")
    return mode if mode in ("files", "pages") else "files"


def get_max_files_per_job():
    """한 번에 고를 수 있는 파일 개수. 0 이면 제한 없음."""
    try:
        return int(get_setting("max_files_per_job", "3"))
    except ValueError:
        return 3


# 받은 파일을 언제 지울지 고르는 방식
#   "time"        보관 시간이 지나면 지운다 (출력 여부와 상관없음)
#   "print"       출력이 끝나면 바로 지운다
#   "print_delay" 출력이 끝나고 정해둔 시간이 지나면 지운다
DELETE_MODES = ("time", "print", "print_delay")


def get_delete_mode():
    mode = get_setting("delete_mode", "print")
    return mode if mode in DELETE_MODES else "print"


def get_print_delete_minutes():
    """출력이 끝난 뒤 몇 분 뒤에 지울지. (출력 후 유예 삭제를 고른 경우)"""
    try:
        return max(0, int(get_setting("print_delete_minutes", "5")))
    except ValueError:
        return 5


def get_idle_minutes():
    """
    아무 조작이 없을 때 몇 분 뒤에 처음 화면으로 돌아갈지. 0 이면 돌아가지 않는다.
    """
    try:
        return int(get_setting("idle_minutes", "5"))
    except ValueError:
        return 5


def get_max_copies():
    """파일 하나당 뽑을 수 있는 최대 부수. 0 이면 제한 없음."""
    try:
        return int(get_setting("max_copies", "5"))
    except ValueError:
        return 5


def get_download_dir():
    """
    파일 공유 기능을 통해 파일을 저장할 폴더.
    설정하지 않았으면 프로그램 폴더 아래 ReceivedFiles 를 쓴다.
    """
    saved = get_setting("download_dir", "").strip()

    if saved:
        return saved

    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        try:
            base = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            base = os.getcwd()

    return os.path.join(base, "ReceivedFiles")


DEFAULT_AUTO_DELETE_MINUTES = 30


def get_auto_delete_minutes():
    """
    받은 파일을 받은 시점부터 몇 분 뒤에 지울지. 0 이면 시간으로는 지우지 않는다.

    출력 후 삭제와 함께 쓰는 값이다. 학생이 파일만 보내놓고 뽑지 않고 가버리는
    일이 흔해서, 출력하지 않은 파일은 이 시간으로 정리된다.
    """
    try:
        return max(0, int(get_setting("auto_delete_minutes", str(DEFAULT_AUTO_DELETE_MINUTES))))
    except (TypeError, ValueError):
        return DEFAULT_AUTO_DELETE_MINUTES


# 다운로드가 끝난 파일을 감시하는 목록: {파일경로: 완료된 시각}
_download_watch = {}
_download_watch_lock = threading.Lock()


def is_download_in_progress(path):
    """
    크롬이 아직 내려받는 중인 파일인지 확인한다.

    크롬은 받는 동안 '파일명.crdownload' 로 저장하다가
    완료되면 원래 이름으로 바꾼다.
    """
    if path.endswith(".crdownload") or path.endswith(".tmp"):
        return True

    # 같은 이름의 .crdownload 가 남아 있으면 아직 진행 중이다
    if os.path.exists(path + ".crdownload"):
        return True

    return False


# 출력이 끝나 "언제 지울지" 를 예약해 둔 파일들. {경로: 지울 시각}
_delete_after_print = {}


def schedule_delete_after_print(path, minutes):
    """출력이 끝난 파일을 정해진 시간 뒤에 지우도록 예약한다."""
    with _download_watch_lock:
        _delete_after_print[os.path.abspath(path)] = time.time() + minutes * 60


def scan_download_folder():
    """
    다운로드 폴더를 훑어서 새로 완료된 파일을 감시 목록에 넣고,
    지울 때가 된 파일을 지운다.

    지우는 기준은 두 가지이며, 둘 중 하나라도 해당하면 지운다.
      · 보관 시간 — 받은 시점부터 정해진 시간이 지난 파일
      · 출력 후 유예 — 출력이 끝나 예약해 둔 시각이 지난 파일
    """
    minutes = get_auto_delete_minutes()
    folder = get_download_dir()

    if not os.path.isdir(folder):
        return

    now = time.time()
    limit_seconds = minutes * 60

    try:
        names = os.listdir(folder)
    except Exception:
        return

    with _download_watch_lock:
        current = set()

        for name in names:
            path = os.path.join(folder, name)

            if not os.path.isfile(path):
                continue

            # 아직 받는 중인 파일은 건너뛴다
            if is_download_in_progress(path):
                continue

            current.add(path)

            # 처음 본 파일이면 이 시점을 완료 시각으로 기록한다
            if path not in _download_watch:
                _download_watch[path] = now

        # 사라진 파일은 목록에서 정리
        for path in list(_download_watch.keys()):
            if path not in current:
                del _download_watch[path]

        for path in list(_delete_after_print.keys()):
            if not os.path.exists(path):
                del _delete_after_print[path]

        # 지울 때가 된 파일 삭제
        for path, done_at in list(_download_watch.items()):
            due = False

            # 보관 시간 기준 (출력하지 않고 방치된 파일도 이때 지워진다)
            if minutes > 0 and now - done_at >= limit_seconds:
                due = True

            # 출력 후 유예 기준
            deadline = _delete_after_print.get(os.path.abspath(path))

            if deadline is not None and now >= deadline:
                due = True

            if not due:
                continue

            try:
                os.remove(path)
            except Exception:
                pass
            else:
                del _download_watch[path]
                _delete_after_print.pop(os.path.abspath(path), None)


def get_student(card_code):
    """
    학생증 코드로 학생을 조회한다.
    일치하는 학생증 코드가 없으면 같은 값으로 학번 조회를 한 번 더 시도한다.
    (학번으로도 조회가 되는 것은 화면에 안내되지 않는 숨겨진 기능이다.)
    """
    normalized = normalize_card_code(card_code)

    conn = connect_db()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT student_id, name, active, unlimited, created_date
        FROM students
        WHERE student_id=?
        """,
        (normalized,)
    )
    row = cur.fetchone()

    if row is None:
        cur.execute(
            """
            SELECT student_id, name, active, unlimited, created_date
            FROM students
            WHERE student_number=?
            """,
            (normalized,)
        )
        row = cur.fetchone()

    conn.close()
    return row


def get_student_number(card_code):
    """학생증 코드로 학번을 가져온다. 없으면 빈 문자열을 돌려준다."""
    conn = connect_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT student_number FROM students WHERE student_id=?",
        (normalize_card_code(card_code),)
    )
    row = cur.fetchone()
    conn.close()

    if not row or not row[0]:
        return ""

    return row[0]


def get_student_by_number(student_number):
    student_number = normalize_card_code(student_number)

    conn = connect_db()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT student_id, name, active, unlimited, created_date
        FROM students
        WHERE student_number=?
        """,
        (student_number,)
    )
    row = cur.fetchone()
    conn.close()
    return row


def get_today_count(card_code):
    """
    오늘 이미 쓴 만큼을 돌려준다.
    limit_mode 가 "files" 면 출력 건수, "pages" 면 소모한 용지 장수 합계.
    """
    today = str(date.today())
    card_code = normalize_card_code(card_code)

    conn = connect_db()
    cur = conn.cursor()

    if get_limit_mode() == "pages":
        cur.execute(
            "SELECT COALESCE(SUM(sheet_count), 0) FROM print_logs "
            "WHERE student_id=? AND print_date=? AND result!='실패'",
            (card_code, today)
        )
    else:
        cur.execute(
            "SELECT COUNT(*) FROM print_logs WHERE student_id=? AND print_date=? AND result!='실패'",
            (card_code, today)
        )

    count = cur.fetchone()[0]
    conn.close()
    return count


def get_today_limit_value():
    """현재 모드에 맞는 하루 제한값. 0 이면 제한 없음."""
    return get_daily_page_limit() if get_limit_mode() == "pages" else get_daily_limit()


def is_unlimited_limit():
    """지금 모드의 하루 제한이 '제한 없음' 인지."""
    return get_today_limit_value() <= UNLIMITED


def save_print_log(card_code, name, file_name, printer_name, result, sheet_count=1, duplex=False):
    now = datetime.now()
    card_code = normalize_card_code(card_code)

    conn = connect_db()
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO print_logs
    (student_id, student_name, file_name, print_date, print_time, printer_name, result, sheet_count, duplex)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        card_code,
        name,
        file_name,
        str(date.today()),
        now.strftime("%H:%M:%S"),
        printer_name,
        result,
        sheet_count,
        1 if duplex else 0
    ))
    conn.commit()
    conn.close()


def upsert_student(card_code, name, student_number="", active=1, unlimited=0):
    card_code = normalize_card_code(card_code)
    name = str(name).strip()
    student_number = normalize_card_code(student_number)

    if not card_code:
        raise ValueError("학생증 코드가 비어 있습니다.")

    if not name:
        raise ValueError("이름이 비어 있습니다.")

    conn = connect_db()
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO students(student_id, name, student_number, active, unlimited, created_date)
    VALUES (?, ?, ?, ?, ?, ?)
    ON CONFLICT(student_id) DO UPDATE SET
        name=excluded.name,
        student_number=excluded.student_number,
        active=excluded.active,
        unlimited=excluded.unlimited
    """, (
        card_code,
        name,
        student_number,
        int(active),
        int(unlimited),
        str(date.today())
    ))
    conn.commit()
    conn.close()


# -----------------------------
# 파일 변환 / 프린터 출력 관련 함수
# -----------------------------
def get_printer_list():
    """
    설치된 프린터 목록을 반환한다.
    문제가 있을 때 원인을 알 수 있도록 예외를 그대로 올려보낸다
    (예: pywin32 미설치, 관리자 권한 문제 등).
    """
    if win32print is None:
        raise RuntimeError(
            "pywin32(win32print) 모듈이 설치되어 있지 않습니다.\n"
            "명령 프롬프트에서 'pip install pywin32'를 실행한 뒤 다시 시도하세요."
        )

    printers = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )
    return [printer[2] for printer in printers]


def find_sumatra_pdf():
    candidates = [
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe"),
    ]

    for path in candidates:
        if os.path.exists(path):
            return path

    found = shutil.which("SumatraPDF.exe")
    if found:
        return found

    return None


def find_libreoffice():
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    for path in candidates:
        if os.path.exists(path):
            return path

    found = shutil.which("soffice.exe")
    if found:
        return found

    return None


CONVERTED_DIR_NAME = "school_printer_converted"
PREVIEW_DIR_NAME = "school_printer_preview"


def get_converted_dir():
    out_dir = os.path.join(tempfile.gettempdir(), CONVERTED_DIR_NAME)
    os.makedirs(out_dir, exist_ok=True)
    return out_dir


def file_key(path):
    """
    파일을 구분하는 짧은 열쇠.

    파일 이름만으로 임시 파일 이름을 지으면, 서로 다른 학생이 같은 이름의 파일
    (예: 과제.hwp)을 냈을 때 먼저 만든 변환 결과가 재사용되어 엉뚱한 내용이
    인쇄될 수 있다. 전체 경로와 수정시각, 크기를 섞어 열쇠를 만든다.
    """
    import hashlib

    try:
        info = os.stat(path)
        raw = f"{os.path.abspath(path)}|{int(info.st_mtime)}|{info.st_size}"
    except OSError:
        raw = os.path.abspath(path)

    return hashlib.md5(raw.encode("utf-8", errors="replace")).hexdigest()[:12]


def safe_stem(path, limit=40):
    base_name = os.path.splitext(os.path.basename(path))[0]
    safe = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in base_name)
    return safe[:limit] or "file"


def clean_temp_dirs(max_age_hours=6):
    """
    오래된 변환·미리보기 임시 파일을 지운다.
    프로그램을 켤 때 한 번 돌려서 임시 폴더가 계속 불어나지 않게 한다.
    """
    limit = max_age_hours * 3600
    now = time.time()

    for name in (CONVERTED_DIR_NAME, PREVIEW_DIR_NAME):
        folder = os.path.join(tempfile.gettempdir(), name)

        if not os.path.isdir(folder):
            continue

        for file_name in os.listdir(folder):
            path = os.path.join(folder, file_name)
            try:
                if os.path.isfile(path) and now - os.path.getmtime(path) > limit:
                    os.remove(path)
            except OSError:
                pass


def make_temp_pdf_path(original_path):
    out_dir = get_converted_dir()
    return os.path.join(
        out_dir, f"{safe_stem(original_path)}_{file_key(original_path)}.pdf"
    )


def send_pdf_to_printer(file_path, printer_name, duplex=False, page_range=None, copies=1):
    """
    SumatraPDF 로 인쇄 명령을 보낸다.

    -print-settings 로 양면(duplex) / 단면(simplex), 페이지 범위, 부수를 지정한다.
    프린터가 자동 양면 유닛을 지원해야 실제로 양면으로 나온다.
    (지원하지 않는 프린터는 이 옵션을 무시하거나 오류를 낼 수 있다)
    """
    sumatra = find_sumatra_pdf()

    if not sumatra:
        raise RuntimeError("SumatraPDF.exe를 찾을 수 없습니다.")

    if not printer_name:
        raise RuntimeError("선택된 프린터가 없습니다. 관리자 화면에서 프린터를 선택하세요.")

    if not printer_exists(printer_name):
        raise RuntimeError(
            f"'{printer_name}' 프린터를 찾을 수 없습니다.\n"
            "프린터가 켜져 있는지 확인하고, 관리자 화면에서 프린터를 다시 선택하세요."
        )

    # 보내기 전에 프린터 상태부터 확인한다.
    # 케이블이 빠졌거나 용지가 없으면 여기서 걸러진다.
    problem = get_printer_problem(printer_name)

    if problem:
        raise RuntimeError(problem)

    settings = ["duplex" if duplex else "simplex"]

    # 특정 페이지만 인쇄할 때는 "1-3,5" 형태를 함께 넘긴다
    if page_range:
        settings.insert(0, page_range)

    # 여러 부 인쇄는 "3x" 형태로 지정한다
    try:
        copies = int(copies)
    except (TypeError, ValueError):
        copies = 1

    if copies > 1:
        settings.append(f"{copies}x")

    cmd = [
        sumatra,
        "-print-to",
        printer_name,
        "-print-settings",
        ",".join(settings),
        "-silent",
        file_path
    ]

    # 지금 대기열에 있는 작업 번호를 기억해 둔다.
    # 이걸 알아야 방금 보낸 작업만 골라 상태를 확인할 수 있다.
    before_ids = snapshot_job_ids(printer_name)

    # Popen 으로 던져만 두면 실패해도 알 수 없어서, 프린터가 꺼져 있어도
    # "전송완료" 로 기록된다. 결과를 기다렸다가 실패면 예외를 낸다.
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            timeout=120,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(
            "프린터로 보내는 데 시간이 너무 오래 걸려 중단했습니다.\n"
            "프린터 상태와 대기 중인 인쇄 작업을 확인해 주세요."
        )

    if result.returncode != 0:
        message = (result.stderr or result.stdout or b"").decode("cp949", errors="replace").strip()
        raise RuntimeError(
            "프린터로 보내지 못했습니다."
            + (f"\n{message}" if message else "")
        )

    # SumatraPDF 가 알려주는 것은 "인쇄 대기열에 넣는 데 성공했다" 까지다.
    # 케이블이 빠져 있어도 대기열에는 들어가므로 여기서 끝내면 안 된다.
    # 방금 넣은 작업이 프린터로 넘어가는지 잠시 지켜본다.
    status, problem = check_print_job_result(printer_name, before_ids)

    if status == "error":
        raise RuntimeError(
            f"{problem}\n\n"
            "인쇄 작업이 대기열에 남아 있을 수 있으니, 문제를 해결한 뒤\n"
            "다시 출력하기 전에 대기 중인 작업을 정리해 주세요."
        )

    return status


# LibreOffice 는 같은 프로필을 쓰기 때문에 동시에 여러 번 실행하면
# 서로 충돌해서 변환에 실패하거나 멈춘다. 한 번에 하나씩만 돌린다.
_libreoffice_lock = threading.Lock()


def convert_libreoffice_to_pdf(file_path):
    soffice = find_libreoffice()

    if not soffice:
        raise RuntimeError("LibreOffice를 찾을 수 없습니다.")

    out_dir = get_converted_dir()

    # 이미 변환해 둔 결과가 있으면 다시 변환하지 않는다.
    #
    # 예전에는 파일 이름만으로 캐시를 찾았기 때문에, 서로 다른 학생이 같은 이름의
    # 파일(예: 과제.hwp)을 내면 앞사람의 변환 결과가 그대로 인쇄될 수 있었다.
    # 이제는 경로·수정시각·크기를 섞은 열쇠로 구분한다.
    cached_path = os.path.join(
        out_dir, f"{safe_stem(file_path)}_{file_key(file_path)}.pdf"
    )

    if os.path.exists(cached_path):
        return cached_path

    with _libreoffice_lock:
        # 기다리는 동안 다른 스레드가 이미 변환했을 수 있다
        if os.path.exists(cached_path):
            return cached_path

        produced = _run_libreoffice_convert(soffice, file_path, out_dir)

        # LibreOffice 는 "원본이름.pdf" 로 내보내므로, 겹치지 않는 이름으로 옮긴다
        if os.path.abspath(produced) != os.path.abspath(cached_path):
            try:
                if os.path.exists(cached_path):
                    os.remove(cached_path)
                os.replace(produced, cached_path)
            except OSError:
                return produced

        return cached_path


def _run_libreoffice_convert(soffice, file_path, out_dir):
    before_files = set(os.listdir(out_dir))

    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nodefault",
        "--convert-to",
        "pdf",
        "--outdir",
        out_dir,
        file_path
    ]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(
            "파일 변환이 너무 오래 걸려 중단했습니다.\n"
            "파일이 너무 크거나 손상되었을 수 있습니다."
        )

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    expected_path = os.path.join(out_dir, base_name + ".pdf")

    if os.path.exists(expected_path):
        return expected_path

    after_files = set(os.listdir(out_dir))
    new_files = [f for f in after_files - before_files if f.lower().endswith(".pdf")]

    if new_files:
        new_files.sort(key=lambda f: os.path.getmtime(os.path.join(out_dir, f)), reverse=True)
        return os.path.join(out_dir, new_files[0])

    raise RuntimeError(
        "파일을 PDF로 변환하지 못했습니다.\n\n"
        f"오류 내용:\n{result.stderr or result.stdout}"
    )


def convert_image_to_pdf(file_path):
    if Image is None:
        raise RuntimeError("이미지 출력을 위해 pillow가 필요합니다.")

    pdf_path = make_temp_pdf_path(file_path)

    with Image.open(file_path) as img:
        if img.mode in ("RGBA", "LA"):
            background = Image.new("RGB", img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[-1])
            background.save(pdf_path, "PDF", resolution=100.0)
        else:
            img.convert("RGB").save(pdf_path, "PDF", resolution=100.0)

    return pdf_path


def print_hwp_by_default_program(file_path, printer_name):
    if win32print is None:
        raise RuntimeError("pywin32가 필요합니다.")

    old_printer = get_default_printer()

    try:
        if printer_name:
            win32print.SetDefaultPrinter(printer_name)

        os.startfile(file_path, "print")
        time.sleep(3)

    finally:
        try:
            if old_printer:
                win32print.SetDefaultPrinter(old_printer)
        except Exception:
            pass


def get_pdf_path_for_counting(file_path):
    """
    페이지 수를 세기 위해 PDF 경로를 얻는다.
    이미 PDF면 그대로, 워드/한글오피스 계열이면 미리 변환해서 그 결과를 쓴다.

    실제 인쇄 시 print_document 가 다시 변환하지 않도록,
    변환 결과 경로를 호출한 쪽에서 재사용하는 것이 좋다.
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".pdf":
        return file_path

    if ext in LIBREOFFICE_EXTENSIONS:
        return convert_libreoffice_to_pdf(file_path)

    if ext in IMAGE_EXTENSIONS:
        return convert_image_to_pdf(file_path)

    # HWP/HWPX 는 변환 경로가 없어 페이지 수를 알 수 없다 (아래 count_pages_for_print 참고)
    return None


def count_pages_for_print(file_path, converted_cache=None):
    """
    파일 하나의 실제 페이지 수를 센다.

    converted_cache 를 넘기면 {원본경로: 변환된PDF경로} 로 결과를 저장해서
    나중에 인쇄할 때 같은 변환을 또 하지 않도록 재사용할 수 있다.

    한글(HWP) 파일은 LibreOffice 한글 확장 프로그램이 설치된 경우
    PDF로 변환하여 페이지 수를 알 수 있다.
    """
    ext = os.path.splitext(file_path)[1].lower()

    pdf_path = get_pdf_path_for_counting(file_path)

    if converted_cache is not None and pdf_path != file_path:
        converted_cache[file_path] = pdf_path

    if fitz is None:
        return None

    try:
        doc = fitz.open(pdf_path)
        try:
            return doc.page_count
        finally:
            doc.close()
    except Exception:
        return None


def sheets_from_pages(pages, duplex):
    """양면이면 2페이지가 1장, 단면이면 1페이지가 1장. 올림 처리."""
    if pages is None:
        return None

    if duplex:
        return (pages + 1) // 2

    return pages


def print_document(file_path, printer_name, duplex=False, converted_cache=None, page_range=None, copies=1):
    """
    돌려주는 값은 (상태, 실제로 인쇄한 PDF 경로).
    상태는 check_print_job_result 와 같다. ("done" 이면 프린터가 받아간 것이 확인된 것)
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".pdf":
        status = send_pdf_to_printer(file_path, printer_name, duplex=duplex,
                                     page_range=page_range, copies=copies)
        return status, file_path

    if ext in LIBREOFFICE_EXTENSIONS:
        # 페이지 수를 셀 때 이미 변환해뒀다면 그 결과를 재사용한다
        # (LibreOffice 변환은 느려서 같은 파일을 두 번 변환하면 그만큼 늦어진다)
        cached = converted_cache.get(file_path) if converted_cache else None
        pdf_path = cached if cached and os.path.exists(cached) else convert_libreoffice_to_pdf(file_path)
        status = send_pdf_to_printer(pdf_path, printer_name, duplex=duplex,
                                     page_range=page_range, copies=copies)
        return status, pdf_path

    if ext in IMAGE_EXTENSIONS:
        cached = converted_cache.get(file_path) if converted_cache else None
        pdf_path = cached if cached and os.path.exists(cached) else convert_image_to_pdf(file_path)
        status = send_pdf_to_printer(pdf_path, printer_name, duplex=duplex,
                                     page_range=page_range, copies=copies)
        return status, pdf_path

    raise RuntimeError("지원하지 않는 파일 형식입니다.")


def is_in_download_dir(file_path):
    """파일 공유로 받은 폴더 안에 있는 파일인지."""
    try:
        download_dir = os.path.normcase(os.path.abspath(get_download_dir()))
        target = os.path.normcase(os.path.abspath(file_path))

        # 드라이브가 다르면 commonpath 가 예외를 낸다 (D: 의 USB 파일 등)
        return os.path.commonpath([download_dir, target]) == download_dir
    except Exception:
        return False


def purge_preview_images(pdf_path):
    """그 파일로 만든 미리보기 썸네일을 지운다."""
    folder = os.path.join(tempfile.gettempdir(), PREVIEW_DIR_NAME)

    if not os.path.isdir(folder):
        return

    prefix = f"{safe_stem(pdf_path)}_{file_key(pdf_path)}"

    for name in os.listdir(folder):
        if name.startswith(prefix):
            try:
                os.remove(os.path.join(folder, name))
            except OSError:
                pass


def cleanup_after_print(file_path, printed_pdf_path):
    """
    출력이 끝난 뒤 남은 흔적을 지운다.

    프린터가 작업을 받아간 것이 확인됐을 때만 부른다.
    지우는 것은 아래 세 가지이며, 학생이 직접 고른 자기 파일(바탕화면, USB 등)은
    건드리지 않는다.
      · 파일 공유로 받은 원본 (받은 파일 폴더 안에 있는 것만)
      · 인쇄를 위해 만든 변환본 PDF
      · 미리보기 썸네일

    받은 원본을 언제 지울지는 출력 파일 설정의 "받은 파일 삭제 시점" 을 따른다.
    돌려주는 값은 ("deleted" / "scheduled" / None, 유예 분).
    """
    # 미리보기와 변환본은 임시 파일이므로 언제든 지워도 된다
    for path in {file_path, printed_pdf_path}:
        try:
            purge_preview_images(path)
        except Exception:
            pass

    if printed_pdf_path and os.path.abspath(printed_pdf_path) != os.path.abspath(file_path):
        try:
            if os.path.abspath(printed_pdf_path).startswith(os.path.abspath(get_converted_dir())):
                os.remove(printed_pdf_path)
        except OSError:
            pass

    if not is_in_download_dir(file_path):
        return None, 0

    mode = get_delete_mode()

    # 보관 시간만 쓰는 방식이면 출력과 상관없이 그대로 둔다
    if mode == "time":
        return None, 0

    if mode == "print_delay":
        minutes = get_print_delete_minutes()

        if minutes > 0:
            schedule_delete_after_print(file_path, minutes)
            return "scheduled", minutes

    # 즉시 삭제 (유예 시간을 0 으로 둔 경우도 여기로 온다)
    try:
        os.remove(file_path)
        return "deleted", 0
    except OSError:
        return None, 0


def is_supported_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    return ext in SUPPORTED_EXTENSIONS


# -----------------------------
# 출력 미리보기 관련 함수
# -----------------------------
def render_pdf_page_thumbnail(pdf_path, page_index, max_width=190):
    """
    PDF 의 특정 페이지를 작은 이미지로 만들어 경로를 돌려준다.
    페이지 선택 창의 썸네일용이라 해상도를 낮게 잡아 빠르게 만든다.
    """
    if fitz is None:
        raise RuntimeError("PyMuPDF 패키지가 필요합니다.")

    doc = fitz.open(pdf_path)
    try:
        if page_index < 0 or page_index >= doc.page_count:
            raise RuntimeError("없는 페이지입니다.")

        page = doc.load_page(page_index)
        zoom = max_width / page.rect.width
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))

        out_dir = os.path.join(tempfile.gettempdir(), PREVIEW_DIR_NAME)
        os.makedirs(out_dir, exist_ok=True)

        # 이름이 같은 다른 파일의 썸네일이 재사용되지 않도록 열쇠를 붙인다
        thumb_path = os.path.join(
            out_dir,
            f"{safe_stem(pdf_path)}_{file_key(pdf_path)}_p{page_index + 1}.png"
        )

        pix.save(thumb_path)
        return thumb_path
    finally:
        doc.close()


def parse_page_range(text, total_pages):
    """
    "1-3,5,8-10" 같은 문자열을 페이지 번호 집합으로 바꾼다.
    잘못된 형식이거나 범위를 벗어나면 ValueError 를 낸다.
    빈 문자열은 전체 페이지를 뜻한다.
    """
    text = (text or "").strip()

    if not text:
        return set(range(1, total_pages + 1))

    pages = set()

    for chunk in text.replace(" ", "").split(","):
        if not chunk:
            continue

        if "-" in chunk:
            parts = chunk.split("-")
            if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
                raise ValueError(f"'{chunk}' 형식을 알 수 없습니다.")

            start, end = int(parts[0]), int(parts[1])
            if start > end:
                start, end = end, start

            if start < 1 or end > total_pages:
                raise ValueError(f"'{chunk}' 는 1~{total_pages} 범위를 벗어납니다.")

            pages.update(range(start, end + 1))
        else:
            if not chunk.isdigit():
                raise ValueError(f"'{chunk}' 형식을 알 수 없습니다.")

            num = int(chunk)
            if num < 1 or num > total_pages:
                raise ValueError(f"'{num}' 는 1~{total_pages} 범위를 벗어납니다.")

            pages.add(num)

    if not pages:
        raise ValueError("선택된 페이지가 없습니다.")

    return pages


def format_page_range(pages, total_pages):
    """
    페이지 번호 집합을 "1-3,5,8-10" 형태의 문자열로 만든다.
    전체 선택이면 빈 문자열을 돌려준다 (= 전체 출력).
    """
    if not pages:
        return ""

    if len(pages) == total_pages:
        return ""

    result = []
    ordered = sorted(pages)
    start = prev = ordered[0]

    for num in ordered[1:]:
        if num == prev + 1:
            prev = num
            continue

        result.append(str(start) if start == prev else f"{start}-{prev}")
        start = prev = num

    result.append(str(start) if start == prev else f"{start}-{prev}")
    return ",".join(result)


def render_pdf_first_page_image(pdf_path, max_width=720):
    if fitz is None:
        raise RuntimeError(
            "미리보기를 위해 PyMuPDF 패키지가 필요합니다.\n'pip install pymupdf'로 설치하세요."
        )

    doc = fitz.open(pdf_path)
    try:
        if doc.page_count == 0:
            raise RuntimeError("미리보기할 페이지가 없습니다.")

        page = doc.load_page(0)
        zoom = max_width / page.rect.width
        matrix = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=matrix)

        out_dir = os.path.join(tempfile.gettempdir(), PREVIEW_DIR_NAME)
        os.makedirs(out_dir, exist_ok=True)

        timestamp = int(time.time() * 1000)
        preview_path = os.path.join(
            out_dir, f"{safe_stem(pdf_path)}_{file_key(pdf_path)}_{timestamp}.png"
        )

        pix.save(preview_path)
        return preview_path
    finally:
        doc.close()


def get_preview_image_path(file_path):
    """
    출력 전 미리볼 이미지 파일 경로를 반환한다.
    이미지 파일은 그대로, PDF/오피스 문서는 첫 페이지를 이미지로 변환해 반환한다.
    HWP/HWPX 는 LibreOffice 한글 확장 프로그램을 통해 변환한다.
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext in IMAGE_EXTENSIONS:
        return file_path

    if ext == ".pdf":
        return render_pdf_first_page_image(file_path)

    if ext in LIBREOFFICE_EXTENSIONS:
        pdf_path = convert_libreoffice_to_pdf(file_path)
        return render_pdf_first_page_image(pdf_path)

    raise RuntimeError("이 파일 형식은 미리보기를 지원하지 않습니다.")


# -----------------------------
# CSV / XLSX 이용자 불러오기, 내보내기
# -----------------------------
def parse_bool(value, default=0):
    if value is None:
        return default

    text = str(value).strip().lower()

    if text in ["1", "true", "yes", "y", "o", "on", "예", "네", "사용", "무제한", "가능"]:
        return 1

    if text in ["0", "false", "no", "n", "x", "off", "아니오", "아니요", "미사용", "중지", "비활성", "불가"]:
        return 0

    return default


def normalize_header(header):
    if header is None:
        return ""

    text = str(header).strip().lower()
    text = text.replace(" ", "").replace("_", "").replace("-", "")
    return text


def find_column_index(headers, candidates):
    normalized_headers = [normalize_header(h) for h in headers]
    normalized_candidates = [normalize_header(c) for c in candidates]

    for candidate in normalized_candidates:
        if candidate in normalized_headers:
            return normalized_headers.index(candidate)

    return None


def read_csv_rows(path):
    encodings = ["utf-8-sig", "cp949", "euc-kr", "utf-8"]

    last_error = None

    for enc in encodings:
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                reader = csv.reader(f)
                return list(reader)
        except Exception as e:
            last_error = e

    raise RuntimeError(f"CSV 파일을 읽지 못했습니다.\n{last_error}")


def read_xlsx_rows(path):
    if openpyxl is None:
        raise RuntimeError("XLSX 불러오기를 위해 openpyxl이 필요합니다.")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(list(row))

    return rows


def import_users_from_file(path):
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        rows = read_csv_rows(path)
    elif ext == ".xlsx":
        rows = read_xlsx_rows(path)
    else:
        raise RuntimeError("CSV 또는 XLSX 파일만 불러올 수 있습니다.")

    rows = [row for row in rows if any(cell is not None and str(cell).strip() for cell in row)]

    if not rows:
        raise RuntimeError("파일에 데이터가 없습니다.")

    headers = rows[0]

    code_idx = find_column_index(headers, [
        "card_code", "cardcode", "student_id", "studentid",
        "학생증코드", "학생증 코드", "카드코드", "코드"
    ])

    number_idx = find_column_index(headers, [
        "student_number", "studentnumber", "student_no", "studentno",
        "학번"
    ])

    name_idx = find_column_index(headers, [
        "name", "이름", "성명", "사용자명"
    ])

    unlimited_idx = find_column_index(headers, [
        "unlimited", "unlimited_print", "무제한", "무제한출력", "무제한 출력"
    ])

    active_idx = find_column_index(headers, [
        "active", "enabled", "사용", "상태", "활성화", "사용여부"
    ])

    if code_idx is None or name_idx is None:
        raise RuntimeError(
            "필수 열이 없습니다.\n"
            "CSV/XLSX 첫 줄에 '학생증코드'와 '이름' 열이 있어야 합니다."
        )

    imported = 0
    skipped = 0

    for row in rows[1:]:
        try:
            code = row[code_idx] if code_idx < len(row) else ""
            name = row[name_idx] if name_idx < len(row) else ""
            number = ""

            if number_idx is not None and number_idx < len(row):
                number = row[number_idx]

            unlimited = 0
            active = 1

            if unlimited_idx is not None and unlimited_idx < len(row):
                unlimited = parse_bool(row[unlimited_idx], default=0)

            if active_idx is not None and active_idx < len(row):
                active = parse_bool(row[active_idx], default=1)

            code = normalize_card_code(code)
            number = normalize_card_code(number)
            name = str(name).strip()

            if not code or not name:
                skipped += 1
                continue

            upsert_student(code, name, student_number=number, active=active, unlimited=unlimited)
            imported += 1

        except Exception:
            skipped += 1

    return imported, skipped


def _normalize_date_input(value, default=None):
    """
    '2026-03-01', '2026/3/1', '20260301' 형태를 모두 'YYYY-MM-DD'로 바꾼다.
    비어 있으면 default를 돌려준다.
    """
    value = str(value or "").strip()

    if not value:
        return default

    cleaned = value.replace("/", "-").replace(".", "-")

    if cleaned.isdigit() and len(cleaned) == 8:
        cleaned = f"{cleaned[:4]}-{cleaned[4:6]}-{cleaned[6:]}"

    parts = cleaned.split("-")
    if len(parts) != 3:
        raise ValueError(f"날짜 형식이 올바르지 않습니다: {value}")

    year, month, day = parts
    if not (year.isdigit() and month.isdigit() and day.isdigit()):
        raise ValueError(f"날짜 형식이 올바르지 않습니다: {value}")

    return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"


def _build_log_filters(start_date=None, end_date=None, card_code=None):
    where = []
    params = []

    if start_date:
        where.append("p.print_date >= ?")
        params.append(start_date)

    if end_date:
        where.append("p.print_date <= ?")
        params.append(end_date)

    if card_code:
        where.append("p.student_id = ?")
        params.append(card_code)

    clause = (" WHERE " + " AND ".join(where)) if where else ""
    return clause, params


def export_logs_detail_to_csv(path, start_date=None, end_date=None, card_code=None):
    """
    출력 기록을 건별로 CSV에 저장한다.
    기간과 특정 사용자 조건은 값이 있을 때만 적용된다.
    반환값: 저장한 행 수
    """
    clause, params = _build_log_filters(start_date, end_date, card_code)

    conn = connect_db()
    cur = conn.cursor()
    cur.execute(f"""
    SELECT p.print_date, p.print_time, p.student_id,
           COALESCE(s.student_number, ''), p.student_name,
           p.file_name, p.sheet_count, p.duplex, p.printer_name, p.result
    FROM print_logs p
    LEFT JOIN students s ON s.student_id = p.student_id
    {clause}
    ORDER BY p.print_date DESC, p.print_time DESC
    """, params)
    rows = cur.fetchall()
    conn.close()

    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            "날짜", "시간", "학생증코드", "학번", "이름",
            "파일명", "용지 장수", "인쇄 방식", "프린터", "결과"
        ])
        for row in rows:
            row = list(row)
            row[7] = "양면" if row[7] else "단면"
            writer.writerow(row)

    return len(rows)


def export_logs_summary_to_csv(path, start_date=None, end_date=None):
    """
    사용자별 출력 횟수 집계를 CSV에 저장한다.
    실패한 출력은 횟수에서 제외한다.
    반환값: 저장한 행 수
    """
    clause, params = _build_log_filters(start_date, end_date)

    # 집계 대상은 print_logs 이므로 s.가 아닌 p. 기준으로 필터가 걸린다
    conn = connect_db()
    cur = conn.cursor()
    cur.execute(f"""
    SELECT p.student_id,
           COALESCE(s.student_number, ''),
           MAX(p.student_name),
           SUM(CASE WHEN p.result != '실패' THEN 1 ELSE 0 END),
           SUM(CASE WHEN p.result != '실패' THEN p.sheet_count ELSE 0 END),
           SUM(CASE WHEN p.result = '실패' THEN 1 ELSE 0 END),
           COUNT(*),
           MIN(p.print_date),
           MAX(p.print_date)
    FROM print_logs p
    LEFT JOIN students s ON s.student_id = p.student_id
    {clause}
    GROUP BY p.student_id
    ORDER BY 5 DESC
    """, params)
    rows = cur.fetchall()
    conn.close()

    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            "학생증코드", "학번", "이름",
            "성공 출력 수", "사용한 용지 장수", "실패 수", "전체 시도 수",
            "첫 출력일", "마지막 출력일"
        ])
        writer.writerows(rows)

    return len(rows)


def export_users_to_csv(path):
    conn = connect_db()
    cur = conn.cursor()
    cur.execute("""
    SELECT student_id, student_number, name, unlimited, active, created_date
    FROM students
    ORDER BY student_id
    """)
    rows = cur.fetchall()
    conn.close()

    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["학생증코드", "학번", "이름", "무제한출력", "사용상태", "등록일"])

        for card_code, student_number, name, unlimited, active, created_date in rows:
            writer.writerow([
                card_code,
                student_number or "",
                name,
                "예" if unlimited else "아니요",
                "사용 가능" if active else "사용 중지",
                created_date
            ])


# -----------------------------
# 이용자 명단 CSV 실시간 연동
#
# 프로그램 폴더에 'user_list.csv' 파일을 하나 두고 명단과 계속 맞춰 놓는다.
#   · 프로그램에서 명단이 바뀌면   → CSV 를 곧바로 새로 쓴다
#   · CSV 를 엑셀 등에서 고쳐 저장하면 → 몇 초 안에 프로그램 명단에 반영된다
#
# CSV 에서 줄을 지워도 이용자가 삭제되지는 않는다.
# 파일을 잘못 저장했을 때 명단이 통째로 날아가는 사고를 막기 위해서이며,
# 삭제는 이용자 관리 화면의 '삭제' 버튼으로만 할 수 있다.
# -----------------------------
# 파일 이름은 영문으로 둔다.
# 한글 이름은 환경(인코딩·백업 도구·네트워크 드라이브)에 따라 문제가 생길 여지가 있다.
SYNC_CSV_NAME = "user_list.csv"

# 우리가 방금 쓴 파일을 '바깥에서 바뀐 것' 으로 착각하지 않도록
# 마지막으로 확인한 파일 상태를 기억해 둔다. (수정시각, 크기)
_sync_csv_state = {"stamp": None}


def get_program_dir():
    """프로그램(또는 exe) 이 놓인 폴더."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)

    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()


# 예전 판에서 쓰던 이름들. 남아 있으면 내용을 합친 뒤 지운다.
LEGACY_SYNC_CSV_NAMES = ["이용자명단.csv"]


def get_sync_csv_name():
    return SYNC_CSV_NAME


def get_sync_csv_path():
    return os.path.join(get_program_dir(), SYNC_CSV_NAME)


def _csv_stamp(path):
    try:
        info = os.stat(path)
        return (int(info.st_mtime), info.st_size)
    except OSError:
        return None


def write_sync_csv():
    """
    지금 명단을 연동용 CSV 로 내보낸다.
    엑셀에서 파일을 열어둔 채라 저장이 막힐 수 있으므로, 실패해도 그냥 넘어간다.
    """
    path = get_sync_csv_path()

    try:
        export_users_to_csv(path)
        ok = True
    except Exception:
        # 엑셀에서 파일을 열어두면 저장이 막히는데, 이건 잠시 뒤 다시 하면 되는
        # 일이라 파일 이름을 바꾸지 않는다. 이름을 바꿔버리면 명단 파일이
        # 두 개로 갈라진다.
        ok = False

    _sync_csv_state["stamp"] = _csv_stamp(path)
    return ok


def merge_legacy_sync_csv():
    """
    예전 판에서 쓰던 이름의 명단 파일이 남아 있으면 하나로 합친다.

    그 파일의 내용을 명단에 한 번 반영(추가·수정)한 뒤 지우고,
    지금 쓰는 이름의 파일만 남긴다. 파일에만 있던 학생도 사라지지 않는다.
    """
    merged = False

    for name in LEGACY_SYNC_CSV_NAMES:
        legacy_path = os.path.join(get_program_dir(), name)

        if not os.path.exists(legacy_path):
            continue

        try:
            import_users_from_file(legacy_path)
        except Exception:
            # 형식이 깨진 파일이면 그냥 지운다 (명단은 DB 에 있다)
            pass

        try:
            os.remove(legacy_path)
            merged = True
        except OSError:
            pass

    if merged:
        write_sync_csv()

    return merged


def sync_csv_changed_outside():
    """CSV 가 프로그램 밖에서 바뀌었는지 확인한다."""
    stamp = _csv_stamp(get_sync_csv_path())

    if stamp is None:
        return False

    return stamp != _sync_csv_state["stamp"]


def load_sync_csv():
    """
    CSV 내용을 명단에 반영한다. (추가와 수정만 하고 삭제는 하지 않는다)
    반영이 끝나면 다듬어진 값으로 CSV 를 다시 써서 양쪽을 같은 모양으로 맞춘다.
    """
    path = get_sync_csv_path()

    try:
        imported, skipped = import_users_from_file(path)
    except Exception:
        # 저장 중이거나 형식이 잘못된 파일. 지금 상태를 기준으로 삼고 다음 변경을 기다린다.
        _sync_csv_state["stamp"] = _csv_stamp(path)
        raise

    write_sync_csv()
    return imported, skipped


def delete_student(card_code):
    """이용자를 명단에서 완전히 지운다. 출력 기록은 그대로 남긴다."""
    card_code = normalize_card_code(card_code)

    conn = connect_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM students WHERE student_id=?", (card_code,))
    conn.commit()
    deleted = cur.rowcount
    conn.close()

    return deleted


# -----------------------------
# 메인 앱
# -----------------------------
# -----------------------------
# 디자인 시스템 (색상 / 폰트)
# -----------------------------
COLOR_BG = "#F1F2F5"            # 화면 배경
COLOR_CARD = "#FFFFFF"          # 카드(패널) 배경
COLOR_BORDER = "#E4E6EB"        # 카드 테두리 / 구분선
COLOR_TEXT = "#1C1F26"          # 기본 텍스트
COLOR_TEXT_MUTED = "#868C98"    # 보조 설명 텍스트
COLOR_PRIMARY = "#2F6FED"       # 포인트 컬러(파랑)
COLOR_PRIMARY_HOVER = "#2557C7"
COLOR_PRIMARY_SOFT = "#EAF1FE"  # 포인트 컬러의 옅은 배경(뱃지 등)
COLOR_SECONDARY_BG = "#EEF0F3"  # 보조 버튼 배경
COLOR_SECONDARY_BG_HOVER = "#E1E4E9"
COLOR_SECONDARY_TEXT = "#3A3F47"
# 고장 신고 안내 배너 (연한 붉은색)
COLOR_REPORT_BANNER_BG   = "#FDECEC"
COLOR_REPORT_BANNER_LINE = "#F3BEBE"
COLOR_REPORT_BANNER_TEXT = "#B4353A"

COLOR_DANGER = "#E5484D"
COLOR_DANGER_HOVER = "#C8383D"
COLOR_SUCCESS = "#1F9254"
COLOR_ENTRY_BG = "#F7F8FA"

FALLBACK_FONT_FAMILY = "맑은 고딕"

# -----------------------------
# 프로그램 아이콘
# icon.ico 파일을 exe와 같은 폴더(또는 PyInstaller로 포함시킨 경우 내부)에 두면
# 창 왼쪽 위와 작업표시줄 아이콘이 그것으로 바뀐다.
# -----------------------------
ICON_FILE_NAME = "icon.ico"


def get_resource_path(file_name):
    """
    PyInstaller로 만든 exe 안에 포함된 파일과
    일반 파이썬 실행 시의 파일을 모두 찾아준다.
    """
    # PyInstaller가 --add-data로 포함시킨 파일은 임시 폴더(_MEIPASS)에 풀린다
    base = getattr(sys, "_MEIPASS", None)

    if base:
        bundled = os.path.join(base, file_name)
        if os.path.exists(bundled):
            return bundled

    # exe(또는 스크립트)와 같은 폴더
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))

    beside = os.path.join(exe_dir, file_name)
    if os.path.exists(beside):
        return beside

    return None


def apply_window_icon(window):
    """창 왼쪽 위와 작업표시줄 아이콘을 icon.ico로 바꾼다."""
    icon_path = get_resource_path(ICON_FILE_NAME)

    if not icon_path:
        return False

    try:
        window.iconbitmap(icon_path)
        # 일부 환경에서 작업표시줄 아이콘이 늦게 적용되는 것을 보완
        window.after(200, lambda: window.iconbitmap(icon_path))
        return True
    except Exception:
        return False

# -----------------------------
# 우측 하단 버튼이 여는 외부 사이트 주소
# 고장 신고 주소는 학교에서 만든 구글 폼 링크로 바꿔서 사용하세요.
# -----------------------------
REPORT_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSfU5cdC3w6QjQnQWkM9tVHLn3PX4BZrPHMMl2efaTDtAhM3Yw/viewform?usp=header"   # 예: "https://forms.gle/XXXXXXXXXXXX"

# ── 고장 신고 상태 확인 ────────────────────────────────
#
# 신고가 들어오거나 처리되면 Apps Script 가 GitHub 저장소의 작은 파일
# (printer_status.json) 을 고쳐 쓴다. 프로그램은 그 파일만 짧은 주기로
# 확인하면 되므로, Apps Script 를 계속 부르지 않아도 된다.
#
#   REPORT_STATUS_FILE : 평소 확인용. GitHub raw 주소 (빠르고 부담이 없다)
#   REPORT_STATUS_URL  : 직접 확인용. Apps Script /exec 주소
#                        (새로고침 버튼처럼 지금 당장 정확한 값이 필요할 때)
#
# 둘 다 비워두면 이 기능을 쓰지 않는다. 하나만 넣어도 동작한다.
REPORT_STATUS_FILE = "https://raw.githubusercontent.com/sicgaonnury/printer/main/printer_status.json"  # 예: "https://raw.githubusercontent.com/사용자/저장소/main/printer_status.json"
REPORT_STATUS_URL = "https://script.google.com/macros/s/AKfycbxODWNIgGZvHPiQeLBYrXIz2mqiRXYKdxHJ9ifpKuDUApiik4kDGoD7wHEtiCjFw2Mr/exec"   # 예: "https://script.google.com/macros/s/AKfy.../exec"

# 몇 초마다 신고 상태를 확인할지.
# GitHub 파일을 보는 것은 가벼워서 짧게 잡아도 된다.
REPORT_CHECK_SECONDS = 15

# 파일 공유용 내장 서버가 사용할 포트
FILE_SHARE_PORT = 8888


def report_status_enabled():
    """고장 신고 상태 확인 기능을 쓸 수 있는지."""
    return bool(REPORT_STATUS_FILE or REPORT_STATUS_URL)


def _read_status_json(url, timeout=10):
    """주소에서 상태 JSON 을 읽어 온다. 실패하면 None."""
    import json
    import ssl
    import urllib.request

    try:
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE

        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": "SchoolPrinterKiosk",
                # 중간 서버가 예전 내용을 돌려주지 않도록
                "Cache-Control": "no-cache",
                "Pragma": "no-cache",
            },
        )

        with urllib.request.urlopen(request, context=context, timeout=timeout) as response:
            data = json.loads(response.read().decode("utf-8"))

        if not isinstance(data, dict):
            return None

        return {
            "open": int(data.get("open", 0) or 0),
            "status": str(data.get("status", "") or ""),
            "id": data.get("id", 0),
            "reason": str(data.get("reason", "") or ""),
            "time": str(data.get("time", "") or ""),
        }

    except Exception:
        return None


def fetch_report_status(direct=False):
    """
    처리되지 않은 고장 신고가 있는지 확인한다.

    direct=False (평소)  : GitHub 파일을 본다. 파일 주소가 없으면 Apps Script 로 간다.
    direct=True  (수동)  : Apps Script 에 직접 물어본다. 값이 가장 정확하다.

    돌려주는 값은 {"open": 개수, "status": 상태, "id": 신고번호, "reason": 원인, "time": 시각}.
    확인에 실패하면 None 을 돌려주고, 이 경우 화면은 지금 상태를 그대로 둔다.
    (인터넷이 잠깐 끊겼다고 안내가 사라지면 안 된다)
    """
    order = []

    if direct:
        order = [REPORT_STATUS_URL, REPORT_STATUS_FILE]
    else:
        order = [REPORT_STATUS_FILE, REPORT_STATUS_URL]

    for url in order:
        if not url:
            continue

        target = url

        # GitHub raw 는 같은 주소를 5분쯤 담아두기 때문에,
        # 뒤에 시각을 붙여 매번 새 내용을 받도록 한다.
        if "raw.githubusercontent.com" in url:
            joiner = "&" if "?" in url else "?"
            target = f"{url}{joiner}t={int(time.time())}"

        data = _read_status_json(target)

        if data is not None:
            return data

    return None


# -----------------------------
# 파일 공유 — 내장 업로드 서버
#
# 외부 파일 전송 서비스는 기기끼리 직접 통신해야 하는데, 학교 망은 보통
# 기기 간 통신을 막아두어(AP 격리) 동작하지 않는다.
# 그래서 제어 컴퓨터가 직접 작은 웹서버를 열고, 학생이 브라우저로 접속해
# 파일을 올리는 방식을 쓴다. 학교 망 안에서만 오가므로 외부 서비스가 필요 없다.
# -----------------------------
_upload_server = None
_upload_server_thread = None


def get_local_ip():
    """학교 망에서 다른 기기가 접속할 수 있는 이 컴퓨터의 주소를 찾는다."""
    import socket

    try:
        # 실제로 연결하지는 않고, 어떤 랜카드를 쓸지만 알아낸다
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
        finally:
            sock.close()
    except Exception:
        pass

    try:
        return socket.gethostbyname(socket.gethostname())
    except Exception:
        return "127.0.0.1"


def get_file_share_url(token=None):
    """
    학생 기기가 접속할 주소.
    토큰을 붙이면 그 학생 전용 주소가 되어, 올린 파일에 학번·이름이 붙는다.
    """
    base = f"http://{get_local_ip()}:{FILE_SHARE_PORT}"
    return f"{base}/?t={token}" if token else base


UPLOAD_DENIED_HTML = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>학생증 인증이 필요합니다</title>
<style>
  body {
    margin: 0; padding: 60px 20px;
    font-family: -apple-system, BlinkMacSystemFont, "Malgun Gothic", sans-serif;
    background: #F1F2F5; color: #1C1F26; text-align: center;
  }
  .card {
    max-width: 420px; margin: 0 auto; background: #fff;
    border-radius: 18px; padding: 40px 26px;
    box-shadow: 0 2px 14px rgba(0,0,0,0.06);
  }
  h1 { font-size: 20px; margin: 0 0 14px; }
  p { color: #4A5058; font-size: 15px; line-height: 1.7; margin: 0; }
  .step {
    background: #F4F5F7; border-radius: 12px;
    padding: 16px 18px; margin-top: 22px;
    font-size: 14px; color: #4A5058; text-align: left; line-height: 1.9;
  }
</style>
</head>
<body>
<div class="card">
  <h1>학생증 인증이 필요합니다</h1>
  <p>이 주소로는 파일을 보낼 수 없습니다.</p>
  <div class="step">
    1. 프린터 앞 화면에서 <b>파일 공유</b>를 누르세요.<br>
    2. <b>학생증 바코드</b>를 스캔하세요.<br>
    3. 화면에 나오는 <b>QR 코드</b>를 찍어 다시 접속하세요.
  </div>
</div>
</body>
</html>
"""


UPLOAD_PAGE_HTML = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>파일 보내기</title>
<style>
  * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
  body {
    margin: 0; padding: 24px 18px;
    font-family: -apple-system, BlinkMacSystemFont, "Malgun Gothic", sans-serif;
    background: #F1F2F5; color: #1C1F26;
  }
  .card {
    max-width: 460px; margin: 0 auto; background: #fff;
    border-radius: 18px; padding: 28px 22px;
    box-shadow: 0 2px 14px rgba(0,0,0,0.06);
  }
  h1 { font-size: 21px; margin: 0 0 6px; }
  .sub { color: #868C98; font-size: 14px; margin-bottom: 12px; }
  .formats {
    font-size: 11px; color: #A5AAB3; text-align: center;
    margin-bottom: 14px; line-height: 1.6;
  }
  .who {
    background: #EAF1FE; color: #2557C7; font-weight: 700;
    font-size: 15px; border-radius: 10px; padding: 10px 14px;
    text-align: center; margin-bottom: 22px;
  }
  .drop {
    border: 2px dashed #C7CBD1; border-radius: 14px;
    padding: 38px 16px; text-align: center; color: #868C98;
    font-size: 15px; margin-bottom: 18px;
  }
  input[type=file] { display: none; }
  .btn {
    display: block; width: 100%; padding: 16px;
    background: #2F6FED; color: #fff; border: none;
    border-radius: 12px; font-size: 17px; font-weight: 700;
    cursor: pointer; margin-bottom: 10px;
  }
  .btn.sub-btn { background: #EDEFF2; color: #1C1F26; }
  .btn:disabled { background: #C7CBD1; }
  #list { margin: 14px 0; font-size: 14px; }
  #list div { padding: 7px 0; border-bottom: 1px solid #EDEFF2; }
  #status { margin-top: 14px; font-size: 15px; text-align: center; min-height: 24px; }
  .ok { color: #1F9254; font-weight: 700; }
  #done {
    display: none; margin-top: 18px; text-align: center;
    background: #EAF7EF; border-radius: 14px; padding: 22px 16px;
  }
  #done .check {
    width: 52px; height: 52px; line-height: 52px; margin: 0 auto 12px;
    background: #1F9254; color: #fff; border-radius: 50%; font-size: 26px;
  }
  #doneText { font-size: 18px; font-weight: 700; color: #1F7A46; }
  #done .donesub { font-size: 14px; color: #4A5058; margin-top: 6px; }
  .err { color: #C8383D; font-weight: 700; }
  .note { margin-top: 20px; font-size: 13px; color: #868C98; line-height: 1.6; }
</style>
</head>
<body>
<div class="card">
  <h1>파일 보내기</h1>
  <div class="sub">공용 프린터로 파일을 보냅니다</div>
  <div class="formats">PDF · DOC · DOCX · HWP · HWPX · PPT · PPTX · XLS · XLSX · CSV · TXT · JPG · PNG</div>
  <div class="who">{{USER}}</div>

  <label for="file">
    <div class="drop" id="drop">여기를 눌러 파일을 고르세요</div>
  </label>
  <input type="file" id="file" multiple accept=".pdf,.doc,.docx,.hwp,.hwpx,.jpg,.jpeg,.png,.ppt,.pptx,.csv,.xls,.xlsx,.txt">

  <div id="list"></div>

  <button class="btn" id="send" disabled>보내기</button>
  <div id="status"></div>

  <div id="done">
    <div class="check">&#10004;</div>
    <div id="doneText"></div>
    <div class="donesub">프린터 옆 컴퓨터에서 출력하세요</div>
  </div>

  <div class="note">
    파일을 보낸 뒤 프린터 옆 컴퓨터에서 출력하면 됩니다.<br>
    보낸 파일은 일정 시간이 지나면 자동으로 지워집니다.
  </div>
</div>

<script>
var input = document.getElementById('file');
var list = document.getElementById('list');
var send = document.getElementById('send');
var status = document.getElementById('status');
var done = document.getElementById('done');
var doneText = document.getElementById('doneText');
var drop = document.getElementById('drop');

input.addEventListener('change', function () {
  list.innerHTML = '';
  for (var i = 0; i < input.files.length; i++) {
    var d = document.createElement('div');
    d.textContent = input.files[i].name;
    list.appendChild(d);
  }
  send.disabled = input.files.length === 0;
  done.style.display = 'none';
  drop.textContent = input.files.length
    ? input.files.length + '개 파일 선택됨'
    : '여기를 눌러 파일을 고르세요';
});

send.addEventListener('click', function () {
  if (!input.files.length) return;

  var form = new FormData();
  for (var i = 0; i < input.files.length; i++) {
    form.append('file', input.files[i]);
  }

  send.disabled = true;
  status.className = '';
  status.textContent = '보내는 중입니다...';

  var xhr = new XMLHttpRequest();
  // 주소에 붙은 학생 토큰을 그대로 넘겨야 서버가 누구인지 알 수 있다
  xhr.open('POST', '/upload' + window.location.search);

  xhr.upload.onprogress = function (e) {
    if (e.lengthComputable) {
      status.textContent = '보내는 중... ' + Math.round(e.loaded / e.total * 100) + '%';
    }
  };

  xhr.onload = function () {
    if (xhr.status === 200) {
      var parts = xhr.responseText.split('|');
      var n = parseInt(parts[0], 10) || 0;

      doneText.textContent = '파일 ' + n + '개를 보냈습니다';
      done.style.display = 'block';
      status.className = '';
      status.textContent = '';

      // 형식이 맞지 않아 빠진 파일이 있으면 알려준다
      if (parts.length > 1 && parts[1]) {
        status.className = 'err';
        status.textContent = '출력할 수 없는 형식이라 제외됨: ' + parts[1];
      }

      drop.textContent = '여기를 눌러 파일을 고르세요';
      list.innerHTML = '';
      input.value = '';
      send.disabled = true;

      // 진동으로도 알려준다 (지원하는 기기에서만)
      if (navigator.vibrate) { navigator.vibrate(120); }
    } else if (xhr.status === 415) {
      status.className = 'err';
      status.textContent = '출력할 수 없는 형식입니다: ' + xhr.responseText.replace('UNSUPPORTED:', '');
      send.disabled = false;
    } else if (xhr.status === 413) {
      status.className = 'err';
      status.textContent = '파일이 너무 큽니다. 한 번에 '
        + xhr.responseText.replace('TOOBIG:', '') + 'MB 까지 보낼 수 있습니다.';
      send.disabled = false;
    } else if (xhr.status === 403) {
      status.className = 'err';
      status.textContent = '인증이 만료되었습니다. 학생증을 다시 스캔하세요.';
      send.disabled = false;
    } else {
      status.className = 'err';
      status.textContent = '전송에 실패했습니다. 다시 시도하세요.';
      send.disabled = false;
    }
  };

  xhr.onerror = function () {
    status.className = 'err';
    status.textContent = '전송에 실패했습니다. 와이파이 연결을 확인하세요.';
    send.disabled = false;
  };

  xhr.send(form);
});
</script>
</body>
</html>
"""


def _safe_file_name(raw_name):
    """업로드된 파일명에서 위험한 문자를 없앤다."""
    name = os.path.basename(raw_name.replace("\\", "/"))
    name = "".join(c for c in name if c not in '\\/:*?"<>|')
    name = name.strip().strip(".")
    return name


def _unique_path(folder, name):
    """같은 이름이 있으면 뒤에 번호를 붙여 겹치지 않는 경로를 만든다."""
    path = os.path.join(folder, name)
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(path):
        path = f"{base} ({i}){ext}"
        i += 1
    return path


# -----------------------------
# 파일 공유 세션
#
# 아무나 파일을 올리지 못하도록, 학생증을 스캔한 사람만 접속 주소를 받는다.
# 주소에 붙는 토큰으로 누가 올린 파일인지 구분해서
# "학번 이름1.pdf" 처럼 이름을 붙여 저장한다.
# -----------------------------
_share_sessions = {}          # {토큰: {"number":학번, "name":이름, "seq":보낸 개수}}
_upload_notify = None         # 파일이 도착하면 화면에 알리기 위한 함수
_share_session_lock = threading.Lock()


def set_upload_notifier(func):
    """파일이 도착했을 때 화면에 알릴 함수를 등록한다."""
    global _upload_notify
    _upload_notify = func


def notify_upload(token, saved_paths):
    """업로드가 끝났음을 화면에 알린다. 화면 쪽 오류가 서버를 멈추지 않게 감싼다."""
    if _upload_notify is None:
        return

    session = get_share_session(token)
    label = ""

    if session:
        label = f"{session['number']} {session['name']}".strip()

    try:
        _upload_notify(label, list(saved_paths))
    except Exception:
        pass


def create_share_session(student_number, name):
    """
    학생 한 명을 위한 업로드 세션을 만들고 토큰을 돌려준다.

    노트북 사용자가 주소를 손으로 입력해야 할 수도 있으므로
    헷갈리기 쉬운 글자(0/O, 1/l/I)를 뺀 6자리 코드를 쓴다.
    """
    import secrets

    alphabet = "ABCDEFGHJKMNPQRSTUVWXYZ23456789"

    with _share_session_lock:
        while True:
            token = "".join(secrets.choice(alphabet) for _ in range(6))
            if token not in _share_sessions:
                break

        _share_sessions[token] = {
            "number": student_number or "",
            "name": name or "",
            "seq": 0,
        }

    return token


def get_share_session(token):
    with _share_session_lock:
        return _share_sessions.get(token)


def end_share_session(token):
    with _share_session_lock:
        _share_sessions.pop(token, None)


def next_share_index(token):
    """이 학생이 몇 번째로 보내는 파일인지 번호를 매긴다."""
    with _share_session_lock:
        session = _share_sessions.get(token)
        if session is None:
            return None
        session["seq"] += 1
        return session["seq"]


def build_share_file_name(token, original_name):
    """
    저장할 파일 이름을 만든다.
    형식: "학번 이름1.확장자"  (같은 학생이 여러 개 보내면 번호가 올라간다)
    """
    session = get_share_session(token)

    if session is None:
        return original_name

    index = next_share_index(token) or 1
    ext = os.path.splitext(original_name)[1]

    number = session["number"].strip()
    name = session["name"].strip()

    if number and name:
        label = f"{number} {name}"
    else:
        label = number or name or "이용자"

    return f"{label}{index}{ext}"


# 한 번에 보낼 수 있는 최대 용량. 학생 과제 파일에는 넉넉하고,
# 디스크를 채우려는 장난은 막을 수 있는 선으로 잡았다.
MAX_UPLOAD_TOTAL = 1024 * 1024 * 1024   # 한 번에 보내는 전체 용량 1GB
MAX_UPLOAD_FILE = 1024 * 1024 * 1024     # 파일 하나당 1GB


def save_multipart_upload(stream, content_type, content_length, save_dir, token=None):
    """
    브라우저가 보낸 multipart/form-data 를 읽어 파일로 저장한다.

    파이썬 3.13 에서 cgi 모듈이 표준 라이브러리에서 빠졌기 때문에
    (파이썬 3.14 를 쓰면 import 자체가 실패한다) 직접 파싱한다.

    큰 파일도 메모리를 적게 쓰도록 조금씩 읽어 바로 파일에 쓴다.
    (저장한 파일 경로 목록, 형식이 맞지 않아 거른 파일 이름 목록) 을 돌려준다.
    """
    if "multipart/form-data" not in content_type.lower():
        return [], []

    # 경계 문자열 찾기 (boundary=----XXXX)
    boundary = None
    for part in content_type.split(";"):
        part = part.strip()
        if part.lower().startswith("boundary="):
            boundary = part[9:].strip().strip('"')
            break

    if not boundary:
        return [], []

    delimiter = b"--" + boundary.encode("utf-8")
    os.makedirs(save_dir, exist_ok=True)

    CHUNK = 64 * 1024
    buffer = b""
    remaining = content_length
    saved_paths = []
    rejected = []          # 지원하지 않는 형식이거나 너무 커서 거른 파일 이름

    def fill():
        """다음 덩어리를 읽어 버퍼에 채운다. 더 읽을 게 없으면 False."""
        nonlocal buffer, remaining
        if remaining <= 0:
            return False
        chunk = stream.read(min(CHUNK, remaining))
        if not chunk:
            remaining = 0
            return False
        remaining -= len(chunk)
        buffer += chunk
        return True

    # 첫 경계까지 건너뛴다
    while delimiter not in buffer:
        if not fill():
            return [], []
    buffer = buffer.split(delimiter, 1)[1]

    while True:
        # 경계 뒤 줄바꿈 또는 마지막 표시(--)
        while len(buffer) < 2:
            if not fill():
                return saved_paths, rejected

        if buffer.startswith(b"--"):
            break   # 마지막 경계

        if buffer.startswith(b"\r\n"):
            buffer = buffer[2:]

        # 헤더 읽기
        while b"\r\n\r\n" not in buffer:
            if not fill():
                return saved_paths, rejected

        header_blob, buffer = buffer.split(b"\r\n\r\n", 1)
        headers = header_blob.decode("utf-8", errors="replace")

        # 파일명 추출
        file_name = ""
        for line in headers.split("\r\n"):
            if "content-disposition" not in line.lower():
                continue
            for piece in line.split(";"):
                piece = piece.strip()
                if piece.lower().startswith("filename="):
                    file_name = piece[9:].strip().strip('"')
            break

        file_name = _safe_file_name(file_name) if file_name else ""

        # 출력할 수 없는 형식은 아예 저장하지 않는다.
        # (저장해두면 출력 목록에 떠서 학생이 고를 수 있게 된다)
        if file_name and not is_supported_file(file_name):
            rejected.append(file_name)
            file_name = ""

        # 학생증을 스캔한 세션이면 "학번 이름1.pdf" 형식으로 이름을 바꾼다
        if file_name and token:
            file_name = _safe_file_name(build_share_file_name(token, file_name))

        # 본문을 경계가 나올 때까지 읽어 저장한다
        target_path = _unique_path(save_dir, file_name) if file_name else None
        target = open(target_path, "wb") if target_path else None
        finished = False
        written = 0
        too_big = False

        def write_chunk(data):
            """파일 하나가 정해진 용량을 넘으면 더 쓰지 않는다."""
            nonlocal written, too_big

            if target is None or too_big:
                return

            if written + len(data) > MAX_UPLOAD_FILE:
                too_big = True
                return

            target.write(data)
            written += len(data)

        try:
            while True:
                index = buffer.find(delimiter)

                if index >= 0:
                    body = buffer[:index]
                    if body.endswith(b"\r\n"):
                        body = body[:-2]
                    write_chunk(body)
                    buffer = buffer[index + len(delimiter):]
                    finished = True
                    break

                # 경계가 덩어리 사이에 걸칠 수 있으므로 끝부분은 남겨둔다
                keep = len(delimiter) + 2
                if len(buffer) > keep:
                    write_chunk(buffer[:-keep])
                    buffer = buffer[-keep:]

                if not fill():
                    write_chunk(buffer)
                    buffer = b""
                    break
        finally:
            if target:
                target.close()

                if too_big:
                    # 용량을 넘긴 파일은 남기지 않는다
                    try:
                        os.remove(target_path)
                    except OSError:
                        pass
                    rejected.append(f"{file_name} (용량 초과)")
                else:
                    saved_paths.append(target_path)

        if not finished:
            break

    return saved_paths, rejected


def start_upload_server():
    """
    파일 업로드용 웹서버를 띄운다. 이미 떠 있으면 그대로 둔다.
    실패하면 오류 메시지를 돌려준다 (성공 시 None).
    """
    global _upload_server, _upload_server_thread

    if _upload_server is not None:
        return None

    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass   # 콘솔에 접속 기록을 남기지 않는다

        def _token(self):
            """주소에 붙은 학생 토큰을 꺼낸다. (예: /?t=abc123)"""
            from urllib.parse import urlparse, parse_qs

            query = parse_qs(urlparse(self.path).query)
            values = query.get("t") or []
            return values[0] if values else ""

        def _deny(self):
            """학생증 인증이 없으면 안내 페이지를 보여준다."""
            body = UPLOAD_DENIED_HTML.encode("utf-8")
            self.send_response(403)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            token = self._token()
            session = get_share_session(token) if token else None

            # 학생증을 스캔해서 받은 주소로만 들어올 수 있다
            if session is None:
                self._deny()
                return

            label = f"{session['number']} {session['name']}".strip()
            body = UPLOAD_PAGE_HTML.replace("{{USER}}", label).encode("utf-8")

            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def do_POST(self):
            from urllib.parse import urlparse

            if urlparse(self.path).path != "/upload":
                self.send_error(404)
                return

            token = self._token()

            if get_share_session(token) is None:
                self.send_error(403, "Forbidden")
                return

            try:
                content_length = int(self.headers.get("Content-Length", 0))
            except (TypeError, ValueError):
                content_length = 0

            if content_length > MAX_UPLOAD_TOTAL:
                limit_mb = MAX_UPLOAD_TOTAL // (1024 * 1024)
                body = f"TOOBIG:{limit_mb}".encode("utf-8")
                self.send_response(413)
                self.send_header("Content-Type", "text/plain; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return

            try:
                saved_paths, rejected = save_multipart_upload(
                    self.rfile,
                    self.headers.get("Content-Type", ""),
                    content_length,
                    get_download_dir(),
                    token=token
                )

                if not saved_paths:
                    # 전부 지원하지 않는 형식이었으면 그 사실을 알려준다
                    if rejected:
                        body = ("UNSUPPORTED:" + ", ".join(rejected[:5])).encode("utf-8")
                        self.send_response(415)
                        self.send_header("Content-Type", "text/plain; charset=utf-8")
                        self.send_header("Content-Length", str(len(body)))
                        self.end_headers()
                        self.wfile.write(body)
                        return

                    self.send_error(400, "No file")
                    return

                notify_upload(token, saved_paths)

                # "저장된 개수|거른 파일 이름" 형식으로 알려준다
                text = str(len(saved_paths))
                if rejected:
                    text += "|" + ", ".join(rejected[:5])

                body = text.encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/plain; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            except Exception as e:
                try:
                    self.send_error(500, "Upload failed")
                except Exception:
                    pass

    try:
        server = ThreadingHTTPServer(("0.0.0.0", FILE_SHARE_PORT), Handler)
    except OSError as e:
        return f"{FILE_SHARE_PORT}번 포트를 열 수 없습니다.\n{e}"

    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    _upload_server = server
    _upload_server_thread = thread
    return None


def stop_upload_server():
    global _upload_server, _upload_server_thread

    if _upload_server is None:
        return

    try:
        _upload_server.shutdown()
        _upload_server.server_close()
    except Exception:
        pass

    _upload_server = None
    _upload_server_thread = None


def make_qr_image(text, size=200):
    """
    QR 코드 이미지를 만든다. qrcode 패키지가 없으면 None 을 돌려준다.
    (그 경우 화면에는 주소만 글자로 표시된다)
    """
    try:
        import qrcode
    except ImportError:
        return None

    try:
        qr = qrcode.QRCode(box_size=10, border=2)
        qr.add_data(text)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white").convert("RGB")

        # QR 은 픽셀이 뭉개지면 인식률이 떨어지므로 보간 없이 확대한다
        if Image is not None:
            return img.resize((size, size), Image.NEAREST)

        return img
    except Exception:
        return None


FIREWALL_RULE_NAME = "School Printer Upload"


def is_firewall_rule_present():
    """파일 공유용 방화벽 규칙이 등록되어 있는지 확인한다."""
    try:
        result = subprocess.run(
            ["netsh", "advfirewall", "firewall", "show", "rule",
             f"name={FIREWALL_RULE_NAME}"],
            capture_output=True,
            timeout=15,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
    except Exception:
        return False

    if result.returncode != 0:
        return False

    output = result.stdout.decode("cp949", errors="replace")

    # 규칙이 있어도 포트가 바뀌었으면 다시 만들어야 한다
    return str(FILE_SHARE_PORT) in output


def is_running_as_admin():
    """지금 프로그램이 관리자 권한으로 돌고 있는지."""
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _run_elevated_and_wait(command, timeout=120):
    """
    관리자 권한으로 명령을 실행하고 끝날 때까지 기다린다. (UAC 창이 뜬다)

    ShellExecuteW 로 던지기만 하면 성공했는지 알 수 없어서, 여기서는
    ShellExecuteEx 로 프로세스 손잡이를 받아 종료 코드까지 확인한다.
    돌려주는 값은 (성공여부, 안내문).
    """
    import ctypes
    import ctypes.wintypes as wt

    SEE_MASK_NOCLOSEPROCESS = 0x00000040
    SEE_MASK_NO_CONSOLE = 0x00008000
    SW_HIDE = 0
    ERROR_CANCELLED = 1223

    class SHELLEXECUTEINFOW(ctypes.Structure):
        _fields_ = [
            ("cbSize", wt.DWORD),
            ("fMask", ctypes.c_ulong),
            ("hwnd", wt.HWND),
            ("lpVerb", wt.LPCWSTR),
            ("lpFile", wt.LPCWSTR),
            ("lpParameters", wt.LPCWSTR),
            ("lpDirectory", wt.LPCWSTR),
            ("nShow", ctypes.c_int),
            ("hInstApp", wt.HINSTANCE),
            ("lpIDList", ctypes.c_void_p),
            ("lpClass", wt.LPCWSTR),
            ("hkeyClass", wt.HKEY),
            ("dwHotKey", wt.DWORD),
            ("hIcon", wt.HANDLE),
            ("hProcess", wt.HANDLE),
        ]

    info = SHELLEXECUTEINFOW()
    info.cbSize = ctypes.sizeof(info)
    info.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NO_CONSOLE
    info.lpVerb = "runas"
    info.lpFile = "cmd.exe"
    info.lpParameters = f"/c {command}"
    info.nShow = SW_HIDE

    try:
        ok = ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(info))
    except Exception as e:
        return False, f"명령을 실행하지 못했습니다.\n{e}"

    if not ok or not info.hProcess:
        code = ctypes.GetLastError()

        if code == ERROR_CANCELLED:
            return False, "관리자 권한 요청이 거부되었습니다."

        return False, f"명령을 실행하지 못했습니다. (코드 {code})"

    kernel32 = ctypes.windll.kernel32

    try:
        wait_result = kernel32.WaitForSingleObject(info.hProcess, int(timeout * 1000))

        if wait_result != 0:   # WAIT_OBJECT_0 가 아니면 시간 초과
            return False, "작업이 끝나기를 기다리다 시간이 초과되었습니다."

        exit_code = wt.DWORD()
        kernel32.GetExitCodeProcess(info.hProcess, ctypes.byref(exit_code))
    finally:
        kernel32.CloseHandle(info.hProcess)

    if exit_code.value != 0:
        return False, f"명령이 실패했습니다. (코드 {exit_code.value})"

    return True, ""


def run_admin_commands(lines, timeout=120):
    """
    여러 명령을 한 번의 관리자 권한 요청으로 처리한다.

    UAC 창이 명령마다 뜨면 쓰기 어려우므로, 임시 배치 파일 하나에 모아
    한 번만 물어본다. 명령의 출력은 파일로 받아 실패했을 때 보여준다.
    돌려주는 값은 (성공여부, 안내문).
    """
    work_dir = os.path.join(tempfile.gettempdir(), "school_printer_admin")
    os.makedirs(work_dir, exist_ok=True)

    stamp = int(time.time() * 1000)
    bat_path = os.path.join(work_dir, f"run_{stamp}.bat")
    log_path = os.path.join(work_dir, f"run_{stamp}.log")

    body = "\r\n".join(lines)

    with open(bat_path, "w", encoding="cp949", errors="replace", newline="\r\n") as f:
        f.write("@echo off\r\n")
        f.write(f'set "SP_LOG={log_path}"\r\n')
        f.write(body + "\r\n")

    try:
        # 이미 관리자 권한이면 UAC 를 다시 물어보지 않는다
        if is_running_as_admin():
            result = subprocess.run(
                ["cmd.exe", "/c", bat_path],
                capture_output=True,
                timeout=timeout,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
            )
            ok = result.returncode == 0
            message = "" if ok else f"명령이 실패했습니다. (코드 {result.returncode})"
        else:
            ok, message = _run_elevated_and_wait(f'"{bat_path}"', timeout=timeout)

        if not ok:
            detail = ""
            try:
                if os.path.exists(log_path):
                    detail = open(log_path, encoding="cp949", errors="replace").read().strip()
            except OSError:
                detail = ""

            if detail:
                message = (message + "\n\n" + detail[:400]).strip()

        return ok, message

    finally:
        for path in (bat_path, log_path):
            try:
                os.remove(path)
            except OSError:
                pass


def _run_as_admin(command):
    """
    관리자 권한으로 명령을 실행하고 결과를 확인한다.
    성공하면 None, 실패하면 오류 메시지를 돌려준다. (기존 호출부 호환용)
    """
    ok, message = run_admin_commands([f'{command} >> "%SP_LOG%" 2>&1'])
    return None if ok else (message or "명령을 실행하지 못했습니다.")


def add_firewall_rule():
    """
    파일 공유용 포트를 Windows 방화벽에서 허용한다.
    포트가 바뀌었을 수 있으므로 같은 이름의 기존 규칙을 지우고 새로 만든다.
    성공하면 None, 실패하면 오류 메시지를 돌려준다.
    """
    command = (
        f'netsh advfirewall firewall delete rule name="{FIREWALL_RULE_NAME}" >nul 2>&1 & '
        f'netsh advfirewall firewall add rule '
        f'name="{FIREWALL_RULE_NAME}" dir=in action=allow '
        f'protocol=TCP localport={FILE_SHARE_PORT}'
    )
    return _run_as_admin(command)


def remove_firewall_rule():
    """등록해둔 방화벽 규칙을 지운다."""
    return _run_as_admin(
        f'netsh advfirewall firewall delete rule name="{FIREWALL_RULE_NAME}"'
    )


def _open_policy_key(hive_path, create=False):
    import winreg

    if create:
        return winreg.CreateKey(winreg.HKEY_CURRENT_USER, hive_path)

    return winreg.OpenKey(winreg.HKEY_CURRENT_USER, hive_path, 0, winreg.KEY_ALL_ACCESS)


def set_windows_lockdown(enabled):
    """
    작업 관리자 / 실행(Win+R) / 시작 화면(윈도우 키)을 잠그거나 푼다.

    - 작업 관리자, 실행창: HKEY_CURRENT_USER 정책값으로 즉시 적용된다.
    - 윈도우 키: 키보드 스캔코드 매핑이라 HKEY_LOCAL_MACHINE 수정 + 재부팅이 필요하다.
      (관리자 권한으로 실행해야 하며, 실패해도 나머지는 정상 적용된다)

    반환값: (성공 메시지 목록, 실패 메시지 목록)
    """
    import winreg

    done = []
    failed = []

    system_key = r"Software\Microsoft\Windows\CurrentVersion\Policies\System"
    explorer_key = r"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

    # 작업 관리자 차단
    try:
        with _open_policy_key(system_key, create=True) as key:
            winreg.SetValueEx(key, "DisableTaskMgr", 0, winreg.REG_DWORD, 1 if enabled else 0)
        done.append("작업 관리자")
    except Exception as e:
        failed.append(f"작업 관리자: {e}")

    # 실행창(Win+R) 차단
    try:
        with _open_policy_key(explorer_key, create=True) as key:
            winreg.SetValueEx(key, "NoRun", 0, winreg.REG_DWORD, 1 if enabled else 0)
        done.append("실행창(Win+R)")
    except Exception as e:
        failed.append(f"실행창: {e}")

    return done, failed


def is_windows_lockdown_enabled():
    import winreg

    try:
        system_key = r"Software\Microsoft\Windows\CurrentVersion\Policies\System"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, system_key) as key:
            value, _ = winreg.QueryValueEx(key, "DisableTaskMgr")
            return value == 1
    except Exception:
        return False


# -----------------------------
# 저수준 키보드 훅 (윈도우 키 / Alt+Tab / Ctrl+Shift+Esc 등 실시간 차단)
# 레지스트리를 건드리지 않고, 관리자 권한이나 재부팅 없이 즉시 적용된다.
# 프로그램이 종료되면 자동으로 원래대로 돌아온다.
# -----------------------------
_KEY_HOOK_THREAD = None
_KEY_HOOK_ACTIVE = False
_KEY_HOOK_CALLBACK = None      # 콜백이 GC되면 훅이 죽으므로 전역으로 잡아둔다
_KEY_HOOK_LAST_ERROR = ""


def get_key_blocker_error():
    return _KEY_HOOK_LAST_ERROR


def start_key_blocker():
    """
    윈도우 키 등 시스템 단축키를 차단하는 훅을 백그라운드에서 실행한다.
    성공하면 True, 실패하면 False를 반환한다 (실패 사유는 get_key_blocker_error()).
    """
    global _KEY_HOOK_THREAD, _KEY_HOOK_ACTIVE, _KEY_HOOK_CALLBACK, _KEY_HOOK_LAST_ERROR

    if _KEY_HOOK_ACTIVE:
        return True

    _KEY_HOOK_LAST_ERROR = ""

    try:
        import ctypes
        import ctypes.wintypes as wintypes
        import threading
    except ImportError as e:
        _KEY_HOOK_LAST_ERROR = f"모듈 없음: {e}"
        return False

    VK_LWIN, VK_RWIN = 0x5B, 0x5C
    VK_TAB, VK_ESCAPE, VK_F4 = 0x09, 0x1B, 0x73
    VK_CONTROL, VK_SHIFT = 0x11, 0x10
    WH_KEYBOARD_LL = 13
    LLKHF_ALTDOWN = 0x20

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    class KBDLLHOOKSTRUCT(ctypes.Structure):
        _fields_ = [
            ("vkCode", wintypes.DWORD),
            ("scanCode", wintypes.DWORD),
            ("flags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", ctypes.c_void_p),
        ]

    LRESULT = ctypes.c_ssize_t
    LPKBDLLHOOKSTRUCT = ctypes.POINTER(KBDLLHOOKSTRUCT)
    HOOKPROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_int, wintypes.WPARAM, LPKBDLLHOOKSTRUCT)

    # 64비트에서 반환형을 지정하지 않으면 핸들이 32비트로 잘려서
    # SetWindowsHookEx가 실패한다. 반드시 restype/argtypes를 명시해야 한다.
    user32.SetWindowsHookExW.restype = wintypes.HHOOK
    user32.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC, wintypes.HINSTANCE, wintypes.DWORD]

    user32.CallNextHookEx.restype = LRESULT
    user32.CallNextHookEx.argtypes = [wintypes.HHOOK, ctypes.c_int, wintypes.WPARAM, LPKBDLLHOOKSTRUCT]

    user32.UnhookWindowsHookEx.restype = wintypes.BOOL
    user32.UnhookWindowsHookEx.argtypes = [wintypes.HHOOK]

    user32.GetAsyncKeyState.restype = ctypes.c_short
    user32.GetAsyncKeyState.argtypes = [ctypes.c_int]

    user32.PeekMessageW.restype = wintypes.BOOL
    user32.PeekMessageW.argtypes = [
        ctypes.POINTER(wintypes.MSG), wintypes.HWND,
        wintypes.UINT, wintypes.UINT, wintypes.UINT
    ]

    kernel32.GetModuleHandleW.restype = wintypes.HMODULE
    kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]

    hook_holder = {"id": None}

    def is_down(vk):
        return bool(user32.GetAsyncKeyState(vk) & 0x8000)

    def low_level_handler(nCode, wParam, lParam):
        if nCode == 0:
            info = lParam[0]
            vk = info.vkCode
            alt_down = bool(info.flags & LLKHF_ALTDOWN)

            # 윈도우 키 (단독 및 Win+R, Win+E 등 모든 조합)
            if vk in (VK_LWIN, VK_RWIN):
                return 1

            # Alt+Tab / Alt+Esc / Alt+F4
            if alt_down and vk in (VK_TAB, VK_ESCAPE, VK_F4):
                return 1

            # Ctrl+Esc (시작 화면), Ctrl+Shift+Esc (작업 관리자)
            if vk == VK_ESCAPE and is_down(VK_CONTROL):
                return 1

        return user32.CallNextHookEx(hook_holder["id"], nCode, wParam, lParam)

    _KEY_HOOK_CALLBACK = HOOKPROC(low_level_handler)

    ready = threading.Event()

    def hook_loop():
        global _KEY_HOOK_ACTIVE, _KEY_HOOK_LAST_ERROR

        hook_id = user32.SetWindowsHookExW(
            WH_KEYBOARD_LL, _KEY_HOOK_CALLBACK, kernel32.GetModuleHandleW(None), 0
        )

        if not hook_id:
            _KEY_HOOK_LAST_ERROR = f"SetWindowsHookEx 실패 (코드 {ctypes.get_last_error()})"
            _KEY_HOOK_ACTIVE = False
            ready.set()
            return

        hook_holder["id"] = hook_id
        _KEY_HOOK_ACTIVE = True
        ready.set()

        msg = wintypes.MSG()
        while _KEY_HOOK_ACTIVE:
            # PM_REMOVE = 1
            while user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            time.sleep(0.005)

        user32.UnhookWindowsHookEx(hook_id)
        hook_holder["id"] = None

    _KEY_HOOK_THREAD = threading.Thread(target=hook_loop, daemon=True)
    _KEY_HOOK_THREAD.start()

    ready.wait(timeout=3)
    return _KEY_HOOK_ACTIVE


def stop_key_blocker():
    global _KEY_HOOK_ACTIVE
    _KEY_HOOK_ACTIVE = False


# -----------------------------
# 자동 종료 (Windows 작업 스케줄러)
#
# 프로그램이 꺼져 있어도 정해진 시간에 PC가 종료되도록
# Windows 작업 스케줄러에 일정을 등록한다.
# -----------------------------
SHUTDOWN_TASK_PREFIX = "SchoolPrinter_Shutdown_"

# 요일 코드 (schtasks 용) 와 화면 표시용 이름
WEEKDAY_CODES = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
WEEKDAY_LABELS = ["월", "화", "수", "목", "금", "토", "일"]


def _run_schtasks(args):
    """
    schtasks 명령을 실행하고 (성공여부, 출력) 을 돌려준다.
    조회처럼 권한이 필요 없는 명령에 쓴다.
    """
    try:
        result = subprocess.run(
            ["schtasks"] + args,
            capture_output=True,
            text=True,
            encoding="cp949",
            errors="ignore",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = (result.stdout or "") + (result.stderr or "")
        return result.returncode == 0, output.strip()
    except Exception as e:
        return False, str(e)


def _quote_schtasks(args):
    """
    schtasks 인자를 배치 파일에 넣을 수 있게 다듬는다.
    /Create 같은 옵션은 그대로 두고, 값은 따옴표로 감싼다.
    """
    parts = []

    for arg in args:
        text = str(arg)
        parts.append(text if text.startswith("/") else f'"{text}"')

    return " ".join(parts)


def _run_schtasks_admin(arg_sets, timeout=120):
    """
    작업 스케줄러 등록·삭제를 관리자 권한으로 실행한다.

    작업 등록(/RL HIGHEST)은 관리자 권한이 필요해서, 예전에는 프로그램 자체를
    "관리자 권한으로 실행"해야만 동작했다. 방화벽 설정과 같은 방식으로
    그때그때 UAC 창을 띄워 처리한다.

    arg_sets 는 [[schtasks 인자...], ...] 형태이며, 앞에서부터 차례로 시도해
    하나라도 성공하면 성공으로 본다. (날짜 형식이 Windows 언어 설정마다 다르기
    때문에 여러 형식을 시도해야 한다)
    """
    lines = []

    for i, args in enumerate(arg_sets):
        command = "schtasks " + _quote_schtasks(args)
        lines.append(f'{command} >> "%SP_LOG%" 2>&1')
        lines.append("if not errorlevel 1 goto sp_done")

    lines.append("exit /b 1")
    lines.append(":sp_done")
    lines.append("exit /b 0")

    return run_admin_commands(lines, timeout=timeout)


SHUTDOWN_SCRIPT_NAME = "auto_shutdown.bat"


def get_shutdown_script_path():
    return os.path.join(get_program_dir(), SHUTDOWN_SCRIPT_NAME)


def ensure_shutdown_script():
    """
    자동 종료에 쓸 배치 파일을 만들어 둔다.

    schtasks 의 /TR 에 따옴표가 들어간 명령을 그대로 넣으면 Windows 언어 설정과
    따옴표 중첩 때문에 등록이 실패하는 경우가 많다. 명령을 배치 파일 하나로 빼고
    그 경로만 등록하면 이런 문제가 없다.
    """
    path = get_shutdown_script_path()

    body = (
        "@echo off\r\n"
        'shutdown.exe /s /f /t 60 /c "자동 종료 시간입니다. 1분 후 컴퓨터가 꺼집니다."\r\n'
    )

    try:
        if os.path.exists(path) and open(path, encoding="cp949", errors="replace").read() == body:
            return path

        with open(path, "w", encoding="cp949", errors="replace", newline="") as f:
            f.write(body)
    except OSError:
        pass

    return path


def _shutdown_command():
    return ensure_shutdown_script()


def format_day_text(day_codes):
    """요일 목록을 읽기 쉬운 문자열로 바꾼다."""
    if not day_codes:
        return ""

    if len(day_codes) == 7:
        return "매일"

    if sorted(day_codes) == sorted(WEEKDAY_CODES[:5]):
        return "평일 (월~금)"

    if sorted(day_codes) == sorted(WEEKDAY_CODES[5:]):
        return "주말 (토·일)"

    labels = [
        WEEKDAY_LABELS[WEEKDAY_CODES.index(code)]
        for code in WEEKDAY_CODES
        if code in day_codes
    ]
    return "·".join(labels)


def _register_task(task_name, kind, days, run_date, run_time):
    """
    schtasks 에 작업을 등록한다.

    특정 날짜(once)의 경우 날짜 형식이 Windows 언어 설정마다 달라서
    여러 형식을 차례로 시도한다.
    """
    base_args = [
        "/Create",
        "/TN", task_name,
        "/TR", _shutdown_command(),
        "/ST", run_time,
        "/RL", "HIGHEST",
        "/F",
    ]

    if kind == "weekly":
        if not days or len(days) == 7:
            schedule_args = ["/SC", "DAILY"]
        else:
            ordered = [d for d in WEEKDAY_CODES if d in days]
            schedule_args = ["/SC", "WEEKLY", "/D", ",".join(ordered)]

        return _run_schtasks_admin([base_args + schedule_args])

    # 특정 날짜 1회 실행.
    # 날짜 형식은 Windows 언어 설정마다 다르므로 여러 형식을 한 번의 권한 요청
    # 안에서 차례로 시도한다. (형식마다 UAC 창이 뜨면 쓸 수 없다)
    year, month, day = run_date.split("-")
    date_formats = [
        f"{year}-{month}-{day}",
        f"{month}/{day}/{year}",
        f"{day}/{month}/{year}",
        f"{year}/{month}/{day}",
    ]

    return _run_schtasks_admin([
        base_args + ["/SC", "ONCE", "/SD", date_text]
        for date_text in date_formats
    ])


def add_shutdown_schedule(kind, days, run_date, run_time):
    """
    자동 종료 일정을 하나 추가한다.
    kind : 'weekly' 또는 'once'
    days : 'weekly' 일 때 요일 코드 목록
    run_date : 'once' 일 때 'YYYY-MM-DD'
    """
    conn = connect_db()
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO shutdown_schedules(kind, days, run_date, run_time, created_date)
    VALUES (?, ?, ?, ?, ?)
    """, (
        kind,
        ",".join(days) if days else "",
        run_date or "",
        run_time,
        str(date.today())
    ))
    schedule_id = cur.lastrowid
    conn.commit()
    conn.close()

    task_name = f"{SHUTDOWN_TASK_PREFIX}{schedule_id}"
    ok, output = _register_task(task_name, kind, days, run_date, run_time)

    if not ok:
        # 등록에 실패하면 DB 기록도 되돌린다
        conn = connect_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM shutdown_schedules WHERE id=?", (schedule_id,))
        conn.commit()
        conn.close()

    return ok, output


def list_shutdown_schedules():
    """등록된 자동 종료 일정을 돌려준다."""
    conn = connect_db()
    cur = conn.cursor()
    cur.execute("""
    SELECT id, kind, days, run_date, run_time
    FROM shutdown_schedules
    ORDER BY kind, run_date, run_time
    """)
    rows = cur.fetchall()
    conn.close()
    return rows


def delete_shutdown_schedule(schedule_id):
    """
    일정 하나를 지운다 (작업 스케줄러 + DB).
    돌려주는 값은 (작업 스케줄러 삭제 성공여부, 안내문).
    """
    ok, output = _run_schtasks_admin(
        [["/Delete", "/TN", f"{SHUTDOWN_TASK_PREFIX}{schedule_id}", "/F"]]
    )

    conn = connect_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM shutdown_schedules WHERE id=?", (schedule_id,))
    conn.commit()
    conn.close()

    return ok, output


def delete_all_shutdown_schedules():
    """모든 일정을 지운다. (권한 요청은 한 번만 뜬다)"""
    names = [f"{SHUTDOWN_TASK_PREFIX}{row[0]}" for row in list_shutdown_schedules()]

    # 예전 버전에서 만든 작업도 함께 정리
    names.append("SchoolPrinter_AutoShutdown")

    lines = [
        f'schtasks /Delete /TN "{name}" /F >> "%SP_LOG%" 2>&1'
        for name in names
    ]
    lines.append("exit /b 0")   # 이미 없는 작업은 실패로 보지 않는다

    if lines:
        run_admin_commands(lines)

    conn = connect_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM shutdown_schedules")
    conn.commit()
    conn.close()

    return True


def cleanup_past_schedules():
    """
    이미 지나간 '특정 날짜' 일정을 목록에서 정리한다.
    (작업 스케줄러는 1회 실행 후 자동으로 비활성화되지만
     목록에 계속 남아 있으면 지저분해지므로 정리한다)
    """
    today = str(date.today())

    conn = connect_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT id FROM shutdown_schedules WHERE kind='once' AND run_date < ?",
        (today,)
    )
    old_ids = [row[0] for row in cur.fetchall()]
    conn.close()

    for schedule_id in old_ids:
        delete_shutdown_schedule(schedule_id)

    return len(old_ids)


def describe_schedule(kind, days, run_date, run_time):
    """일정을 화면에 표시할 문자열로 만든다."""
    if kind == "weekly":
        day_list = [d for d in days.split(",") if d] if days else []
        return format_day_text(day_list) or "매일", run_time

    # 특정 날짜
    try:
        d = datetime.strptime(run_date, "%Y-%m-%d").date()
        weekday = WEEKDAY_LABELS[d.weekday()]
        return f"{d.month}월 {d.day}일 ({weekday})", run_time
    except Exception:
        return run_date, run_time


# -----------------------------
# 자동 업데이트
# -----------------------------
def _parse_version(text):
    """'v1.2.3' 또는 '1.2.3' 을 (1, 2, 3) 으로 바꾼다."""
    text = str(text or "").strip().lstrip("vV")
    parts = []

    for chunk in text.split("."):
        digits = "".join(c for c in chunk if c.isdigit())
        parts.append(int(digits) if digits else 0)

    while len(parts) < 3:
        parts.append(0)

    return tuple(parts[:3])


def check_for_update():
    """
    GitHub 최신 릴리스를 확인한다.

    반환값: (상태, 정보)
      ("최신", None)                    이미 최신 버전
      ("있음", {버전, 다운로드주소, 설명})  새 버전 있음
      ("오류", 메시지)                   확인 실패
    """
    import json
    import ssl
    import urllib.request

    url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

    try:
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE

        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": "SchoolPrinterKiosk",
                "Accept": "application/vnd.github+json",
            },
        )

        with urllib.request.urlopen(request, context=context, timeout=15) as response:
            data = json.loads(response.read().decode("utf-8"))

    except Exception as e:
        message = str(e)

        if "404" in message:
            return "오류", (
                "저장소에서 릴리스를 찾을 수 없습니다.\n\n"
                "· 아직 배포된 버전이 없거나\n"
                "· 저장소가 비공개로 설정되어 있거나\n"
                "· 저장소 주소가 잘못되었습니다."
            )

        if "403" in message:
            return "오류", (
                "업데이트 서버 접속이 거부되었습니다.\n\n"
                "잠시 후 다시 시도하거나, 학교 네트워크에서\n"
                "GitHub 접속이 차단되어 있는지 확인해 주세요."
            )

        return "오류", f"업데이트 서버에 연결하지 못했습니다.\n\n{message}"

    latest_tag = data.get("tag_name", "")

    if not latest_tag:
        return "오류", "릴리스 정보를 읽을 수 없습니다."

    if _parse_version(latest_tag) <= _parse_version(APP_VERSION):
        return "최신", None

    # 첨부된 exe 파일을 찾는다
    download_url = ""
    file_name = ""

    for asset in data.get("assets", []):
        name = asset.get("name", "")
        if name.lower().endswith(".exe"):
            download_url = asset.get("browser_download_url", "")
            file_name = name
            break

    if not download_url:
        return "오류", (
            f"새 버전({latest_tag})이 있지만 실행 파일이 첨부되어 있지 않습니다.\n"
            "릴리스에 exe 파일을 올려 주세요."
        )

    return "있음", {
        "version": latest_tag.lstrip("vV"),
        "url": download_url,
        "file_name": file_name,
        "notes": (data.get("body") or "").strip(),
    }


def download_update(url, progress_callback=None):
    """새 버전 파일을 임시 폴더에 내려받고 그 경로를 돌려준다."""
    import ssl
    import urllib.request

    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE

    request = urllib.request.Request(
        url, headers={"User-Agent": "SchoolPrinterKiosk"}
    )

    out_dir = os.path.join(tempfile.gettempdir(), "school_printer_update")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "new_version.exe")

    with urllib.request.urlopen(request, context=context, timeout=120) as response:
        total = int(response.headers.get("Content-Length", 0))
        downloaded = 0
        last_percent = -1

        with open(out_path, "wb") as f:
            while True:
                chunk = response.read(256 * 1024)
                if not chunk:
                    break

                f.write(chunk)
                downloaded += len(chunk)

                if total and progress_callback:
                    percent = int(downloaded * 100 / total)
                    if percent >= last_percent + 5:
                        progress_callback(percent)
                        last_percent = percent

    return out_path


def apply_update(new_file_path):
    """
    내려받은 파일로 현재 실행 파일을 교체한다.

    실행 중인 exe 는 자기 자신을 덮어쓸 수 없다.
    그래서 교체를 대신 해줄 배치 파일을 만들어 띄우고 프로그램은 종료한다.

    배치 파일이 하는 일:
      1) 프로그램 프로세스가 완전히 사라질 때까지 기다림
      2) PyInstaller 가 남긴 임시 폴더(_MEIxxxx) 정리
      3) 기존 exe 를 새 파일로 교체
      4) 잠시 기다린 뒤 새 버전 실행
      5) 자기 자신 삭제

    2번이 없으면 새 exe 가 예전 임시 폴더를 참조하다가
    "Failed to load Python DLL" 오류를 낸다.
    """
    if not getattr(sys, "frozen", False):
        raise RuntimeError(
            "exe 로 빌드한 상태에서만 업데이트할 수 있습니다.\n"
            "파이썬 파일로 실행 중일 때는 코드를 직접 받아 주세요."
        )

    current_exe = sys.executable
    backup_exe = current_exe + ".old"
    exe_name = os.path.basename(current_exe)
    work_dir = os.path.dirname(current_exe)

    bat_path = os.path.join(
        tempfile.gettempdir(), "school_printer_update", "apply_update.bat"
    )

    script = f"""@echo off
chcp 65001 > nul

rem ── 1. 프로그램이 완전히 종료될 때까지 기다린다 ──
set /a tries=0
:waitloop
tasklist /fi "IMAGENAME eq {exe_name}" 2>nul | find /i "{exe_name}" >nul
if not errorlevel 1 (
    set /a tries+=1
    if !tries! GEQ 30 goto forcekill
    timeout /t 1 /nobreak > nul
    goto waitloop
)
goto killed

:forcekill
taskkill /f /im "{exe_name}" > nul 2>&1
timeout /t 2 /nobreak > nul

:killed
rem 파일 핸들이 완전히 풀릴 때까지 여유를 준다
timeout /t 2 /nobreak > nul

rem ── 2. PyInstaller 임시 폴더 정리 ──
for /d %%D in ("%TEMP%\\_MEI*") do rd /s /q "%%D" > nul 2>&1

rem ── 3. 기존 파일 교체 ──
set /a swap=0
:retry
del "{backup_exe}" > nul 2>&1
move "{current_exe}" "{backup_exe}" > nul 2>&1
if errorlevel 1 (
    set /a swap+=1
    if !swap! GEQ 15 goto failed
    timeout /t 2 /nobreak > nul
    goto retry
)

move "{new_file_path}" "{current_exe}" > nul 2>&1
if errorlevel 1 goto rollback

del "{backup_exe}" > nul 2>&1

rem ── 4. 새 버전 실행 ──
timeout /t 2 /nobreak > nul
cd /d "{work_dir}"
start "" "{current_exe}"
goto cleanup

:rollback
move "{backup_exe}" "{current_exe}" > nul 2>&1
:failed
echo.
echo  업데이트에 실패했습니다.
echo  프로그램을 직접 실행해 주세요.
echo.
pause
goto cleanup

:cleanup
del "%~f0" > nul 2>&1
"""

    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(script)

    # setlocal enabledelayedexpansion 이 필요하므로 cmd /v:on 으로 실행한다
    subprocess.Popen(
        ["cmd", "/v:on", "/c", bat_path],
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def prepare_browser_download_dir(profile_dir, download_dir):
    """
    크롬/엣지 프로필에 다운로드 폴더를 지정한다.

    다운로드 위치는 명령줄 옵션으로 줄 수 없고 프로필의 Preferences 파일에
    들어 있다. 그래서 브라우저를 띄우기 전에 그 파일을 직접 써준다.
    저장 위치를 묻는 창도 뜨지 않게 해서 학생이 헤매지 않도록 한다.
    """
    import json

    try:
        os.makedirs(download_dir, exist_ok=True)
    except Exception:
        pass

    pref_path = os.path.join(profile_dir, "Default", "Preferences")

    try:
        os.makedirs(os.path.dirname(pref_path), exist_ok=True)
    except Exception:
        return

    data = {}
    if os.path.exists(pref_path):
        try:
            with open(pref_path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}

    download = data.get("download", {})
    download["default_directory"] = download_dir
    download["prompt_for_download"] = False
    download["directory_upgrade"] = True
    data["download"] = download

    # 첫 실행 안내 화면이 뜨지 않게 한다
    data.setdefault("profile", {})["exit_type"] = "Normal"
    data["profile"]["exited_cleanly"] = True

    try:
        with open(pref_path, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        pass


def resolve_font_family(root):
    """
    Pretendard 폰트가 시스템에 설치되어 있으면 그것을 쓰고,
    없으면 맑은 고딕으로 자동 대체한다.
    """
    try:
        families = set(tkfont.families(root))
    except Exception:
        families = set()

    for candidate in ("Pretendard", "Pretendard Variable", "Pretendard JP"):
        if candidate in families:
            return candidate

    return FALLBACK_FONT_FAMILY


class PrinterKioskApp:
    def __init__(self, root):
        self.root = root
        self.root.title("서인천고등학교 공용 프린터 제어 시스템")
        apply_window_icon(self.root)

        self.root.update_idletasks()
        self.font_family = resolve_font_family(self.root)
        self._setup_treeview_style()

        init_db()

        self.external_browser_open = False
        self._last_external_open = 0
        self._external_process = None
        self._close_overlay = None
        self._overlay_job = None
        self._auto_delete_overlay = None
        self._auto_delete_overlay_job = None

        # 페이지 수를 셀 때 만든 변환 PDF 를 인쇄 때 재사용하기 위한 캐시
        # {원본파일경로: 변환된PDF경로}
        self._print_converted_cache = {}

        # 첫 화면에서 학생증을 스캔하면 여기에 담아두고,
        # 파일 공유를 누를 때 다시 스캔하지 않아도 되게 한다.
        self._current_user = None

        self.fullscreen = get_setting("fullscreen", "1") == "1"
        self.apply_fullscreen()

        self.root.protocol("WM_DELETE_WINDOW", self.block_close)
        self.bind_blocked_shortcuts()
        self.setup_scanner_focus()

        # 윈도우 키 등 시스템 단축키 차단 (설정에 따라)
        if get_setting("block_system_keys", "0") == "1":
            start_key_blocker()

        # 받은 파일 자동 삭제 감시 시작
        self.start_download_cleanup()

        # 이용자 명단 CSV 실시간 연동 시작
        self._user_list_refresh = None
        self.start_user_csv_sync()

        # 지난번에 남은 변환·미리보기 임시 파일 정리
        try:
            clean_temp_dirs()
        except Exception:
            pass

        # 학생이 화면을 켜둔 채 자리를 뜨면 처음 화면으로 되돌린다
        self._share_token = None
        self.start_idle_watch()

        # 접수된 고장 신고가 있으면 화면에 안내를 띄운다
        self._report_banner = None
        self._report_refresh_button = None
        self._report_checking = False
        self._report_state = None
        self.start_report_watch()

        self.start_screen()

    def start_report_watch(self):
        """
        접수된 고장 신고가 있는지 주기적으로 확인한다.

        학생들이 이미 신고한 문제를 또 신고하거나, 고장인 줄 모르고 계속
        시도하는 일을 줄이기 위한 안내다. 신고가 있어도 출력은 그대로 할 수
        있게 두고, 화면에 알림만 띄운다.
        """
        if not report_status_enabled():
            return

        def worker():
            while True:
                state = fetch_report_status()

                if state is not None and not self.direct_result_is_fresher(state):
                    # 화면 조작은 반드시 메인 스레드에서 해야 한다
                    try:
                        self.root.after(0, lambda s=state: self.apply_report_banner(s))
                    except Exception:
                        return

                time.sleep(max(15, REPORT_CHECK_SECONDS))

        threading.Thread(target=worker, daemon=True).start()

    # 직접 확인(Apps Script)한 결과를 이 시간(초) 동안은 GitHub 파일보다 믿는다.
    # 직접 확인할 때 Apps Script 가 파일도 고쳐 쓰지만, 반영되기까지 잠깐 걸린다.
    # 그 사이 옛 파일 내용으로 안내가 다시 떴다 사라지는 일을 막는다.
    DIRECT_RESULT_HOLD_SECONDS = 90

    def remember_direct_result(self, state):
        """방금 직접 확인한 결과를 기억해 둔다."""
        self._direct_state = state
        self._direct_state_at = time.time()

    def direct_result_is_fresher(self, file_state):
        """
        GitHub 파일에서 읽은 값을 무시해야 하는지.
        직접 확인한 지 얼마 안 됐는데 파일이 다른 말을 하면, 파일이 아직
        옛 내용인 것으로 보고 무시한다.
        """
        direct = getattr(self, "_direct_state", None)
        at = getattr(self, "_direct_state_at", 0)

        if direct is None or time.time() - at > self.DIRECT_RESULT_HOLD_SECONDS:
            return False

        def key(state):
            return (state.get("open", 0), state.get("status", ""), str(state.get("id", "")))

        return key(direct) != key(file_state)

    def check_report_status_soon(self, delay_seconds=25):
        """
        고장 신고 창을 열고 돌아온 직후처럼, 곧 상태가 바뀔 만한 때에
        한 번 더 확인한다. (다음 정기 확인까지 기다리지 않도록)
        """
        if not report_status_enabled():
            return

        def worker():
            time.sleep(delay_seconds)
            # 신고 직후에는 Apps Script 에 직접 물어본다 (파일 반영보다 빠르다)
            state = fetch_report_status(direct=True)

            if state is None:
                return

            self.remember_direct_result(state)

            try:
                self.root.after(0, lambda s=state: self.apply_report_banner(s))
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def refresh_report_status(self):
        """
        지금 바로 고장 신고 상태를 확인한다.

        정기 확인(기본 60초)을 기다리지 않고 결과를 보고 싶을 때 쓴다.
        확인하는 동안 화면이 멈추지 않도록 별도 스레드에서 처리한다.
        """
        if not report_status_enabled():
            messagebox.showinfo(
                "안내",
                "고장 신고 상태 주소가 설정되어 있지 않습니다.\n"
                "프로그램 파일의 REPORT_STATUS_FILE 또는 REPORT_STATUS_URL 을 확인해 주세요."
            )
            return

        if getattr(self, "_report_checking", False):
            return

        self._report_checking = True
        self.set_refresh_button_text("확인 중...")

        def worker():
            # 수동 확인은 Apps Script 에 직접 물어본다
            state = fetch_report_status(direct=True)

            try:
                self.root.after(0, lambda s=state: self.finish_refresh(s))
            except Exception:
                self._report_checking = False

        threading.Thread(target=worker, daemon=True).start()

    def finish_refresh(self, state):
        """새로고침 결과를 화면에 반영한다."""
        self._report_checking = False

        if state is None:
            # 확인에 실패했으면 지금 안내를 그대로 두고 알리기만 한다
            self.set_refresh_button_text("새로고침")
            messagebox.showwarning(
                "확인 실패",
                "고장 신고 상태를 확인하지 못했습니다.\n"
                "인터넷 연결을 확인한 뒤 다시 시도해 주세요."
            )
            return

        self.remember_direct_result(state)
        self.apply_report_banner(state)
        self.set_refresh_button_text("새로고침")

    def set_refresh_button_text(self, text):
        button = getattr(self, "_report_refresh_button", None)

        if button is None:
            return

        try:
            button.configure(text=text)
        except Exception:
            self._report_refresh_button = None

    def apply_report_banner(self, state):
        """확인한 결과에 맞춰 안내를 띄우거나 지운다."""
        self._report_state = state

        if not state or state.get("open", 0) <= 0:
            self.hide_report_banner()
            return

        self.show_report_banner(state)

    def hide_report_banner(self):
        banner = getattr(self, "_report_banner", None)

        if banner is None:
            return

        self._report_banner = None
        self._report_refresh_button = None

        try:
            banner.destroy()
        except Exception:
            pass

    def render_report_banner(self):
        """
        지금 화면에 맞춰 고장 안내를 다시 그린다.

        첫 화면에서는 학생 눈에 잘 띄도록 카드 아래쪽에 크게 띄우고,
        다른 화면에서는 내용을 가리지 않도록 위쪽에 작게 띄운다.
        """
        state = getattr(self, "_report_state", None)

        if not state or state.get("open", 0) <= 0:
            self.hide_report_banner()
            return

        self.show_report_banner(state)

    def show_report_banner(self, state):
        """
        연한 붉은색 안내를 띄운다.

        카드 안에 넣으면 화면마다 높이가 달라져 아래가 잘릴 수 있으므로,
        창 위에 얹는 방식으로 띄운다.
        """
        status = state.get("status") or "접수"
        reason = state.get("reason") or ""
        count = state.get("open", 0)

        headline = "고장 신고가 접수되었습니다"

        if status == "처리중":
            headline = "고장 신고를 처리하는 중입니다"

        if count > 1:
            headline += f" (신고 {count}건)"

        detail = "출력은 그대로 할 수 있습니다."

        if reason:
            detail = f"신고 내용: {reason}"

        # 첫 화면에서는 크게, 다른 화면에서는 작게
        big = (getattr(self, "current_screen", "") == "start")

        if big:
            place_args = {"relx": 0.5, "rely": 0.75, "anchor": "n"}
            head_size, detail_size = 30, 19
            pad_x, pad_y = 40, 22
            btn_width, btn_height, btn_font = 130, 52, 17
        else:
            place_args = {"relx": 0.5, "rely": 0.022, "anchor": "n"}
            head_size, detail_size = 20, 14
            pad_x, pad_y = 26, 12
            btn_width, btn_height, btn_font = 94, 38, 14

        self.hide_report_banner()

        banner = ctk.CTkFrame(
            self.root,
            fg_color=COLOR_REPORT_BANNER_BG,
            border_color=COLOR_REPORT_BANNER_LINE,
            border_width=3 if big else 2,
            corner_radius=16
        )
        banner.place(**place_args)

        texts = ctk.CTkFrame(banner, fg_color="transparent")
        texts.grid(row=0, column=0, padx=(pad_x, 18), pady=pad_y, sticky="w")

        ctk.CTkLabel(
            texts,
            text="🔧  " + headline,
            font=(self.font_family, head_size, "bold"),
            text_color=COLOR_REPORT_BANNER_TEXT
        ).pack(anchor="w")

        ctk.CTkLabel(
            texts,
            text=detail,
            font=(self.font_family, detail_size),
            text_color=COLOR_REPORT_BANNER_TEXT
        ).pack(anchor="w", pady=(4 if big else 2, 0))

        if big and reason:
            # 큰 안내에서는 '출력은 그대로 가능' 을 한 줄 더 두어 확실히 알린다
            ctk.CTkLabel(
                texts,
                text="출력은 그대로 할 수 있습니다.",
                font=(self.font_family, detail_size),
                text_color=COLOR_REPORT_BANNER_TEXT
            ).pack(anchor="w", pady=(2, 0))

        # 완료 처리한 직후에 바로 확인하고 싶을 때 쓰는 버튼
        refresh = ctk.CTkButton(
            banner,
            text="새로고침",
            command=lambda: self.refresh_report_status(),
            width=btn_width, height=btn_height, corner_radius=10,
            font=(self.font_family, btn_font, "bold"),
            fg_color="#FFFFFF", hover_color="#FBE3E3",
            text_color=COLOR_REPORT_BANNER_TEXT,
            border_color=COLOR_REPORT_BANNER_LINE, border_width=1
        )
        refresh.grid(row=0, column=1, padx=(0, pad_x - 6), pady=pad_y)

        self._report_banner = banner
        self._report_refresh_button = refresh

        try:
            banner.lift()
        except Exception:
            pass

    def start_idle_watch(self):
        """
        일정 시간 아무 조작이 없으면 처음 화면으로 되돌린다.

        학생이 파일을 고르다 말고 자리를 뜨면 다음 사람이 그 화면을 그대로 보게
        되고, 남의 파일 목록이 노출되거나 그 사람 계정으로 출력될 수 있다.
        """
        self._last_activity = time.time()

        def mark_activity(event=None):
            self._last_activity = time.time()

        for sequence in ("<Any-KeyPress>", "<Any-Button>", "<Motion>", "<MouseWheel>"):
            try:
                self.root.bind_all(sequence, mark_activity, add="+")
            except Exception:
                pass

        def tick():
            try:
                minutes = get_idle_minutes()

                if minutes > 0 and self.current_screen != "start":
                    idle = time.time() - getattr(self, "_last_activity", time.time())

                    if idle > minutes * 60:
                        self._last_activity = time.time()
                        end_admin_session("자리 비움으로 자동 종료")
                        self.start_screen()
            except Exception:
                pass

            self.root.after(10_000, tick)

        self.root.after(10_000, tick)

    def setup_scanner_focus(self):
        """
        바코드 스캐너는 키보드처럼 입력되기 때문에, 입력창에 포커스가 없으면
        학생증을 찍어도 아무 데도 입력되지 않는다.
        시작 화면에서는 입력창을 클릭하지 않아도 바로 찍히도록
        (1) 포커스를 주기적으로 되돌려 놓고
        (2) 혹시 다른 곳에 입력이 들어가면 입력창으로 넘겨준다.
        """
        self.current_screen = ""
        self.card_entry = None

        self.root.bind_all("<Key>", self.redirect_key_to_scan_entry, add="+")
        self.root.after(400, self.keep_scan_focus)

    def focus_scan_entry(self):
        if self.current_screen != "start":
            return

        if self.card_entry is None or not self.card_entry.winfo_exists():
            return

        try:
            self.card_entry.focus_force()
            # 커서를 항상 끝으로 보내 스캔값이 중간에 끼어들지 않게 한다
            self.card_entry.icursor(tk.END)
        except Exception:
            pass

    def keep_scan_focus(self):
        """시작 화면이 떠 있는 동안 입력창 포커스를 유지한다."""
        try:
            if (
                self.current_screen == "start"
                and not self.external_browser_open
                and self.card_entry is not None
                and self.card_entry.winfo_exists()
            ):
                # 창이 활성 상태일 때만 (다른 프로그램 쓰는 중에 뺏지 않도록)
                if self.root.focus_displayof() is not None:
                    focused = self.root.focus_get()

                    if focused is not self.card_entry:
                        self.focus_scan_entry()
        except Exception:
            pass

        self.root.after(400, self.keep_scan_focus)

    def redirect_key_to_scan_entry(self, event):
        """
        시작 화면에서 입력창 밖으로 들어온 글자를 입력창으로 넘긴다.
        스캔 첫 글자가 유실되는 것을 막기 위한 안전장치다.
        """
        if self.current_screen != "start":
            return

        if self.card_entry is None or not self.card_entry.winfo_exists():
            return

        try:
            focused = self.root.focus_get()
        except Exception:
            return

        if focused is self.card_entry:
            return

        # 다른 입력창(관리자 로그인 등)이 떠 있으면 건드리지 않는다
        if isinstance(focused, (ctk.CTkEntry, tk.Entry)):
            return

        char = event.char

        if not char or not char.isprintable():
            return

        self.focus_scan_entry()
        self.card_entry.insert(tk.END, char)
        return "break"

    def apply_fullscreen(self):
        """
        전체화면을 적용한다.

        해상도를 직접 계산해서 geometry로 지정하면 모니터 배율(125%, 150%)이
        걸린 PC에서 실제 화면보다 크거나 작게 잡힌다.
        -fullscreen 속성을 쓰면 Windows가 알아서 실제 모니터 크기에 맞춰주므로
        해상도와 무관하게 정확히 채워진다.
        """
        # 창이 너무 작아지면 우측 하단 버튼이 흰색 카드 위로 겹친다.
        # 최소 크기를 정해서 그 아래로는 줄어들지 않게 한다.
        self.root.minsize(1100, 780)

        if self.fullscreen:
            self.root.attributes("-fullscreen", True)
            self.root.attributes("-topmost", True)
            self.root.focus_force()
        else:
            self.root.attributes("-fullscreen", False)
            self.root.attributes("-topmost", False)
            self.root.state("zoomed")

        self.root.update_idletasks()

    def bind_blocked_shortcuts(self):
        """
        프로그램 창 안에서 창을 벗어나게 하는 키 조합을 막는다.
        (Alt+F4, Alt+Tab, F11, Ctrl+W, Ctrl+Shift+Esc 등)
        """
        blocked = [
            "<Alt-F4>", "<Alt-Tab>", "<Alt-Escape>",
            "<Control-Escape>", "<Control-w>", "<Control-W>",
            "<F11>", "<Escape>",
        ]

        for seq in blocked:
            self.root.bind_all(seq, lambda e: "break")

        # 외부 사이트 단축키 (Ctrl+Shift+R / Ctrl+Shift+S)
        #
        # "<Control-Shift-R>" 처럼 글자로 바인딩하면 한글 입력 상태일 때
        # keysym이 달라져서 단축키가 먹지 않는다.
        # keycode는 한/영 상태와 무관하게 같은 값이므로 keycode로 판별한다.
        # (Windows 가상 키코드: R=82, S=83)
        self.root.bind_all("<Control-Shift-KeyPress>", self.handle_shortcut, add="+")

    def handle_shortcut(self, event):
        """Ctrl+Shift 조합 단축키 처리."""
        keycode = getattr(event, "keycode", None)

        if keycode == 82:      # R
            self.open_report_form()
            return "break"

        if keycode == 83:      # S
            self.file_share_screen()
            return "break"

        if keycode == 70:      # F — 고장 신고 상태 새로고침
            self.refresh_report_status()
            return "break"

    # -----------------------------
    # 디자인 시스템 헬퍼
    # -----------------------------
    def _setup_treeview_style(self):
        style = ttk.Style()

        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "App.Treeview",
            background=COLOR_CARD,
            fieldbackground=COLOR_CARD,
            foreground=COLOR_TEXT,
            rowheight=38,
            borderwidth=0,
            font=(self.font_family, 14)
        )
        style.configure(
            "App.Treeview.Heading",
            background=COLOR_SECONDARY_BG,
            foreground=COLOR_SECONDARY_TEXT,
            font=(self.font_family, 14, "bold"),
            borderwidth=0,
            relief="flat"
        )
        style.map(
            "App.Treeview",
            background=[("selected", COLOR_PRIMARY)],
            foreground=[("selected", "white")]
        )
        style.map(
            "App.Treeview.Heading",
            background=[("active", COLOR_SECONDARY_BG_HOVER)]
        )

    def clear(self):
        """
        화면 내용을 지운다.

        브라우저 위에 띄우는 '닫기' 버튼(Toplevel)은 창 밖에 따로 떠 있는
        것이므로 여기서 지우면 안 된다. 지우면 키오스크 브라우저를 닫을
        방법이 사라진다.
        """
        keep = (
            getattr(self, "_close_overlay", None),
            getattr(self, "_auto_delete_overlay", None),
            getattr(self, "_report_banner", None),
        )

        for widget in self.root.winfo_children():
            if any(widget is item for item in keep if item is not None):
                continue

            widget.destroy()

    def end_active_share_session(self):
        """
        열려 있던 파일 공유 세션을 닫는다.

        QR 화면의 버튼을 누르지 않고 학생증을 다시 찍어 화면을 벗어나는 경우가
        있어서, 화면이 바뀔 때마다 여기서 확실히 정리한다.
        (닫지 않으면 그 주소로 계속 파일을 보낼 수 있다)
        """
        token = getattr(self, "_share_token", None)

        if not token:
            return

        self._share_token = None

        try:
            set_upload_notifier(None)
            end_share_session(token)
        except Exception:
            pass

    def open_report_form(self):
        """
        고장 신고 창을 연다.
        닫고 나오면 잠시 뒤 상태를 한 번 더 확인해, 방금 한 신고가
        화면 안내에 바로 반영되도록 한다.
        """
        self.open_external_site(REPORT_FORM_URL)
        self.check_report_status_soon()

    def build_card(self):
        """
        화면 배경 위에 중앙 정렬된 카드 패널을 만들고, 그 카드를 반환한다.
        각 화면은 이 카드 안에 위젯을 채운다.
        """
        self.end_active_share_session()
        self.clear()

        # 화면이 바뀌면 안내의 크기와 위치도 그 화면에 맞게 다시 잡는다
        self.root.after(10, self.render_report_banner)

        # 시작 화면을 벗어나면 스캐너 자동 포커스를 끈다
        self.current_screen = "other"
        self.card_entry = None

        outer = ctk.CTkFrame(self.root, fg_color=COLOR_BG)
        outer.pack(fill="both", expand=True)
        self.outer = outer

        card = ctk.CTkFrame(
            outer,
            fg_color=COLOR_CARD,
            corner_radius=22,
            border_width=1,
            border_color=COLOR_BORDER
        )
        # 아래쪽 여백을 더 크게 준다.
        # 우측 하단 버튼이 카드 위로 겹쳐 보이지 않게 하기 위한 공간이다.
        card.pack(expand=True, padx=40, pady=(40, 92))

        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(padx=56, pady=44)

        self.build_corner_buttons()
        self.build_admin_badge()

        return content

    def build_admin_badge(self):
        """
        관리자 화면이면 왼쪽 위에 누구로 들어왔는지 띄운다. ('학번 이름 (직함)')
        화면마다 높이가 달라 카드 안에 넣으면 아래가 잘릴 수 있으므로,
        오른쪽 아래 버튼처럼 배경 위에 따로 얹는다.
        """
        if not admin_session_active():
            return

        shown = describe_admin(current_admin()) or current_admin()

        badge = ctk.CTkFrame(
            self.outer,
            fg_color=COLOR_PRIMARY_SOFT,
            corner_radius=12
        )
        badge.place(relx=0.0, rely=0.0, x=28, y=10, anchor="nw")

        ctk.CTkLabel(
            badge,
            text="관리자",
            font=(self.font_family, 13, "bold"),
            text_color="#FFFFFF",
            fg_color=COLOR_PRIMARY,
            corner_radius=8,
            width=58, height=26
        ).pack(side="left", padx=(6, 8), pady=5)

        ctk.CTkLabel(
            badge,
            text=shown,
            font=(self.font_family, 15, "bold"),
            text_color=COLOR_PRIMARY
        ).pack(side="left", padx=(0, 14), pady=5)

    def make_title(self, parent, text, subtitle=None):
        title_frame = ctk.CTkFrame(parent, fg_color="transparent")
        title_frame.pack(pady=(8, 28))

        title_label = ctk.CTkLabel(
            title_frame,
            text=text,
            font=(self.font_family, 32, "bold"),
            text_color=COLOR_TEXT
        )
        title_label.pack()

        if subtitle:
            ctk.CTkLabel(
                title_frame,
                text=subtitle,
                font=(self.font_family, 18),
                text_color=COLOR_TEXT_MUTED
            ).pack(pady=(12, 0))

        title_label.bind("<Button-1>", lambda e: self.admin_login_screen())

    def body_label(self, parent, text, size=18, muted=False, bold=False, justify="center", wraplength=0):
        kwargs = {}
        if wraplength:
            kwargs["wraplength"] = wraplength

        return ctk.CTkLabel(
            parent,
            text=text,
            font=(self.font_family, size, "bold" if bold else "normal"),
            text_color=COLOR_TEXT_MUTED if muted else COLOR_TEXT,
            justify=justify,
            **kwargs
        )

    def make_entry(self, parent, width=320, height=44, font_size=15, show=None, justify="center"):
        return ctk.CTkEntry(
            parent,
            width=width,
            height=height,
            corner_radius=10,
            justify=justify,
            font=(self.font_family, font_size),
            fg_color=COLOR_ENTRY_BG,
            border_color=COLOR_BORDER,
            border_width=1,
            text_color=COLOR_TEXT,
            show=show if show else ""
        )

    def primary_button(self, parent, text, command, width=220, height=46, font_size=16):
        return ctk.CTkButton(
            parent, text=text, command=command,
            width=width, height=height, corner_radius=12,
            font=(self.font_family, font_size, "bold"),
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
            text_color="white"
        )

    def secondary_button(self, parent, text, command, width=200, height=44, font_size=16,
                         fg_color=None, hover_color=None):
        return ctk.CTkButton(
            parent, text=text, command=command,
            width=width, height=height, corner_radius=12,
            font=(self.font_family, font_size),
            fg_color=fg_color if fg_color else COLOR_SECONDARY_BG,
            hover_color=hover_color if hover_color else COLOR_SECONDARY_BG_HOVER,
            text_color=COLOR_SECONDARY_TEXT
        )

    def danger_button(self, parent, text, command, width=150, height=44, font_size=16):
        return ctk.CTkButton(
            parent, text=text, command=command,
            width=width, height=height, corner_radius=12,
            font=(self.font_family, font_size, "bold"),
            fg_color=COLOR_DANGER, hover_color=COLOR_DANGER_HOVER,
            text_color="white"
        )

    def link_button(self, parent, text, command, width=160, height=40, font_size=15):
        return ctk.CTkButton(
            parent, text=text, command=command,
            width=width, height=height, corner_radius=10,
            font=(self.font_family, font_size),
            fg_color="transparent", hover_color=COLOR_SECONDARY_BG,
            text_color=COLOR_TEXT_MUTED
        )

    def block_close(self):
        messagebox.showwarning(
            "안내",
            "이 프로그램은 종료할 수 없습니다."
        )

    def admin_exit_program(self):
        result = messagebox.askyesno(
            "프로그램 종료",
            "정말 키오스크 프로그램을 종료하시겠습니까?"
        )

        if result:
            end_admin_session("프로그램 종료")
            self.root.destroy()

    # -----------------------------
    # 학생용 화면
    # -----------------------------
    def start_screen(self):
        # 처음 화면으로 돌아오면 인증 상태를 지운다
        # (다음 학생이 앞 사람 이름으로 파일을 보내면 안 된다)
        self._current_user = None

        # 관리자 화면에서 나온 것이면 기록을 남기고 세션을 닫는다
        end_admin_session("나감")

        card = self.build_card()
        self.make_title(card, "서인천고등학교 공용 프린터 제어 시스템")

        # 입력 영역을 한데 묶어 가운데에 둔다.
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(pady=(0, 4))

        left = ctk.CTkFrame(row, fg_color="transparent")
        left.grid(row=0, column=0, padx=(0, 0), sticky="n")

        # 이 상자를 누르면 고장 신고 상태를 바로 다시 확인한다.
        # (학생이 신고한 직후 안내가 아직 안 떴을 때 눌러보게 된다)
        badge = ctk.CTkFrame(left, fg_color=COLOR_PRIMARY_SOFT, corner_radius=14)
        badge.pack(pady=(0, 28))

        badge_label = ctk.CTkLabel(
            badge,
            text="학생증 바코드를 스캔하세요",
            font=(self.font_family, 26, "bold"),
            text_color=COLOR_PRIMARY
        )
        badge_label.pack(padx=28, pady=20)

        for widget in (badge, badge_label):
            widget.bind("<Button-1>", lambda event: self.refresh_report_status())
            widget.configure(cursor="hand2")

        self.card_entry = self.make_entry(left, width=460, height=60, font_size=28)
        self.card_entry.pack(pady=(0, 24))
        self.card_entry.bind("<Return>", lambda event: self.check_user())

        # 학생증 코드(5자리)보다 길게 입력되면 관리자코드일 수 있으므로 가린다.
        # 옆에서 보고 따라 적는 일을 막기 위해서다.
        self._scan_masked = False
        self.card_entry.bind("<KeyRelease>", lambda event: self.update_scan_mask(), add="+")

        # 학생이 입력창을 클릭하지 않아도 바코드가 바로 입력되도록
        # 이 화면에서는 항상 입력창에 포커스를 유지한다
        self.current_screen = "start"
        self.card_entry.after(50, self.focus_scan_entry)
        self.root.after(20, self.render_report_banner)

        self.primary_button(left, "확인", self.check_user, width=240, height=60, font_size=20).pack(pady=(0, 28))

        divider = ctk.CTkFrame(card, fg_color=COLOR_BORDER, height=1)
        divider.pack(fill="x", pady=(4, 24))

        self.body_label(card, "바코드 리더기로 학생증을 찍거나 학생증 코드를 직접 입력하세요.", size=16, muted=True).pack(pady=(0, 6))
        self.body_label(card, "한글 입력 상태여도 학생증 코드는 자동으로 영어 코드로 보정됩니다.", size=14, muted=True).pack()

    def build_corner_buttons(self):
        """
        화면 우측 하단에 파일 공유 / 고장 신고 버튼을 띄운다.

        컨테이너 프레임에 담으면 안 된다.
        customtkinter 의 fg_color="transparent" 는 진짜 투명이 아니라
        부모 위젯의 색을 따라가는 방식이라서, 회색 배경(outer)을 부모로 둔
        프레임이 흰색 카드 위에 겹치면 회색 사각형처럼 보인다.
        그래서 버튼을 배경 위에 하나씩 직접 배치한다.
        """
        self.secondary_button(
            self.outer, "고장 신고",
            lambda: self.open_report_form(),
            width=150, height=44, font_size=14,
            fg_color=COLOR_BG, hover_color=COLOR_SECONDARY_BG
        ).place(relx=1.0, rely=1.0, anchor="se", x=-28, y=-24)

        self.secondary_button(
            self.outer, "파일 공유",
            self.file_share_screen,
            width=150, height=44, font_size=14,
            fg_color=COLOR_BG, hover_color=COLOR_SECONDARY_BG
        ).place(relx=1.0, rely=1.0, anchor="se", x=-188, y=-24)

    def start_download_cleanup(self):
        """
        받은 파일 폴더를 주기적으로 확인해서 보관 시간이 지난 파일을 지운다.
        30초마다 한 번씩 돌며, 화면 동작을 막지 않도록 별도 스레드에서 처리한다.
        """
        def tick():
            threading.Thread(target=scan_download_folder, daemon=True).start()
            self.root.after(30_000, tick)

        self.root.after(5_000, tick)

    def start_user_csv_sync(self):
        """
        이용자 명단 CSV 를 주기적으로 확인해서, 밖에서 고친 내용을 명단에 반영한다.
        3초마다 파일의 수정시각과 크기만 확인하므로 화면이 느려지지 않는다.
        """
        # 예전 판에서 갈라진 명단 파일이 있으면 하나로 합친다
        try:
            merge_legacy_sync_csv()
        except Exception:
            pass

        if os.path.exists(get_sync_csv_path()):
            # 이미 있는 파일이면 지금 상태를 기준으로 잡는다 (시작하자마자 덮어쓰지 않도록)
            _sync_csv_state["stamp"] = _csv_stamp(get_sync_csv_path())
        else:
            write_sync_csv()

        # 관리자 명단도 처음 한 번 읽어 둔다
        try:
            sync_admin_csv(force=True)
        except Exception:
            pass

        def tick():
            # 관리자 명단 파일이 바뀌었으면 반영한다 (비밀번호 칸은 곧바로 가려진다)
            try:
                if sync_admin_csv():
                    refresh_admins = getattr(self, "_admin_list_refresh", None)
                    if refresh_admins is not None:
                        try:
                            refresh_admins()
                        except Exception:
                            self._admin_list_refresh = None
            except Exception:
                pass

            try:
                if sync_csv_changed_outside():
                    try:
                        load_sync_csv()
                    except Exception:
                        pass
                    else:
                        # 이용자 관리 화면이 떠 있으면 목록도 바로 새로 그린다
                        refresh = getattr(self, "_user_list_refresh", None)

                        if refresh is not None:
                            try:
                                refresh()
                            except Exception:
                                self._user_list_refresh = None
            except Exception:
                pass

            self.root.after(3_000, tick)

        self.root.after(3_000, tick)

    def open_external_site(self, url):
        """
        외부 사이트를 전체화면 브라우저로 연다.

        --kiosk 로 띄우면 주소창·탭·닫기 버튼이 모두 사라져서 학생이
        브라우저를 벗어나 컴퓨터를 조작할 수 없다.
        대신 닫을 방법이 없어지므로, 이 프로그램이 직접 만든 "닫기" 버튼을
        항상 위에 띄워준다. (close_overlay)

        브라우저는 진짜 크롬/엣지이므로 내려받은 파일은
        평소처럼 다운로드 폴더에 그대로 저장된다.
        """
        if not url:
            messagebox.showinfo("안내", "아직 주소가 설정되지 않았습니다. 관리자에게 문의하세요.")
            return

        # 이미 떠 있으면 무시 (연타 방지)
        if self.external_browser_open:
            return

        previous = getattr(self, "_external_process", None)
        if previous is not None and previous.poll() is None:
            return

        now = time.time()
        if now - getattr(self, "_last_external_open", 0) < 3:
            return

        self._last_external_open = now

        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ]
        edge_paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]

        process = None

        # 전용 프로필을 써야 항상 독립 프로세스로 뜬다.
        # (이미 크롬이 켜져 있으면 새 프로세스가 창만 넘기고 바로 종료된다)
        profile_dir = os.path.join(tempfile.gettempdir(), "school_printer_browser")

        # 받은 파일이 항상 지정한 폴더로 가도록 프로필 설정을 미리 심어둔다.
        # (크롬은 다운로드 폴더를 명령줄 옵션으로 받지 않는다)
        download_dir = get_download_dir()
        prepare_browser_download_dir(profile_dir, download_dir)

        for path in chrome_paths + edge_paths:
            if os.path.exists(path):
                try:
                    process = subprocess.Popen([
                        path,
                        "--kiosk",
                        url,
                        f"--user-data-dir={profile_dir}",
                        "--no-first-run",
                        "--no-default-browser-check",
                        "--disable-features=TranslateUI",
                    ])
                    break
                except Exception:
                    continue

        if process is None:
            # 크롬·엣지가 없으면 기본 브라우저로 연다 (이때는 닫기 버튼 없음)
            try:
                webbrowser.open(url)
            except Exception as e:
                self._last_external_open = 0
                messagebox.showerror("오류", f"사이트를 열 수 없습니다.\n{e}")
            return

        self._external_process = process
        self.external_browser_open = True

        # 닫기 버튼을 먼저 띄워 화면 공백을 없앤다
        self.show_close_overlay()

        # 이 프로그램이 전체화면 + 항상 위로 떠 있으면 브라우저가 뒤에 가려진다.
        # topmost 만 풀고(창은 내리지 않는다 — 내리면 작업표시줄이 드러난다),
        # 브라우저 창을 Windows API 로 직접 앞으로 끌어온다.
        if self.fullscreen:
            try:
                self.root.attributes("-topmost", False)
            except Exception:
                pass

        self.bring_browser_to_front(process.pid)

    def bring_browser_to_front(self, pid):
        """
        방금 띄운 브라우저 창을 찾아 화면 맨 앞으로 올린다.

        전체화면 프로그램이 화면을 덮고 있으면 브라우저가 뒤에서 열려
        아무것도 안 뜬 것처럼 보인다. 창이 뜨는 데 시간이 걸리므로
        잠시 기다리며 여러 번 시도한다.
        """
        def worker():
            try:
                import ctypes
                import ctypes.wintypes as wintypes
            except ImportError:
                return

            user32 = ctypes.windll.user32
            SW_SHOW = 5

            EnumWindowsProc = ctypes.WINFUNCTYPE(
                ctypes.c_bool, wintypes.HWND, wintypes.LPARAM
            )

            found = []

            def callback(hwnd, lparam):
                if not user32.IsWindowVisible(hwnd):
                    return True

                window_pid = wintypes.DWORD()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))

                if window_pid.value == pid and user32.GetWindowTextLengthW(hwnd) > 0:
                    found.append(hwnd)
                    return False

                return True

            # 최대 8초 동안 창을 찾는다
            for _ in range(40):
                time.sleep(0.2)
                found.clear()

                try:
                    user32.EnumWindows(EnumWindowsProc(callback), 0)
                except Exception:
                    return

                if found:
                    hwnd = found[0]
                    try:
                        HWND_TOPMOST = -1
                        SWP_NOMOVE = 0x0002
                        SWP_NOSIZE = 0x0001
                        SWP_SHOWWINDOW = 0x0040

                        user32.ShowWindow(hwnd, SW_SHOW)

                        # SetForegroundWindow 는 Windows 보안 정책 때문에
                        # 다른 프로그램이 포그라운드일 때 무시될 수 있다.
                        # SetWindowPos 로 최상위 지정하는 쪽이 확실하다.
                        user32.SetWindowPos(
                            hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
                        )
                        user32.BringWindowToTop(hwnd)
                        user32.SetForegroundWindow(hwnd)
                    except Exception:
                        pass
                    return

        threading.Thread(target=worker, daemon=True).start()

    def show_close_overlay(self):
        """
        전체화면 브라우저 위에 항상 떠 있는 '닫기' 버튼을 만든다.

        --kiosk 브라우저에는 X 버튼이 없으므로 이 버튼이 유일한 탈출구다.
        화면 오른쪽 위에 작게 띄우고, 브라우저가 어떤 이유로든 먼저 종료되면
        같이 사라진다.
        """
        self.close_external_overlay()   # 혹시 남아 있으면 정리

        overlay = tk.Toplevel(self.root)
        overlay.overrideredirect(True)          # 제목표시줄 없음
        overlay.attributes("-topmost", True)    # 항상 맨 위
        overlay.configure(bg=COLOR_DANGER)

        try:
            screen_w = overlay.winfo_screenwidth()
        except Exception:
            screen_w = 1920

        btn_w, btn_h = 132, 46
        margin = 18
        overlay.geometry(f"{btn_w}x{btn_h}+{screen_w - btn_w - margin}+{margin}")

        button = tk.Button(
            overlay,
            text="✕  닫기",
            command=self.close_external_site,
            bg=COLOR_DANGER,
            fg="#FFFFFF",
            activebackground=COLOR_DANGER_HOVER,
            activeforeground="#FFFFFF",
            font=(self.font_family, 14, "bold"),
            relief="flat",
            borderwidth=0,
            cursor="hand2",
        )
        button.pack(fill="both", expand=True)

        self._close_overlay = overlay

        # 브라우저가 살아 있는지 주기적으로 확인
        self._watch_external_process()

    def _watch_external_process(self):
        """브라우저가 종료되면 닫기 버튼도 없애고 원래 화면으로 돌아온다."""
        process = getattr(self, "_external_process", None)

        if process is None or process.poll() is not None:
            self.close_external_site()
            return

        # 다른 창이 닫기 버튼을 가리지 않도록 계속 위로 올린다.
        # 브라우저를 최상위(TOPMOST)로 만들었기 때문에 이 버튼도
        # 계속 최상위로 다시 올려주지 않으면 브라우저 뒤로 숨는다.
        overlay = getattr(self, "_close_overlay", None)
        if overlay is not None:
            try:
                overlay.attributes("-topmost", True)
                overlay.lift()
            except Exception:
                pass

        self._overlay_job = self.root.after(400, self._watch_external_process)

    def close_external_site(self):
        """브라우저를 종료하고 제어 프로그램 화면으로 돌아온다."""
        process = getattr(self, "_external_process", None)

        if process is not None and process.poll() is None:
            try:
                process.terminate()
                process.wait(timeout=3)
            except Exception:
                try:
                    process.kill()
                except Exception:
                    pass

        self._external_process = None
        self.close_external_overlay()
        self.close_auto_delete_overlay()
        self.restore_kiosk_focus()

    def file_share_screen(self):
        """
        파일 공유 1단계 — 학생증 인증.

        아무나 파일을 보내지 못하도록 학생증을 먼저 스캔하게 한다.
        스캔한 학생의 학번·이름으로 파일 이름이 붙어 저장된다.
        """
        error = start_upload_server()

        card = self.build_card()
        self.make_title(card, "파일 보내기")

        if error:
            self.body_label(
                card,
                "파일 공유 기능을 시작할 수 없습니다.\n관리자에게 문의하세요.",
                size=18
            ).pack(pady=(0, 14))

            self.body_label(card, error, size=14, muted=True, wraplength=620).pack(pady=(0, 26))
            self.link_button(card, "처음 화면으로", self.start_screen, width=200, height=46).pack()
            return

        # 이미 학생증을 스캔해서 들어온 상태라면 다시 찍게 하지 않는다
        current = getattr(self, "_current_user", None)
        if current:
            card_code, name, number = current
            token = create_share_session(number, name)
            self.file_share_qr_screen(token, number, name)
            return

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(pady=(0, 4))

        left = ctk.CTkFrame(row, fg_color="transparent")
        left.grid(row=0, column=0, sticky="n")

        self.body_label(
            left,
            "먼저 학생증을 스캔해 주세요.\n"
            "누가 보낸 파일인지 표시하기 위해 필요합니다.",
            size=17
        ).pack(pady=(0, 26))

        entry = self.make_entry(left, width=460, height=72, font_size=26)
        entry.pack(pady=(0, 20))
        entry.focus_set()

        def confirm():
            raw = entry.get().strip()
            code = normalize_card_code(raw)

            if not code:
                messagebox.showwarning("입력 오류", "학생증 코드를 입력하세요.")
                entry.focus_set()
                return

            user = get_student(code)

            if user is None:
                messagebox.showerror(
                    "등록되지 않음",
                    "등록되지 않은 학생증입니다.\n관리자에게 문의하세요."
                )
                entry.delete(0, tk.END)
                entry.focus_set()
                return

            card_code, name, active, unlimited, _ = user

            if active == 0:
                messagebox.showerror("사용 제한", "비활성화된 사용자입니다. 관리자에게 문의하세요.")
                entry.delete(0, tk.END)
                entry.focus_set()
                return

            number = get_student_number(card_code)

            # 파일을 보낸 뒤 바로 출력하러 갈 수 있도록 기억해 둔다
            self._current_user = (card_code, name, number)

            token = create_share_session(number, name)
            self.file_share_qr_screen(token, number, name)

        entry.bind("<Return>", lambda e: confirm())

        self.primary_button(left, "확인", confirm, width=240, height=58, font_size=19).pack(pady=(0, 20))

        self.body_label(
            card,
            "바코드 리더기로 학생증을 찍거나 학생증 코드를 직접 입력하세요.",
            size=14, muted=True
        ).pack(pady=(0, 22))

        self.link_button(card, "처음 화면으로", self.start_screen, width=200, height=44).pack()

    def file_share_qr_screen(self, token, number, name):
        """
        파일 공유 2단계 — 접속 주소와 QR 코드 안내.
        이 주소는 학생증을 스캔한 사람에게만 발급되는 전용 주소다.
        """
        card = self.build_card()
        self.make_title(card, "파일 보내기")

        url = get_file_share_url(token)
        label = f"{number} {name}".strip() if number else name

        self.body_label(card, f"{label} 님", size=19, bold=True).pack(pady=(0, 6))

        self.body_label(
            card,
            "휴대폰 또는 태블릿PC에서는 QR코드를 스캔하시고,\n노트북은 아래 주소를 입력하세요.",
            size=16
        ).pack(pady=(0, 16))

        # ── QR 코드와 주소를 나란히 ──
        # 휴대폰은 QR, 노트북은 주소 입력이라 둘 다 보여야 한다.
        share_row = ctk.CTkFrame(card, fg_color="transparent")
        share_row.pack(pady=(0, 16))

        qr_image = make_qr_image(url, size=210)

        if qr_image is not None and ImageTk is not None:
            qr_frame = ctk.CTkFrame(
                share_row, fg_color="#FFFFFF", corner_radius=14,
                border_width=1, border_color=COLOR_BORDER
            )
            qr_frame.grid(row=0, column=0, padx=(0, 22))

            photo = ImageTk.PhotoImage(qr_image)
            qr_label = tk.Label(qr_frame, image=photo, bg="#FFFFFF", borderwidth=0)
            qr_label.image = photo      # 참조를 유지해야 이미지가 사라지지 않는다
            qr_label.pack(padx=12, pady=12)

        addr_frame = ctk.CTkFrame(share_row, fg_color=COLOR_ENTRY_BG, corner_radius=16)
        addr_frame.grid(row=0, column=1)

        ctk.CTkLabel(
            addr_frame, text="브라우저 주소창에 입력",
            font=(self.font_family, 13), text_color=COLOR_TEXT_MUTED
        ).pack(padx=30, pady=(20, 8))

        ctk.CTkLabel(
            addr_frame, text=url.replace("http://", ""),
            font=(self.font_family, 23, "bold"), text_color=COLOR_PRIMARY,
            justify="center"
        ).pack(padx=30, pady=(0, 20))

        # ── 도착 알림 ──
        # 파일이 오면 이 자리에 초록색으로 표시된다
        status_frame = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=12)
        status_frame.pack(pady=(0, 16))

        # 여백은 라벨 쪽에 준다.
        # 프레임에 ipady 를 주면 글자가 위로 치우쳐 보인다.
        status_label = ctk.CTkLabel(
            status_frame,
            text="파일을 기다리는 중입니다...",
            font=(self.font_family, 16),
            text_color=COLOR_TEXT_MUTED
        )
        status_label.pack(padx=28, pady=14)

        # 지금까지 받은 파일 경로를 모아둔다 (바로 출력할 때 쓴다)
        received = {"paths": []}

        # "출력할까요?" 버튼은 파일이 처음 도착할 때 만든다
        ask_frame = ctk.CTkFrame(status_frame, fg_color="transparent")

        def on_upload(user_label, paths):
            """
            서버 스레드에서 불리므로 화면을 직접 건드리면 안 된다.
            after 로 화면 스레드에 넘겨서 처리한다.
            """
            self.root.after(0, lambda: show_received(paths))

        def show_received(paths):
            # 같은 파일이 두 번 들어가지 않게 걸러서 모은다
            for path in paths:
                if path and path not in received["paths"]:
                    received["paths"].append(path)

            count = len(received["paths"])

            try:
                if not status_label.winfo_exists():
                    return

                status_label.configure(
                    text=f"파일 {count}개를 받았습니다\n받은 파일을 지금 출력할까요?",
                    text_color=COLOR_SUCCESS,
                    font=(self.font_family, 17, "bold")
                )
                status_frame.configure(fg_color="#EAF7EF")

                if not ask_frame.winfo_ismapped():
                    ask_frame.pack(padx=28, pady=(0, 16))
            except Exception:
                pass

        def print_received():
            """받은 파일을 바로 출력 확인 화면으로 넘긴다."""
            paths = [p for p in received["paths"] if p and os.path.exists(p)]

            if not paths:
                messagebox.showwarning(
                    "안내",
                    "출력할 파일을 찾을 수 없습니다.\n보관 시간이 지나 삭제되었을 수 있습니다."
                )
                return

            current = getattr(self, "_current_user", None)

            if not current:
                messagebox.showwarning("안내", "학생 정보를 확인할 수 없습니다.")
                return

            card_code = current[0]
            user = get_student(card_code)

            if user is None:
                messagebox.showwarning("안내", "등록 정보를 확인할 수 없습니다.")
                return

            _, user_name, active, unlimited, _ = user

            if active == 0:
                messagebox.showerror("사용 제한", "비활성화된 사용자입니다.")
                return

            cleanup()

            # 한 번에 고를 수 있는 개수를 넘으면 출력 화면에서 막히므로 미리 알린다
            if not unlimited:
                max_files = get_max_files_per_job()
                if max_files > UNLIMITED and len(paths) > max_files:
                    messagebox.showwarning(
                        "파일 개수 제한",
                        f"한 번에 최대 {max_files}개까지만 출력할 수 있습니다.\n"
                        f"받은 파일: {len(paths)}개\n\n"
                        "파일 선택 화면에서 나눠서 출력해 주세요."
                    )
                    self.file_screen(card_code, user_name, unlimited)
                    return

            self.print_confirm_screen(card_code, user_name, unlimited, paths)

        def skip_print():
            """지금은 출력하지 않고 계속 파일을 받는다."""
            try:
                ask_frame.pack_forget()
                status_label.configure(
                    text=f"파일 {len(received['paths'])}개를 받았습니다",
                    font=(self.font_family, 17, "bold")
                )
            except Exception:
                pass

        # 초록 알림 상자 안이라 버튼도 초록 계열로 맞춘다
        ctk.CTkButton(
            ask_frame, text="예, 출력합니다", command=print_received,
            width=170, height=46, corner_radius=12,
            font=(self.font_family, 15, "bold"),
            fg_color=COLOR_SUCCESS, hover_color="#1A7A45",
            text_color="white"
        ).grid(row=0, column=0, padx=6)

        self.secondary_button(
            ask_frame, "아니요", skip_print,
            width=120, height=46, font_size=15
        ).grid(row=0, column=1, padx=6)

        set_upload_notifier(on_upload)

        minutes = get_auto_delete_minutes()
        notice = "학교 와이파이에 연결되어 있어야 합니다."
        notice += "\n파일이 도착하면 바로 출력할 수 있습니다."
        if minutes > 0:
            notice += f"\n보낸 파일은 {minutes}분 후 자동으로 삭제됩니다."

        self.body_label(card, notice, size=14, muted=True).pack(pady=(0, 22))

        # 화면이 바뀔 때 자동으로 정리되도록 지금 세션을 기억해 둔다
        self._share_token = token

        def cleanup():
            self._share_token = None
            set_upload_notifier(None)   # 화면을 벗어나면 알림을 끊는다
            end_share_session(token)

        def go_print():
            """보낸 파일을 바로 출력하러 간다 (다시 스캔할 필요 없이)."""
            cleanup()

            current = getattr(self, "_current_user", None)
            if not current:
                self.start_screen()
                return

            card_code, user_name, _ = current
            user = get_student(card_code)

            if user is None:
                self.start_screen()
                return

            _, name2, active, unlimited, _ = user

            if active == 0:
                self.start_screen()
                return

            self.file_screen(card_code, name2, unlimited)

        def go_home():
            cleanup()
            self.start_screen()

        self.primary_button(
            card, "파일을 보냈어요 — 출력하러 가기",
            go_print, width=340, height=56, font_size=18
        ).pack(pady=(0, 12))

        self.link_button(card, "처음 화면으로", go_home, width=200, height=44).pack()

    def show_auto_delete_overlay(self):
        """
        파일 공유 화면 우측 하단에 자동 삭제 안내를 띄운다.

        닫기 버튼과 같은 방식(제목표시줄 없는 항상-위 Toplevel)으로 만들되,
        위치만 화면 오른쪽 아래로 둔다. 설정된 보관 시간이 바뀌어도
        안내 문구에 실제 값이 반영되도록 매번 새로 읽어온다.
        """
        self.close_auto_delete_overlay()   # 혹시 남아 있으면 정리

        minutes = get_auto_delete_minutes()

        # 자동 삭제를 꺼둔 경우(0분)에는 안내할 내용이 없으므로 띄우지 않는다
        if minutes <= 0:
            return

        overlay = tk.Toplevel(self.root)
        overlay.overrideredirect(True)
        overlay.attributes("-topmost", True)
        overlay.configure(bg=COLOR_TEXT)

        text = f"받은 파일은 {minutes}분 후 자동으로 삭제됩니다"

        label = tk.Label(
            overlay,
            text=text,
            bg=COLOR_TEXT,
            fg="#FFFFFF",
            font=(self.font_family, 14, "bold"),
            padx=22,
            pady=12,
        )
        label.pack()

        overlay.update_idletasks()

        try:
            screen_w = overlay.winfo_screenwidth()
            screen_h = overlay.winfo_screenheight()
        except Exception:
            screen_w, screen_h = 1920, 1080

        box_w = overlay.winfo_reqwidth()
        box_h = overlay.winfo_reqheight()
        margin = 24

        x = screen_w - box_w - margin
        y = screen_h - box_h - margin
        overlay.geometry(f"{box_w}x{box_h}+{x}+{y}")

        self._auto_delete_overlay = overlay

        # 닫기 버튼과 마찬가지로 브라우저가 최상위이므로 계속 위로 올려야 한다
        self._watch_auto_delete_overlay()

    def _watch_auto_delete_overlay(self):
        """브라우저가 떠 있는 동안 안내 창이 가려지지 않게 계속 위로 올린다."""
        if not self.external_browser_open:
            self.close_auto_delete_overlay()
            return

        overlay = getattr(self, "_auto_delete_overlay", None)
        if overlay is not None:
            try:
                overlay.attributes("-topmost", True)
                overlay.lift()
            except Exception:
                pass

        self._auto_delete_overlay_job = self.root.after(400, self._watch_auto_delete_overlay)

    def close_auto_delete_overlay(self):
        """자동 삭제 안내 창과 감시 작업을 정리한다."""
        job = getattr(self, "_auto_delete_overlay_job", None)
        if job is not None:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
            self._auto_delete_overlay_job = None

        overlay = getattr(self, "_auto_delete_overlay", None)
        if overlay is not None:
            try:
                overlay.destroy()
            except Exception:
                pass
            self._auto_delete_overlay = None

    def close_external_overlay(self):
        """닫기 버튼 창과 감시 작업을 정리한다."""
        job = getattr(self, "_overlay_job", None)
        if job is not None:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
            self._overlay_job = None

        overlay = getattr(self, "_close_overlay", None)
        if overlay is not None:
            try:
                overlay.destroy()
            except Exception:
                pass
            self._close_overlay = None

    def restore_kiosk_focus(self):
        self.external_browser_open = False

        if not self.fullscreen:
            return

        self.root.attributes("-topmost", True)
        self.root.lift()
        self.root.focus_force()

    # 이 글자 수를 넘으면 첫 화면 입력을 * 로 가린다 (학생증 코드는 5자리)
    SCAN_MASK_AFTER = 5

    def update_scan_mask(self):
        """입력 길이에 따라 첫 화면 입력칸을 가리거나 다시 보이게 한다."""
        entry = getattr(self, "card_entry", None)

        if entry is None:
            return

        try:
            if not entry.winfo_exists():
                return

            hide = len(entry.get()) > self.SCAN_MASK_AFTER

            # 같은 상태로 계속 다시 설정하지 않도록 지금 상태를 기억해 둔다
            if getattr(self, "_scan_masked", False) != hide:
                entry.configure(show="*" if hide else "")
                self._scan_masked = hide
        except Exception:
            pass

    def reset_scan_entry(self):
        """입력창을 비우고 다시 포커스를 준다 (다음 스캔 대기 상태)."""
        if self.card_entry is None or not self.card_entry.winfo_exists():
            return

        try:
            self.card_entry.delete(0, tk.END)
        except Exception:
            pass

        self.update_scan_mask()
        self.focus_scan_entry()

    def check_user(self):
        raw_code = self.card_entry.get().strip()
        card_code = normalize_card_code(raw_code)

        if not card_code:
            messagebox.showwarning("입력 오류", "학생증 코드를 입력하세요.")
            return

        # 관리자코드 감지 (관리자 카드의 QR 을 찍은 경우)
        admin_name = self.match_admin_code(card_code)
        if admin_name:
            self.reset_scan_entry()
            begin_admin_session(admin_name, "관리자코드")
            self.admin_menu()
            return

        user = get_student(card_code)

        if not user:
            messagebox.showerror(
                "인증 실패",
                f"등록되지 않은 학생증 코드입니다.\n\n입력 코드: {card_code}"
            )
            self.reset_scan_entry()
            return

        card_code, name, active, unlimited, _ = user

        if active == 0:
            messagebox.showerror("사용 제한", "비활성화된 사용자입니다. 관리자에게 문의하세요.")
            self.reset_scan_entry()
            return

        # 제한값이 0 이면 제한을 두지 않는다
        if not unlimited and not is_unlimited_limit():
            today_count = get_today_count(card_code)
            daily_limit = get_today_limit_value()

            if today_count >= daily_limit:
                unit = "장" if get_limit_mode() == "pages" else "회"
                messagebox.showerror(
                    "출력 제한",
                    f"{name} 사용자는 오늘 출력 가능 {'용지 장수' if get_limit_mode()=='pages' else '횟수'}를 초과했습니다.\n"
                    f"오늘 사용량: {today_count}/{daily_limit}{unit}"
                )
                self.start_screen()
                return

        # 파일 공유에서 다시 스캔하지 않아도 되도록 기억해 둔다
        self._current_user = (card_code, name, get_student_number(card_code))

        self.file_screen(card_code, name, unlimited)

    def file_screen(self, card_code, name, unlimited):
        card = self.build_card()
        self.make_title(card, "파일 선택 화면")

        limit_mode = get_limit_mode()
        today_count = get_today_count(card_code)
        daily_limit = get_today_limit_value()
        remain = daily_limit - today_count
        max_files = get_max_files_per_job()

        unit = "장" if limit_mode == "pages" else "회"
        limit_label = "오늘 사용한 용지" if limit_mode == "pages" else "오늘 출력 횟수"
        remain_label = "오늘 남은 용지" if limit_mode == "pages" else "남은 출력 가능 횟수"

        if unlimited:
            limit_text = "무제한 출력 사용자"
            remain_text = "제한 없음"
        elif daily_limit <= UNLIMITED:
            # 제한값이 0 이면 제한을 두지 않는다
            limit_text = f"{limit_label}: {today_count}{unit}"
            remain_text = "제한 없음"
        else:
            limit_text = f"{limit_label}: {today_count}/{daily_limit}{unit}"
            remain_text = f"{remain}{unit}"

        if unlimited or max_files <= UNLIMITED:
            max_text = "제한 없음"
        else:
            max_text = f"최대 {max_files}개"

        info_card = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=16)
        info_card.pack(pady=(0, 28), padx=6, ipadx=24, ipady=22)

        student_number = get_student_number(card_code)

        info_rows = [("학생증 코드", card_code)]

        if student_number:
            info_rows.append(("학번", student_number))

        info_rows += [
            ("이름", name),
            (limit_label, limit_text),
            (remain_label, remain_text),
            ("한 번에 선택 가능 파일 수", max_text),
        ]

        # 위와 같은 이유로 ipady 여백이 아래쪽에만 몰린다.
        # 행들을 감싸는 틀을 하나 두고 expand=True 를 주면
        # 남는 공간이 위아래로 나뉘어 내용이 상자 한가운데 온다.
        rows_wrap = ctk.CTkFrame(info_card, fg_color="transparent")
        rows_wrap.pack(fill="x", expand=True)

        for label_text, value_text in info_rows:
            row = ctk.CTkFrame(rows_wrap, fg_color="transparent")
            row.pack(fill="x", pady=4, padx=10)
            self.body_label(row, label_text, size=16, muted=True).pack(side="left")
            ctk.CTkLabel(
                row, text=value_text,
                font=(self.font_family, 16, "bold"),
                text_color=COLOR_TEXT
            ).pack(side="right")

        self.primary_button(
            card, "문서 파일 선택",
            lambda: self.select_files(card_code, name, unlimited),
            width=320, height=60, font_size=20
        ).pack(pady=(0, 14))

        self.body_label(
            card,
            "지원 형식: PDF, DOC, DOCX, HWP, HWPX, JPG, PNG, PPT, PPTX, CSV, XLS, XLSX, TXT",
            size=14, muted=True
        ).pack(pady=(0, 24))

        self.link_button(card, "처음 화면으로", self.start_screen, width=180, height=42).pack()

    def select_files(self, card_code, name, unlimited):
        initial_dir = get_setting("file_open_dir", "")
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = os.path.expanduser("~\\Desktop")

        file_paths = filedialog.askopenfilenames(
            title="출력할 파일 선택",
            initialdir=initial_dir,
            filetypes=[
                ("지원 파일", "*.pdf *.doc *.docx *.hwp *.hwpx *.jpg *.jpeg *.png *.ppt *.pptx *.csv *.xls *.xlsx *.txt"),
                ("PDF 파일", "*.pdf"),
                ("Word 파일", "*.doc *.docx"),
                ("한컴 파일", "*.hwp *.hwpx"),
                ("이미지 파일", "*.jpg *.jpeg *.png"),
                ("PowerPoint 파일", "*.ppt *.pptx"),
                ("Excel/CSV 파일", "*.xls *.xlsx *.csv"),
                ("TXT 파일", "*.txt"),
                ("모든 파일", "*.*")
            ]
        )

        if not file_paths:
            return

        file_paths = list(file_paths)

        unsupported = [p for p in file_paths if not is_supported_file(p)]
        if unsupported:
            messagebox.showerror("파일 형식 오류", "지원하지 않는 파일이 포함되어 있습니다.")
            return

        if not unlimited:
            max_files = get_max_files_per_job()

            # 0 이면 한 번에 고를 수 있는 개수에 제한을 두지 않는다
            if max_files > UNLIMITED and len(file_paths) > max_files:
                messagebox.showerror(
                    "파일 개수 제한",
                    f"한 번에 최대 {max_files}개까지만 선택할 수 있습니다."
                )
                return

            # "파일 개수" 모드는 여기서 바로 확인할 수 있다.
            # "용지 장수" 모드는 실제 페이지 수를 세어봐야 알 수 있으므로
            # (워드/한글 파일은 변환까지 해봐야 한다) 다음 확인 화면에서 검사한다.
            if get_limit_mode() == "files" and not is_unlimited_limit():
                today_count = get_today_count(card_code)
                remain = get_today_limit_value() - today_count

                if len(file_paths) > remain:
                    messagebox.showerror(
                        "출력 횟수 제한",
                        f"오늘 남은 출력 가능 횟수는 {remain}회입니다.\n"
                        f"선택한 파일 수: {len(file_paths)}개"
                    )
                    return

        self.print_confirm_screen(card_code, name, unlimited, file_paths)

    def print_confirm_screen(self, card_code, name, unlimited, file_paths):
        card = self.build_card()
        self.make_title(card, "출력 정보 확인")

        printer_name = get_setting("selected_printer", "")
        limit_mode = get_limit_mode()

        # 페이지 수는 모드와 관계없이 항상 세어서 보여준다.
        # 학생이 몇 장이 나가는지 미리 알 수 있어야 종이 낭비를 줄일 수 있다.
        #
        # 워드/한글 파일은 LibreOffice 로 PDF 변환을 해야 페이지 수를 알 수 있는데
        # 이 작업이 수 초씩 걸린다. 화면 스레드에서 그대로 돌리면 그동안 창이 멈춰
        # Windows 가 "응답 없음"으로 표시하므로, 반드시 별도 스레드에서 처리하고
        # 결과만 화면으로 돌려받는다.
        page_counts = {}
        unknown_files = []
        selected_pages = {}   # {파일경로: {인쇄할 페이지 번호}} — 비어 있으면 전체 출력
        file_copies = {}      # {파일경로: 부수} — 없으면 1부

        progress_label = self.body_label(
            card, "파일을 확인하는 중입니다...", size=17
        )
        progress_label.pack(pady=(40, 8))

        detail_label = self.body_label(card, "", size=14, muted=True)
        detail_label.pack(pady=(0, 40))

        counting_done = {"value": False}

        def count_worker():
            converted_cache = getattr(self, "_print_converted_cache", {})
            total = len(file_paths)

            for i, path in enumerate(file_paths, start=1):
                file_name = os.path.basename(path)

                # 진행 상황 갱신은 화면 스레드에서 해야 안전하다
                self.root.after(
                    0,
                    lambda n=file_name, i=i, t=total: update_progress(n, i, t)
                )

                try:
                    pages = count_pages_for_print(path, converted_cache)
                except Exception:
                    pages = None

                if pages is None:
                    unknown_files.append(file_name)
                    pages = 1  # 확인할 수 없으면 최소 1장으로 어림잡는다

                page_counts[path] = pages

            self._print_converted_cache = converted_cache
            self.root.after(0, finish_counting)

        def update_progress(file_name, index, total):
            if counting_done["value"]:
                return
            try:
                progress_label.configure(text=f"파일을 확인하는 중입니다... ({index}/{total})")
                detail_label.configure(text=file_name)
            except Exception:
                pass

        def finish_counting():
            counting_done["value"] = True

            # 계산이 끝나기 전에 사용자가 다른 화면으로 넘어갔을 수 있다.
            # 그 경우 화면 요소가 이미 사라졌으므로 아무것도 하지 않는다.
            try:
                if not card.winfo_exists():
                    return
            except Exception:
                return

            try:
                progress_label.destroy()
                detail_label.destroy()
            except Exception:
                pass

            # tkinter 의 after 콜백에서 난 예외는 조용히 묻혀서
            # 화면이 중간까지만 그려진 채 멈춘 것처럼 보인다.
            # 원인을 알 수 있도록 여기서 붙잡아 보여준다.
            try:
                build_confirm_ui()
            except Exception as e:
                import traceback
                traceback.print_exc()
                messagebox.showerror(
                    "화면 오류",
                    f"출력 정보를 표시하는 중 문제가 생겼습니다.\n\n{type(e).__name__}: {e}"
                )
                self.file_screen(card_code, name, unlimited)

        self._confirm_page_counts = page_counts

        def build_confirm_ui():
            student_number = get_student_number(card_code)
            user_text = f"{student_number} {name} ({card_code})" if student_number else f"{name} ({card_code})"

            info_card = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=16)
            info_card.pack(pady=(0, 12), padx=6, ipadx=24, ipady=18, fill="x")

            info_label = self.body_label(info_card, "", size=16, justify="left")

            # pack 의 ipady 로 늘어난 여백은 기본적으로 아래쪽에만 남는다.
            # (자식 위젯이 위에서부터 채워지고 남는 공간이 밑에 몰리기 때문)
            # expand=True 를 주면 남는 공간을 위아래로 나눠 가져서 글자가
            # 상자 한가운데 오게 된다. anchor="w" 이므로 좌우는 왼쪽 정렬 그대로다.
            info_label.pack(anchor="w", padx=8, expand=True)

            # ── 양면/단면 선택 ──
            # 프린터가 자동 양면 유닛을 지원해야 실제로 양면 인쇄가 된다.
            duplex_var = tk.IntVar(value=0)

            duplex_frame = ctk.CTkFrame(card, fg_color="transparent")
            duplex_frame.pack(pady=(0, 12))

            self.body_label(duplex_frame, "인쇄 방식", size=15, bold=True).grid(row=0, column=0, padx=(0, 14))

            ctk.CTkRadioButton(
                duplex_frame, text="단면 인쇄", variable=duplex_var, value=0,
                font=(self.font_family, 15), text_color=COLOR_TEXT,
                fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
                command=lambda: refresh_summary()
            ).grid(row=0, column=1, padx=8)

            ctk.CTkRadioButton(
                duplex_frame, text="양면 인쇄", variable=duplex_var, value=1,
                font=(self.font_family, 15), text_color=COLOR_TEXT,
                fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
                command=lambda: refresh_summary()
            ).grid(row=0, column=2, padx=8)

            max_copies = get_max_copies()

            def effective_pages(path):
                """이 파일에서 실제로 인쇄될 페이지 수 (페이지 선택을 반영)."""
                chosen = selected_pages.get(path)
                if chosen:
                    return len(chosen)
                return page_counts.get(path, 1)

            def copies_of(path):
                """이 파일을 몇 부 뽑을지. 지정하지 않았으면 1부."""
                return max(1, file_copies.get(path, 1))

            def compute_summary():
                duplex = bool(duplex_var.get())

                # 실제로 소모될 용지 장수. 파일마다 부수가 다를 수 있다.
                total_sheets = sum(
                    sheets_from_pages(effective_pages(p), duplex) * copies_of(p)
                    for p in file_paths
                )

                # 제한에 차감되는 양은 모드에 따라 다르다.
                # 파일 개수 모드에서도 부수만큼 출력 건수가 늘어난다.
                job_amount = (
                    total_sheets if limit_mode == "pages"
                    else sum(copies_of(p) for p in file_paths)
                )

                today_count = get_today_count(card_code)
                daily_limit = get_today_limit_value()
                after_count = today_count + job_amount

                return job_amount, after_count, daily_limit, duplex, total_sheets

            def refresh_summary():
                job_amount, after_count, daily_limit, duplex, total_sheets = compute_summary()
                total_pages = sum(effective_pages(p) for p in file_paths)
                duplex_text = "양면" if duplex else "단면"

                total_copies = sum(copies_of(p) for p in file_paths)
                copies_text = f"  ·  총 {total_copies}부" if total_copies > len(file_paths) else ""

                lines = [
                    f"사용자: {user_text}",
                    f"선택된 프린터: {printer_name}",
                    f"선택한 파일 수: {len(file_paths)}개{copies_text}",
                    f"전체 페이지: {total_pages}쪽  →  사용 용지: {total_sheets}장 ({duplex_text})",
                ]

                if unlimited:
                    lines.append("무제한 출력 사용자이므로 출력 제한 없음")
                elif daily_limit <= UNLIMITED:
                    # 제한값이 0 이면 제한을 두지 않는다
                    unit = "장" if limit_mode == "pages" else "회"
                    label = "사용 용지" if limit_mode == "pages" else "출력 횟수"
                    lines.append(f"출력 후 오늘 {label}: {after_count}{unit} (제한 없음)")
                elif limit_mode == "pages":
                    lines.append(f"출력 후 오늘 사용 용지: {after_count}/{daily_limit}장")
                else:
                    lines.append(f"출력 후 오늘 출력 횟수: {after_count}/{daily_limit}회")

                if unknown_files:
                    lines.append(
                        f"※ {', '.join(unknown_files)} 은(는) 페이지 수를 확인할 수 없어 1장으로 계산했습니다."
                    )

                info_label.configure(text="\n".join(lines))

            refresh_summary()

            columns = ("no", "file_name", "file_type", "pages", "copies")
            tree = ttk.Treeview(card, columns=columns, show="headings", height=8, style="App.Treeview")
            tree.heading("no", text="번호")
            tree.heading("file_name", text="파일명")
            tree.heading("file_type", text="형식")
            tree.heading("pages", text="페이지 수")
            tree.heading("copies", text="부수")

            tree.column("no", width=50, anchor="center")
            tree.column("file_name", width=400)
            tree.column("file_type", width=80, anchor="center")
            tree.column("pages", width=130, anchor="center")
            tree.column("copies", width=100, anchor="center")

            tree.pack(pady=(0, 8))

            def refresh_tree(keep_selection=True):
                selected_rows = tree.selection() if keep_selection else ()
                selected_idx = None
                if selected_rows:
                    try:
                        selected_idx = int(tree.item(selected_rows[0], "values")[0]) - 1
                    except (ValueError, IndexError):
                        selected_idx = None

                for row in tree.get_children():
                    tree.delete(row)

                items = []
                for idx, path in enumerate(file_paths, start=1):
                    file_name = os.path.basename(path)
                    ext = os.path.splitext(path)[1].lower()

                    total = page_counts.get(path, 1)
                    chosen = selected_pages.get(path)

                    if file_name in unknown_files:
                        pages_text = "확인 불가"
                    elif chosen:
                        pages_text = f"{len(chosen)}쪽 / {total}쪽"
                    else:
                        pages_text = f"{total}쪽"

                    items.append(
                        tree.insert("", "end",
                                    values=(idx, file_name, ext, pages_text, f"{copies_of(path)}부"))
                    )

                # 목록을 다시 그려도 고르고 있던 줄은 그대로 두어야
                # 부수를 연달아 조절할 때 선택이 풀리지 않는다
                if selected_idx is not None and 0 <= selected_idx < len(items):
                    tree.selection_set(items[selected_idx])

            refresh_tree(keep_selection=False)

            # 첫 줄을 미리 골라둬서 바로 부수를 조절할 수 있게 한다
            first_rows = tree.get_children()
            if first_rows:
                tree.selection_set(first_rows[0])

            def open_page_selector(index=None):
                """
                미리보기와 페이지 선택을 겸하는 창을 연다.
                모든 페이지를 썸네일로 훑어보면서 인쇄할 쪽을 고를 수 있다.
                """
                if index is None:
                    sel = tree.selection()
                    if not sel:
                        messagebox.showwarning("선택 오류", "먼저 목록에서 파일을 선택하세요.")
                        return
                    index = int(tree.item(sel[0], "values")[0]) - 1

                if index < 0 or index >= len(file_paths):
                    return

                path = file_paths[index]
                file_name = os.path.basename(path)

                if file_name in unknown_files:
                    messagebox.showinfo(
                        "안내",
                        "이 파일은 내용을 확인할 수 없어 미리보기와 페이지 선택을 지원하지 않습니다.\n"
                        "출력은 정상적으로 진행됩니다."
                    )
                    return

                # 미리보기·페이지 선택은 변환된 PDF 기준으로 동작한다
                cache = getattr(self, "_print_converted_cache", {})
                pdf_path = cache.get(path, path)

                if not str(pdf_path).lower().endswith(".pdf"):
                    messagebox.showinfo("안내", "이 파일은 미리보기를 지원하지 않습니다.")
                    return

                total = page_counts.get(path, 1)

                def on_confirm(pages):
                    if len(pages) == total:
                        selected_pages.pop(path, None)   # 전체 선택은 저장하지 않는다
                    else:
                        selected_pages[path] = pages
                    refresh_tree()
                    refresh_summary()

                self.show_page_selector(
                    pdf_path, file_name, total, selected_pages.get(path), on_confirm
                )

            def open_selected_from_tree(event=None):
                sel = tree.selection()
                if not sel:
                    return
                open_page_selector(int(tree.item(sel[0], "values")[0]) - 1)

            tree.bind("<Double-1>", open_selected_from_tree)

            # ── 파일별 부수 조절 ──
            # 목록에서 파일을 고른 뒤 부수를 정한다. 파일마다 다르게 지정할 수 있다.
            copies_frame = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=12)
            copies_frame.pack(pady=(0, 10))

            self.body_label(
                copies_frame, "선택한 파일 부수", size=15, bold=True
            ).grid(row=0, column=0, padx=(20, 12), pady=12)

            def current_path():
                """목록에서 지금 고른 파일의 경로."""
                rows = tree.selection()
                if not rows:
                    return None
                try:
                    idx = int(tree.item(rows[0], "values")[0]) - 1
                except (ValueError, IndexError):
                    return None
                return file_paths[idx] if 0 <= idx < len(file_paths) else None

            def apply_copies(value, silent=False):
                """고른 파일의 부수를 바꾼다. 범위를 벗어나면 되돌린다."""
                path = current_path()

                if path is None:
                    if not silent:
                        messagebox.showinfo("안내", "먼저 목록에서 파일을 선택하세요.")
                    sync_copies_entry()
                    return

                if value < 1:
                    value = 1

                if max_copies > UNLIMITED and value > max_copies:
                    if not silent:
                        messagebox.showinfo("안내", f"한 파일당 최대 {max_copies}부까지 뽑을 수 있습니다.")
                    value = max_copies

                file_copies[path] = value
                sync_copies_entry()
                refresh_tree()
                refresh_summary()

            def sync_copies_entry():
                """입력칸을 지금 고른 파일의 부수로 맞춘다."""
                path = current_path()
                value = copies_of(path) if path else 1
                copies_entry.delete(0, tk.END)
                copies_entry.insert(0, str(value))

            def change_copies(delta):
                path = current_path()
                if path is None:
                    messagebox.showinfo("안내", "먼저 목록에서 파일을 선택하세요.")
                    return
                apply_copies(copies_of(path) + delta)

            def commit_entry(event=None):
                """직접 입력한 값을 적용한다."""
                text = copies_entry.get().strip()

                if not text.isdigit() or int(text) < 1:
                    messagebox.showwarning("입력 오류", "부수는 1 이상의 숫자여야 합니다.")
                    sync_copies_entry()
                    return

                apply_copies(int(text))

            self.secondary_button(
                copies_frame, "−", lambda: change_copies(-1),
                width=46, height=42, font_size=18
            ).grid(row=0, column=1, padx=3, pady=12)

            copies_entry = self.make_entry(copies_frame, width=76, height=42, font_size=16)
            copies_entry.grid(row=0, column=2, padx=3, pady=12)
            copies_entry.insert(0, "1")
            copies_entry.bind("<Return>", commit_entry)
            copies_entry.bind("<FocusOut>", lambda e: commit_entry())

            self.secondary_button(
                copies_frame, "+", lambda: change_copies(1),
                width=46, height=42, font_size=18
            ).grid(row=0, column=3, padx=3, pady=12)

            def apply_all():
                """지금 입력한 부수를 모든 파일에 똑같이 적용한다."""
                text = copies_entry.get().strip()

                if not text.isdigit() or int(text) < 1:
                    messagebox.showwarning("입력 오류", "부수는 1 이상의 숫자여야 합니다.")
                    sync_copies_entry()
                    return

                value = int(text)
                if max_copies > UNLIMITED and value > max_copies:
                    messagebox.showinfo("안내", f"한 파일당 최대 {max_copies}부까지 뽑을 수 있습니다.")
                    value = max_copies

                for path in file_paths:
                    file_copies[path] = value

                sync_copies_entry()
                refresh_tree()
                refresh_summary()

            self.secondary_button(
                copies_frame, "모두 같게", apply_all,
                width=110, height=42, font_size=14
            ).grid(row=0, column=4, padx=(12, 20), pady=12)

            # 다른 파일을 고르면 입력칸도 그 파일 부수로 바뀐다
            tree.bind("<<TreeviewSelect>>", lambda e: sync_copies_entry())
            sync_copies_entry()

            self.body_label(
                card,
                "파일을 선택한 뒤 '미리보기 · 페이지 선택'을 누르거나 파일을 더블클릭하세요.",
                size=14, muted=True
            ).pack(pady=(0, 10))

            def start_print():
                job_amount, after_count, daily_limit, duplex, total_sheets = compute_summary()

                # 페이지 수를 확인하지 못한 파일은 1장으로 어림잡았을 뿐이라
                # 실제로는 훨씬 많은 종이가 나갈 수 있다. 그대로 진행하지 않고 한 번 묻는다.
                if unknown_files:
                    proceed = messagebox.askyesno(
                        "확인",
                        f"{', '.join(unknown_files)} 은(는) 내용을 확인할 수 없어\n"
                        "몇 장이 나올지 알 수 없습니다.\n\n"
                        "생각보다 많은 종이가 나갈 수 있고, 변환에 실패해\n"
                        "인쇄되지 않을 수도 있습니다.\n\n"
                        "그래도 출력할까요?"
                    )

                    if not proceed:
                        return

                if not unlimited and daily_limit > UNLIMITED and after_count > daily_limit:
                    unit_label = "용지 장수" if limit_mode == "pages" else "출력 횟수"
                    unit = "장" if limit_mode == "pages" else "회"
                    remain = max(daily_limit - (after_count - job_amount), 0)

                    messagebox.showerror(
                        f"{unit_label} 제한",
                        f"오늘 남은 {unit_label}는 {remain}{unit}입니다.\n"
                        f"이번 출력에 필요한 양: {job_amount}{unit}"
                    )
                    return

                # 페이지 수는 항상 세므로 기록에도 항상 실제 장수를 남긴다.
                # (나중에 모드를 바꿔도 과거 기록의 용지 사용량을 그대로 볼 수 있다)
                # 인쇄에 넘길 페이지 수는 실제로 고른 만큼이어야 한다
                actual_counts = {p: effective_pages(p) for p in file_paths}

                self.do_print_files(
                    card_code, name, file_paths,
                    duplex=duplex,
                    page_counts=actual_counts,
                    converted_cache=getattr(self, "_print_converted_cache", None),
                    selected_pages=dict(selected_pages),
                    file_copies={p: copies_of(p) for p in file_paths}
                )

            btn_frame = ctk.CTkFrame(card, fg_color="transparent") 
            btn_frame.pack(pady=(0, 8))

            self.secondary_button(
                btn_frame, "미리보기 · 페이지 선택",
                open_page_selector,
                width=230, height=54, font_size=15
            ).grid(row=0, column=0, padx=8)

            self.primary_button(
                btn_frame, "출력 시작",
                start_print,
                width=230, height=54, font_size=16
            ).grid(row=0, column=1, padx=8)

            self.secondary_button(
                btn_frame, "취소", 
                lambda: self.file_screen(card_code, name, unlimited),
                width=230, height=54, font_size=16
            ).grid(row=0, column=2, padx=8)


        threading.Thread(target=count_worker, daemon=True).start()

    def show_page_selector(self, file_path, file_name, total_pages, current_pages, on_confirm):
        """
        미리보기와 인쇄 페이지 선택을 함께 하는 창.

        모든 페이지를 썸네일로 보여주고, 각 페이지의 체크박스나 직접 입력으로
        인쇄할 쪽을 고른다. 썸네일을 클릭하면 크게 볼 수 있다.

        페이지가 많은 문서는 썸네일을 한꺼번에 만들면 시간이 오래 걸리므로,
        화면에 실제로 보이는 부분만 그때그때 렌더링한다(스크롤 시 추가 렌더링).
        렌더링은 백그라운드 스레드에서 하고 결과만 화면에 반영해 창이 멈추지 않게 한다.
        """
        win = ctk.CTkToplevel(self.root)
        win.title(f"미리보기 · 페이지 선택 - {file_name}")
        win.after(250, lambda: apply_window_icon(win))
        win.geometry("1080x760")
        win.configure(fg_color=COLOR_BG)
        win.transient(self.root)
        win.attributes("-topmost", True)
        win.lift()
        win.after(120, win.grab_set)

        selected = set(current_pages) if current_pages else set(range(1, total_pages + 1))
        check_vars = {}
        thumb_labels = {}      # {페이지번호: 이미지 라벨}
        rendered = set()       # 이미 썸네일을 만든 페이지
        requested = set()      # 렌더링을 요청해둔 페이지 (중복 요청 방지)
        photo_refs = {}        # 이미지 객체 유지 (없으면 가비지 컬렉션에 지워진다)
        closed = {"value": False}

        # ── 상단: 제목과 일괄 선택 ──
        ctk.CTkLabel(
            win, text=file_name,
            font=(self.font_family, 19, "bold"), text_color=COLOR_TEXT
        ).pack(pady=(22, 4))

        ctk.CTkLabel(
            win, text=f"전체 {total_pages}쪽",
            font=(self.font_family, 13), text_color=COLOR_TEXT_MUTED
        ).pack(pady=(0, 14))

        top_frame = ctk.CTkFrame(win, fg_color="transparent")
        top_frame.pack(pady=(0, 10))

        def set_selection(pages):
            selected.clear()
            selected.update(pages)
            for num, var in check_vars.items():
                var.set(1 if num in selected else 0)
            range_entry.delete(0, tk.END)
            range_entry.insert(0, format_page_range(selected, total_pages))
            update_summary()

        self.secondary_button(
            top_frame, "전체 선택",
            lambda: set_selection(range(1, total_pages + 1)),
            width=110, height=40, font_size=14
        ).grid(row=0, column=0, padx=4)

        self.secondary_button(
            top_frame, "전체 해제",
            lambda: set_selection([]),
            width=110, height=40, font_size=14
        ).grid(row=0, column=1, padx=4)

        self.secondary_button(
            top_frame, "홀수 쪽",
            lambda: set_selection(range(1, total_pages + 1, 2)),
            width=110, height=40, font_size=14
        ).grid(row=0, column=2, padx=4)

        self.secondary_button(
            top_frame, "짝수 쪽",
            lambda: set_selection(range(2, total_pages + 1, 2)),
            width=110, height=40, font_size=14
        ).grid(row=0, column=3, padx=4)

        # ── 직접 입력 ──
        input_frame = ctk.CTkFrame(win, fg_color="transparent")
        input_frame.pack(pady=(0, 12))

        ctk.CTkLabel(
            input_frame, text="직접 입력",
            font=(self.font_family, 14), text_color=COLOR_TEXT
        ).grid(row=0, column=0, padx=(0, 8))

        range_entry = self.make_entry(input_frame, width=300, height=40, font_size=14, justify="left")
        range_entry.grid(row=0, column=1, padx=4)
        range_entry.insert(0, format_page_range(selected, total_pages))

        def apply_range():
            try:
                pages = parse_page_range(range_entry.get(), total_pages)
            except ValueError as e:
                messagebox.showwarning("입력 오류", str(e), parent=win)
                return

            selected.clear()
            selected.update(pages)
            for num, var in check_vars.items():
                var.set(1 if num in selected else 0)
            update_summary()

        self.secondary_button(
            input_frame, "적용", apply_range, width=90, height=40, font_size=14
        ).grid(row=0, column=2, padx=4)

        range_entry.bind("<Return>", lambda e: apply_range())

        ctk.CTkLabel(
            win, text="예: 1-3,5,8-10  ·  비워두면 전체 페이지  ·  썸네일을 클릭하면 크게 볼 수 있습니다",
            font=(self.font_family, 12), text_color=COLOR_TEXT_MUTED
        ).pack(pady=(0, 12))

        # ── 썸네일 목록 (스크롤) ──
        scroll = ctk.CTkScrollableFrame(win, fg_color=COLOR_CARD, corner_radius=12, height=380)
        scroll.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        COLUMNS = 5
        THUMB_W, THUMB_H = 190, 250

        for num in range(1, total_pages + 1):
            row, col = divmod(num - 1, COLUMNS)

            cell = ctk.CTkFrame(scroll, fg_color="transparent")
            cell.grid(row=row, column=col, padx=8, pady=8)

            var = tk.IntVar(value=1 if num in selected else 0)
            check_vars[num] = var

            def toggle(n=num):
                if check_vars[n].get():
                    selected.add(n)
                else:
                    selected.discard(n)
                range_entry.delete(0, tk.END)
                range_entry.insert(0, format_page_range(selected, total_pages))
                update_summary()

            ctk.CTkCheckBox(
                cell, text=f"{num}쪽", variable=var, onvalue=1, offvalue=0,
                font=(self.font_family, 13), text_color=COLOR_TEXT,
                fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
                command=toggle
            ).pack(pady=(0, 4))

            # 썸네일 자리표시자 — 스크롤해서 보일 때 실제 이미지로 바뀐다
            holder = tk.Label(
                cell, text="", bg=COLOR_ENTRY_BG,
                borderwidth=1, relief="solid", cursor="hand2"
            )
            holder.configure(width=THUMB_W // 8, height=THUMB_H // 16)
            holder.pack()
            holder.bind("<Button-1>", lambda e, n=num: show_large(n))
            thumb_labels[num] = holder

        summary_label = ctk.CTkLabel(
            win, text="",
            font=(self.font_family, 15, "bold"), text_color=COLOR_TEXT
        )
        summary_label.pack(pady=(0, 10))

        def update_summary():
            count = len(selected)
            summary_label.configure(text=f"선택: {count}쪽 / 전체 {total_pages}쪽")

        update_summary()

        def show_large(page_num):
            """썸네일을 클릭하면 그 페이지를 크게 보여준다."""
            large = ctk.CTkToplevel(win)
            large.title(f"{page_num}쪽 - {file_name}")
            large.geometry("820x900")
            large.configure(fg_color=COLOR_BG)
            large.transient(win)
            large.attributes("-topmost", True)
            large.lift()
            large.after(120, large.grab_set)

            ctk.CTkLabel(
                large, text=f"{page_num}쪽 / 전체 {total_pages}쪽",
                font=(self.font_family, 17, "bold"), text_color=COLOR_TEXT
            ).pack(pady=(22, 12))

            status = ctk.CTkLabel(
                large, text="불러오는 중입니다...",
                font=(self.font_family, 14), text_color=COLOR_TEXT_MUTED
            )
            status.pack(pady=20)

            self.secondary_button(
                large, "닫기", large.destroy, width=150, height=46
            ).pack(side="bottom", pady=20)

            def load():
                try:
                    path = render_pdf_page_thumbnail(file_path, page_num - 1, max_width=760)
                except Exception as e:
                    self.root.after(0, lambda: status.configure(text=str(e)))
                    return
                self.root.after(0, lambda: place(path))

            def place(path):
                try:
                    if not large.winfo_exists():
                        return
                    with Image.open(path) as img:
                        img = img.copy()
                    img.thumbnail((760, 700))
                    photo = ImageTk.PhotoImage(img)

                    status.destroy()
                    holder = tk.Label(large, image=photo, bg=COLOR_CARD, borderwidth=0)
                    holder.image = photo
                    holder.pack(pady=6)
                except Exception:
                    pass

            threading.Thread(target=load, daemon=True).start()

        # ── 보이는 썸네일만 렌더링 ──
        def visible_pages():
            """스크롤 영역에서 지금 화면에 보이는 페이지 번호를 찾는다."""
            try:
                canvas = scroll._parent_canvas
                top = canvas.canvasy(0)
                bottom = top + canvas.winfo_height()
            except Exception:
                return list(range(1, min(total_pages, COLUMNS * 2) + 1))

            result = []
            for num, label in thumb_labels.items():
                try:
                    y = label.winfo_y()
                    h = label.winfo_height()
                except Exception:
                    continue

                # 화면 위아래로 한 화면씩 여유를 두고 미리 준비한다
                margin = canvas.winfo_height()
                if y + h >= top - margin and y <= bottom + margin:
                    result.append(num)

            return result

        def request_visible():
            if closed["value"]:
                return

            targets = [n for n in visible_pages() if n not in requested]

            for num in targets:
                requested.add(num)
                threading.Thread(target=render_one, args=(num,), daemon=True).start()

        def render_one(num):
            if closed["value"]:
                return

            try:
                path = render_pdf_page_thumbnail(file_path, num - 1, max_width=THUMB_W)
            except Exception:
                return

            if closed["value"]:
                return

            self.root.after(0, lambda: place_thumb(num, path))

        def place_thumb(num, path):
            if closed["value"] or num in rendered:
                return

            label = thumb_labels.get(num)
            if label is None:
                return

            try:
                if not label.winfo_exists():
                    return

                with Image.open(path) as img:
                    img = img.copy()
                img.thumbnail((THUMB_W, THUMB_H))
                photo = ImageTk.PhotoImage(img)

                photo_refs[num] = photo
                label.configure(image=photo, width=img.width, height=img.height)
                rendered.add(num)
            except Exception:
                pass

        # 스크롤할 때마다 새로 보이는 부분을 준비한다
        def on_scroll(event=None):
            if closed["value"]:
                return
            win.after(60, request_visible)

        try:
            scroll._parent_canvas.bind("<Configure>", on_scroll)
            scroll._parent_canvas.bind("<MouseWheel>", on_scroll, add="+")
        except Exception:
            pass

        win.after(200, request_visible)

        # 스크롤 위치가 바뀌는 것을 주기적으로도 확인한다
        # (마우스휠 외에 스크롤바 드래그 등으로 움직일 수 있다)
        def poll_scroll():
            if closed["value"]:
                return
            request_visible()
            win.after(700, poll_scroll)

        win.after(700, poll_scroll)

        # ── 하단 버튼 ──
        btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        btn_frame.pack(pady=(0, 20))

        def confirm():
            if not selected:
                messagebox.showwarning("선택 오류", "최소 한 페이지는 선택해야 합니다.", parent=win)
                return

            closed["value"] = True
            result = set(selected)
            win.destroy()
            on_confirm(result)

        def cancel():
            closed["value"] = True
            win.destroy()

        self.primary_button(btn_frame, "확인", confirm, width=160, height=48).grid(row=0, column=0, padx=8)
        self.secondary_button(btn_frame, "취소", cancel, width=160, height=48).grid(row=0, column=1, padx=8)

        win.protocol("WM_DELETE_WINDOW", cancel)

    def do_print_files(self, card_code, name, file_paths, duplex=False, page_counts=None,
                       converted_cache=None, selected_pages=None, file_copies=None):
        printer_name = get_setting("selected_printer", "")

        if not printer_name:
            messagebox.showerror("프린터 오류", "선택된 프린터가 없습니다. 관리자 화면에서 프린터를 선택하세요.")
            return

        success_count = 0
        fail_count = 0
        errors = []
        removed_files = []   # 출력이 끝나 바로 지운 받은 파일
        delayed_files = []   # 출력이 끝나 삭제를 예약한 받은 파일
        delay_minutes = 0

        for file_path in file_paths:
            file_name = os.path.basename(file_path)

            # 용지 장수 모드일 때만 실제 페이지 수를 안다.
            # 파일 개수 모드에서는 세지 않았으므로 1장으로 기록한다
            # (이 모드는 애초에 "몇 개를 출력했나"만 관리하기 때문이다).
            copies = max(1, int((file_copies or {}).get(file_path, 1)))

            if page_counts is not None:
                pages = page_counts.get(file_path)
                sheet_count = (sheets_from_pages(pages, duplex) or 1) * copies
            else:
                sheet_count = copies

            # 페이지를 골라 인쇄하는 경우 "1-3,5" 형태로 넘긴다.
            # (selected_pages 에는 전체 선택이 아닌 경우만 들어 있다)
            page_range = None
            if selected_pages:
                chosen = selected_pages.get(file_path)
                if chosen:
                    # total 을 실제 개수보다 크게 잡아 "전체"로 축약되지 않게 한다
                    page_range = format_page_range(chosen, max(chosen) + 1)

            try:
                status, printed_pdf = print_document(
                    file_path, printer_name, duplex=duplex,
                    converted_cache=converted_cache, page_range=page_range,
                    copies=copies
                )
                save_print_log(card_code, name, file_name, printer_name, "전송완료",
                                sheet_count=sheet_count, duplex=duplex)
                success_count += 1

                # 프린터가 작업을 받아간 것이 확인되면 받은 파일을 바로 지운다.
                # 30분을 기다리지 않으므로 남의 과제물이 남아 있는 시간이 줄어든다.
                if status == "done":
                    result, delay = cleanup_after_print(file_path, printed_pdf)

                    if result == "deleted":
                        removed_files.append(file_name)
                    elif result == "scheduled":
                        delayed_files.append(file_name)
                        delay_minutes = delay

            except Exception as e:
                save_print_log(card_code, name, file_name, printer_name, "실패",
                                sheet_count=sheet_count, duplex=duplex)
                fail_count += 1
                errors.append(f"{file_name}: {e}")

        # 인쇄가 끝난 변환 캐시는 더 이상 필요 없으므로 정리한다
        self._print_converted_cache = {}

        duplex_text = "양면" if duplex else "단면"

        total_copies = sum((file_copies or {}).get(p, 1) for p in file_paths)
        copies_text = f" · 총 {total_copies}부" if total_copies > len(file_paths) else ""

        message = (
            f"출력 전송 완료 ({duplex_text} 인쇄{copies_text})\n\n"
            f"성공: {success_count}개\n"
            f"실패: {fail_count}개"
        )

        if removed_files:
            message += (
                f"\n\n파일 공유로 받은 파일 {len(removed_files)}개는\n"
                "출력이 끝나 바로 삭제했습니다."
            )

        if delayed_files:
            message += (
                f"\n\n파일 공유로 받은 파일 {len(delayed_files)}개는\n"
                f"{delay_minutes}분 뒤에 삭제됩니다. 다시 뽑아야 하면 그 전에 출력하세요."
            )

        if errors:
            message += "\n\n오류 내용:\n" + "\n".join(errors[:5])

        if fail_count:
            # 실패한 출력은 횟수에서 빠지지만, 프린터에 문제가 있는 것일 수 있으므로
            # 다음 사람을 위해 신고하도록 안내한다.
            message += (
                "\n\n실패한 출력은 오늘 사용량에서 빠집니다.\n"
                "프린터에 문제가 있을 수 있으니 고장 신고를 해 주세요.\n"
                "화면 오른쪽 아래 '고장 신고' 버튼 또는 프린터 옆 QR 코드를 쓰면 됩니다."
            )

            answer = messagebox.askyesno(
                "출력 결과",
                message + "\n\n지금 고장 신고를 하시겠습니까?"
            )

            if answer:
                # 화면을 먼저 처음 상태로 되돌리고 브라우저를 연다.
                # 순서가 반대면 화면을 새로 그리는 사이에 '닫기' 버튼이 묻힌다.
                self.start_screen()
                self.open_report_form()
                return

            self.start_screen()
            return

        messagebox.showinfo("출력 결과", message)
        self.start_screen()

    # -----------------------------
    # 관리자 화면
    # -----------------------------
    def match_admin_code(self, card_code):
        """
        첫 화면에 입력된 값이 관리자코드인지 확인하고, 맞으면 기록에 남길 이름을 돌려준다.

        단일 방식  : 메인 관리자 코드와 비교 → '관리자'
        개별 방식  : 관리자 명단에서 찾음   → 그 사람 이름
                     (명단이 비어 있으면 잠기지 않도록 메인 코드를 받아 준다)
        """
        main_code = normalize_card_code(get_setting("admin_barcode", ""))

        if get_admin_mode() == "individual":
            name = find_admin_by_code(card_code)
            if name:
                return name

            if not list_admins() and main_code and main_code == card_code:
                return "관리자(메인 코드)"

            return None

        if main_code and main_code == card_code:
            return "관리자"

        return None

    def match_admin_password(self, password):
        """
        비밀번호가 누구의 것인지 확인한다. 돌려주는 값은 (이름, 들어온 방법).
        맞는 사람이 없으면 (None, "").
        """
        if not password:
            return None, ""

        real_password = get_setting("admin_password", DEFAULT_ADMIN_PASSWORD)

        def main_ok():
            if not verify_password(password, real_password):
                return False
            # 예전 버전에서 평문으로 저장된 비밀번호는 이때 해시로 바꿔 둔다
            if is_legacy_password(real_password):
                set_setting("admin_password", hash_password(password))
            return True

        if get_admin_mode() == "individual":
            name = find_admin_by_password(password)
            if name:
                return name, "관리자 비밀번호"

            # 명단이 비어 있거나 읽지 못하면 아무도 못 들어오게 되므로,
            # 그때만 메인 비밀번호를 받아 준다.
            if not list_admins() and main_ok():
                return "관리자(메인 비밀번호)", "메인 비밀번호 · 명단 없음"

            return None, ""

        if main_ok():
            return "관리자", "비밀번호"

        return None, ""

    def admin_login_screen(self):
        card = self.build_card()
        self.make_title(card, "관리자 로그인")

        if get_admin_mode() == "individual" and list_admins():
            prompt = "본인의 관리자 비밀번호를 입력하세요"
        else:
            prompt = "관리자 비밀번호 입력"

        self.body_label(card, prompt, size=15).pack(pady=(0, 14))

        password_entry = self.make_entry(card, width=280, height=48, font_size=16, show="*")
        password_entry.pack(pady=(0, 24))
        password_entry.focus()

        def check_password():
            password = password_entry.get()
            who, via = self.match_admin_password(password)

            if who:
                begin_admin_session(who, via)
                self.admin_menu()
            else:
                admin_log_as("-", "로그인 실패", "비밀번호가 틀림")
                messagebox.showerror("접근 거부", "관리자 비밀번호가 틀렸습니다.")
                password_entry.delete(0, tk.END)

        password_entry.bind("<Return>", lambda event: check_password())

        self.primary_button(card, "로그인", check_password, width=240, height=54).pack(pady=(0, 14))
        self.link_button(card, "처음 화면으로", self.start_screen, width=180, height=42).pack()

    def admin_menu(self):
        # 로그인하지 않고 관리자 화면에 들어오는 길은 막는다
        if not admin_session_active():
            self.admin_login_screen()
            return

        card = self.build_card()
        self.make_title(card, "관리자 모드")

        self.body_label(
            card, f"{describe_admin(current_admin()) or current_admin()} 님으로 들어왔습니다 · 이 화면에서 한 작업은 모두 기록됩니다",
            size=13, muted=True
        ).pack(pady=(0, 14))

        buttons = [
            ("이용자 관리", self.user_manage_screen),
            ("출력 파일 설정", self.limit_setting_screen),
            ("프린터 설정", self.printer_setting_screen),
            ("출력 기록 관리", self.log_manage_screen),
            ("비밀번호 설정", self.change_admin_password_screen),
            ("보호 기능 설정", self.kiosk_setting_screen),
        ]

        grid_frame = ctk.CTkFrame(card, fg_color="transparent")
        grid_frame.pack(pady=(0, 24))

        for i, (text, cmd) in enumerate(buttons):
            row, col = divmod(i, 2)
            self.secondary_button(
                grid_frame, text, cmd, width=340, height=64, font_size=18
            ).grid(row=row, column=col, padx=10, pady=10)

        # 관리자 방식 · 명단 · 작업 기록 (두 칸을 합친 넓은 버튼)
        self.secondary_button(
            grid_frame, "관리자 계정 · 작업 기록", self.admin_account_screen,
            width=700, height=56, font_size=17
        ).grid(row=len(buttons) // 2, column=0, columnspan=2, padx=10, pady=(6, 10))

        bottom_frame = ctk.CTkFrame(card, fg_color="transparent")
        bottom_frame.pack()

        self.link_button(bottom_frame, "처음 화면으로", self.start_screen, width=200).grid(row=0, column=0, padx=10)
        self.danger_button(bottom_frame, "프로그램 종료", self.admin_exit_program, width=200).grid(row=0, column=1, padx=10)

    def user_manage_screen(self):
        card = self.build_card()
        self.make_title(card, "이용자 관리")

        # 입력 항목을 가로 3칸으로 늘어놓는다.
        # 세로로 쌓으면 FHD(1080p) 화면에서 아래가 잘린다.
        input_frame = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=14)
        input_frame.pack(pady=(0, 8), ipadx=16, ipady=12)

        self.body_label(input_frame, "학생증 코드", size=16).grid(row=0, column=0, padx=(12, 8), pady=(4, 2))
        self.body_label(input_frame, "학번", size=16).grid(row=0, column=1, padx=8, pady=(4, 2))
        self.body_label(input_frame, "이름", size=16).grid(row=0, column=2, padx=8, pady=(4, 2))

        card_entry = self.make_entry(input_frame, width=230, height=44, font_size=15, justify="left")
        card_entry.grid(row=1, column=0, padx=(12, 8), pady=(0, 6))

        number_entry = self.make_entry(input_frame, width=190, height=44, font_size=15, justify="left")
        number_entry.grid(row=1, column=1, padx=8, pady=(0, 6))

        name_entry = self.make_entry(input_frame, width=190, height=44, font_size=15, justify="left")
        name_entry.grid(row=1, column=2, padx=8, pady=(0, 6))

        unlimited_var = tk.IntVar(value=0)
        ctk.CTkCheckBox(
            input_frame,
            text="무제한 출력 사용자",
            variable=unlimited_var,
            onvalue=1,
            offvalue=0,
            font=(self.font_family, 16),
            text_color=COLOR_TEXT,
            fg_color=COLOR_PRIMARY,
            hover_color=COLOR_PRIMARY_HOVER
        ).grid(row=1, column=3, padx=(16, 12), pady=(0, 6))

        search_frame = ctk.CTkFrame(card, fg_color="transparent")
        search_frame.pack(pady=(0, 6))

        self.body_label(search_frame, "학생증 코드", size=17, bold=True).grid(row=0, column=0, padx=(0, 10))
        search_entry = self.make_entry(search_frame, width=260, height=44, font_size=15, justify="left")
        search_entry.grid(row=0, column=1, padx=6)

        tree = ttk.Treeview(
            card,
            columns=("card_code", "student_number", "name", "unlimited", "active", "created_date"),
            show="headings",
            height=9,
            style="App.Treeview"
        )

        tree.heading("card_code", text="학생증 코드")
        tree.heading("student_number", text="학번")  
        tree.heading("name", text="이름")
        tree.heading("unlimited", text="무제한")
        tree.heading("active", text="상태")
        tree.heading("created_date", text="등록일")

        tree.column("card_code", width=220, minwidth=220, stretch=False)
        tree.column("student_number", width=150, minwidth=150, stretch=False)
        tree.column("name", width=170, minwidth=170, stretch=False)
        tree.column("unlimited", width=110, minwidth=110, anchor="center", stretch=False)
        tree.column("active", width=140, minwidth=140, anchor="center", stretch=False)
        tree.column("created_date", width=170, minwidth=170, anchor="center", stretch=False)

        def refresh_users(keyword=None):
            for row in tree.get_children():
                tree.delete(row)

            conn = connect_db()
            cur = conn.cursor()

            if keyword:
                like_keyword = f"%{keyword}%"
                cur.execute("""
                SELECT student_id, student_number, name, unlimited, active, created_date
                FROM students
                WHERE student_id LIKE ? OR student_number LIKE ?
                ORDER BY student_id
                """, (like_keyword, like_keyword))
            else:
                cur.execute("""
                SELECT student_id, student_number, name, unlimited, active, created_date
                FROM students
                ORDER BY student_id
                """)

            rows = cur.fetchall()
            conn.close()

            for card_code, student_number, name, unlimited, active, created in rows:
                tree.insert(
                    "",
                    "end",
                    values=(
                        card_code,
                        student_number or "",
                        name,
                        "예" if unlimited else "아니오",
                        "사용 가능" if active else "사용 중지",
                        created
                    )
                )

        def search_students():
            keyword = normalize_card_code(search_entry.get())
            refresh_users(keyword if keyword else None)

        def show_all_students():
            search_entry.delete(0, tk.END)
            refresh_users()

        self.secondary_button(search_frame, "검색", search_students, width=100, height=40, font_size=16).grid(row=0, column=2, padx=6)
        self.secondary_button(search_frame, "전체 보기", show_all_students, width=100, height=40, font_size=16).grid(row=0, column=3, padx=6)

        tree.pack(pady=(0, 6))

        def on_select(event):
            selected = tree.selection()
            if not selected:
                return

            values = tree.item(selected[0], "values")

            card_entry.delete(0, tk.END)
            card_entry.insert(0, values[0])

            number_entry.delete(0, tk.END)
            number_entry.insert(0, values[1])

            name_entry.delete(0, tk.END)
            name_entry.insert(0, values[2])

            unlimited_var.set(1 if values[3] == "예" else 0)

        tree.bind("<<TreeviewSelect>>", on_select)

        def save_user():
            card_code = normalize_card_code(card_entry.get())
            student_number = normalize_card_code(number_entry.get())
            name = name_entry.get().strip()
            unlimited = unlimited_var.get()

            if not card_code or not name:
                messagebox.showwarning("입력 오류", "학생증 코드와 이름을 입력하세요.")
                return

            try:
                existing = get_student(card_code)
                active = existing[2] if existing else 1

                upsert_student(card_code, name, student_number=student_number, active=active, unlimited=unlimited)

                write_sync_csv()

                admin_log(
                    "이용자 수정" if existing else "이용자 추가",
                    f"{student_number} {name}({card_code})" + (" · 무제한" if unlimited else "")
                )

                messagebox.showinfo("완료", "이용자 정보를 저장했습니다.")
                card_entry.delete(0, tk.END)
                number_entry.delete(0, tk.END)
                name_entry.delete(0, tk.END)
                unlimited_var.set(0)
                refresh_users()

            except Exception as e:
                messagebox.showerror("오류", str(e))

        def set_active(value):
            card_code = normalize_card_code(card_entry.get())

            if not card_code:
                messagebox.showwarning("입력 오류", "학생증 코드를 입력하세요.")
                return

            conn = connect_db()
            cur = conn.cursor()
            cur.execute("UPDATE students SET active=? WHERE student_id=?", (value, card_code))
            conn.commit()
            changed = cur.rowcount
            conn.close()

            if changed == 0:
                messagebox.showerror("오류", "해당 학생증 코드의 사용자가 없습니다.")
            else:
                write_sync_csv()
                admin_log("이용자 상태 변경", f"{card_code} → {'사용 가능' if value else '사용 중지'}")
                messagebox.showinfo("완료", "상태를 변경했습니다.")
                refresh_users()

        def remove_user():
            """이용자를 명단에서 완전히 지운다. 되돌릴 수 없으므로 두 번 확인한다."""
            card_code = normalize_card_code(card_entry.get())

            if not card_code:
                messagebox.showwarning("입력 오류", "지울 이용자를 목록에서 고르거나 학생증 코드를 입력하세요.")
                return

            student = get_student(card_code)

            if not student:
                messagebox.showerror("오류", "해당 학생증 코드의 사용자가 없습니다.")
                return

            name = student[1]

            answer = messagebox.askyesno(
                "이용자 삭제",
                f"{name} ({card_code}) 님을 명단에서 완전히 지웁니다.\n\n"
                "지운 뒤에는 되돌릴 수 없고, 다시 쓰려면 새로 등록해야 합니다.\n"
                "출력 기록은 그대로 남습니다.\n\n"
                "잠시 못 쓰게 하려는 것이라면 '비활성화' 를 쓰세요.\n\n"
                "정말 지울까요?"
            )

            if not answer:
                return

            try:
                deleted = delete_student(card_code)
            except Exception as e:
                messagebox.showerror("오류", str(e))
                return

            if deleted == 0:
                messagebox.showerror("오류", "해당 학생증 코드의 사용자가 없습니다.")
                return

            write_sync_csv()

            card_entry.delete(0, tk.END)
            number_entry.delete(0, tk.END)
            name_entry.delete(0, tk.END)
            unlimited_var.set(0)

            admin_log("이용자 삭제", f"{name}({card_code})")
            messagebox.showinfo("완료", f"{name} ({card_code}) 님을 명단에서 지웠습니다.")
            refresh_users()

        def open_sync_csv():
            """연동용 CSV 파일을 엑셀 등 기본 프로그램으로 연다."""
            path = get_sync_csv_path()

            if not os.path.exists(path):
                write_sync_csv()

            try:
                os.startfile(path)
            except Exception:
                messagebox.showinfo("안내", f"CSV 파일 위치\n\n{path}")

        def import_users():
            path = filedialog.askopenfilename(
                title="이용자 명단 불러오기",
                filetypes=[
                    ("CSV/XLSX 파일", "*.csv *.xlsx"),
                    ("CSV 파일", "*.csv"),
                    ("Excel 파일", "*.xlsx")
                ]
            )

            if not path:
                return

            try:
                imported, skipped = import_users_from_file(path)
                admin_log("이용자 명단 불러오기", f"{os.path.basename(path)} · {imported}명 반영, {skipped}줄 건너뜀")
                messagebox.showinfo(
                    "불러오기 완료",
                    f"불러온 이용자: {imported}명\n건너뛴 행: {skipped}개"
                )
                write_sync_csv()
                refresh_users()

            except Exception as e:
                messagebox.showerror("불러오기 오류", str(e))

        def export_users():
            path = filedialog.asksaveasfilename(
                title="이용자 명단 CSV 내보내기",
                defaultextension=".csv",
                filetypes=[("CSV 파일", "*.csv")]
            )

            if not path:
                return

            try:
                export_users_to_csv(path)
                admin_log("이용자 명단 내보내기", os.path.basename(path))
                messagebox.showinfo("내보내기 완료", "이용자 명단을 CSV로 저장했습니다.")
            except Exception as e:
                messagebox.showerror("내보내기 오류", str(e))

        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.pack(pady=(2, 4))

        # 줄 수를 늘리면 화면 아래가 잘리므로 삭제 버튼도 같은 줄에 넣는다.
        # 실수로 누르지 않도록 '비활성화' 와 떨어뜨려 맨 끝에 둔다.
        self.primary_button(btn_frame, "추가/수정 저장", save_user, width=150, height=44, font_size=16).grid(row=0, column=0, padx=5)
        self.danger_button(btn_frame, "비활성화", lambda: set_active(0), width=120, height=44, font_size=16).grid(row=0, column=1, padx=5)
        self.secondary_button(btn_frame, "다시 활성화", lambda: set_active(1), width=125, height=44, font_size=16).grid(row=0, column=2, padx=5)
        self.secondary_button(btn_frame, "새로고침", refresh_users, width=110, height=44, font_size=16).grid(row=0, column=3, padx=5)
        self.danger_button(btn_frame, "삭제", remove_user, width=110, height=44, font_size=16).grid(row=0, column=4, padx=(22, 5))

        file_btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        file_btn_frame.pack(pady=(0, 6))

        self.secondary_button(file_btn_frame, "CSV/XLSX 불러오기", import_users, width=200, height=42, font_size=16).grid(row=0, column=0, padx=5)
        self.secondary_button(file_btn_frame, "명단 CSV 내보내기", export_users, width=200, height=42, font_size=16).grid(row=0, column=1, padx=5)
        self.secondary_button(file_btn_frame, "연동 CSV 열기", open_sync_csv, width=170, height=42, font_size=16).grid(row=0, column=2, padx=5)

        self.body_label(
            card,
            f"{get_sync_csv_name()} 실시간 연동 중  ·  파일을 고쳐 저장하면 몇 초 안에 명단에 반영됩니다 "
            "(줄을 지워도 삭제되지는 않습니다)",
            size=13, muted=True
        ).pack(pady=(0, 6))

        self.link_button(card, "관리자 메뉴로", self.admin_menu, width=180, height=38).pack()

        # CSV 가 밖에서 바뀌었을 때 이 화면의 목록도 함께 새로 그리도록 등록해 둔다
        self._user_list_refresh = refresh_users

        if not os.path.exists(get_sync_csv_path()):
            write_sync_csv()

        refresh_users()

    def limit_setting_screen(self):
        """
        출력 파일 설정.

        항목이 많아 세로로 쌓으면 FHD 화면에서 잘리므로 좌우 두 칸으로 나눈다.
          왼쪽 : 출력 제한 · 파일 선택 기본 폴더
          오른쪽: 받은 파일 보관 (다운로드 폴더 · 자동 삭제)
        """
        card = self.build_card()
        self.make_title(card, "출력 파일 설정")

        columns = ctk.CTkFrame(card, fg_color="transparent")
        columns.pack(pady=(0, 2))

        left = ctk.CTkFrame(columns, fg_color="transparent")
        left.grid(row=0, column=0, padx=(0, 24), sticky="n")

        ctk.CTkFrame(columns, fg_color=COLOR_BORDER, width=1).grid(
            row=0, column=1, sticky="ns", pady=6
        )

        right = ctk.CTkFrame(columns, fg_color="transparent")
        right.grid(row=0, column=2, padx=(24, 0), sticky="n")

        # ==========================================================
        # 왼쪽 — 출력 제한
        # ==========================================================
        self.body_label(left, "출력 제한", size=20, bold=True).pack(pady=(0, 10))

        limit_mode_var = tk.StringVar(value=get_limit_mode())

        mode_frame = ctk.CTkFrame(left, fg_color="transparent")
        mode_frame.pack(pady=(0, 12))

        ctk.CTkRadioButton(
            mode_frame, text="파일 개수로 제한", variable=limit_mode_var, value="files",
            font=(self.font_family, 14), text_color=COLOR_TEXT,
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
            command=lambda: refresh_mode_ui()
        ).grid(row=0, column=0, padx=(0, 14))

        ctk.CTkRadioButton(
            mode_frame, text="용지 장수로 제한", variable=limit_mode_var, value="pages",
            font=(self.font_family, 14), text_color=COLOR_TEXT,
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
            command=lambda: refresh_mode_ui()
        ).grid(row=0, column=1)

        frame = ctk.CTkFrame(left, fg_color=COLOR_ENTRY_BG, corner_radius=14)
        frame.pack(pady=(0, 6), ipadx=14, ipady=12)

        # 안쪽을 grid 로 채우면 ipadx/ipady 로 늘어난 여백이 전부
        # 오른쪽·아래쪽에만 남는다. grid_anchor 를 center 로 두면
        # 그 여백을 사방으로 나눠 가져서 내용이 상자 한가운데 온다.
        frame.grid_anchor("center")

        self.body_label(frame, "하루 출력 제한 (파일 수)", size=14).grid(row=0, column=0, padx=10, pady=4, sticky="e")
        daily_entry = self.make_entry(frame, width=110, height=42, font_size=15)
        daily_entry.grid(row=0, column=1, padx=10, pady=4)
        daily_entry.insert(0, str(get_daily_limit()))

        page_limit_label = self.body_label(frame, "하루 출력 제한 (용지 장수)", size=14)
        page_limit_label.grid(row=1, column=0, padx=10, pady=4, sticky="e")
        page_limit_entry = self.make_entry(frame, width=110, height=42, font_size=15)
        page_limit_entry.grid(row=1, column=1, padx=10, pady=4)
        page_limit_entry.insert(0, str(get_daily_page_limit()))

        def refresh_mode_ui():
            # 지금 적용되는 모드의 입력칸만 강조하고, 안 쓰는 쪽은 흐리게 표시한다.
            # (두 값 다 저장은 되므로 모드를 나중에 바꿔도 이전 설정이 남아 있다)
            using_pages = limit_mode_var.get() == "pages"

            daily_entry.configure(
                text_color=(COLOR_TEXT_MUTED if using_pages else COLOR_TEXT)
            )
            page_limit_entry.configure(
                text_color=(COLOR_TEXT if using_pages else COLOR_TEXT_MUTED)
            )

        refresh_mode_ui()

        self.body_label(frame, "한 번에 선택 가능한 파일 수", size=14).grid(row=2, column=0, padx=10, pady=4, sticky="e")
        max_file_entry = self.make_entry(frame, width=110, height=42, font_size=15)
        max_file_entry.grid(row=2, column=1, padx=10, pady=4)
        max_file_entry.insert(0, str(get_max_files_per_job()))

        self.body_label(frame, "파일 하나당 최대 부수", size=14).grid(row=3, column=0, padx=10, pady=4, sticky="e")
        max_copies_entry = self.make_entry(frame, width=110, height=42, font_size=15)
        max_copies_entry.grid(row=3, column=1, padx=10, pady=4)
        max_copies_entry.insert(0, str(get_max_copies()))

        self.body_label(
            left,
            "0 을 입력하면 그 항목은 제한하지 않습니다.\n"
            "용지 장수 기준일 때는 양면 인쇄 시 2쪽이 1장으로 계산됩니다.\n"
            "무제한 출력 사용자는 위 제한을 적용받지 않습니다.",
            size=13, muted=True, wraplength=340
        ).pack(pady=(6, 6))

        # ── 왼쪽 — 파일 선택 기본 폴더 ──
        ctk.CTkFrame(left, fg_color=COLOR_BORDER, height=1).pack(fill="x", pady=(0, 12))

        self.body_label(left, "파일 선택 기본 폴더", size=20, bold=True).pack(pady=(0, 8))

        current_dir = get_setting("file_open_dir", "")
        self.body_label(
            left,
            f"현재: {current_dir if current_dir else '설정 안 됨 (바탕화면)'}",
            size=13, muted=True, wraplength=340
        ).pack(pady=(0, 10))

        dir_entry = self.make_entry(left, width=340, height=44, font_size=14, justify="left")
        dir_entry.pack(pady=(0, 8))
        if current_dir:
            dir_entry.insert(0, current_dir)

        self.body_label(
            left,
            "학생이 '문서 파일 선택'을 눌렀을 때 열리는 폴더입니다.\n"
            "비워두면 바탕화면에서 시작합니다.",
            size=13, muted=True
        ).pack(pady=(0, 12))

        def browse_open_dir():
            selected = filedialog.askdirectory(title="파일 선택 기본 폴더")
            if selected:
                dir_entry.delete(0, tk.END)
                dir_entry.insert(0, selected.replace("/", "\\"))

        self.secondary_button(left, "폴더 선택", browse_open_dir, width=150, height=42).pack()

        # ==========================================================
        # 오른쪽 — 받은 파일 보관
        # ==========================================================
        self.body_label(right, "받은 파일 저장 폴더", size=20, bold=True).pack(pady=(0, 8))

        current_download = get_setting("download_dir", "")
        actual_download = get_download_dir()

        self.body_label(
            right,
            f"현재: {actual_download}",
            size=13, muted=True, wraplength=340
        ).pack(pady=(0, 10))

        download_entry = self.make_entry(right, width=340, height=44, font_size=14, justify="left")
        download_entry.pack(pady=(0, 8))
        if current_download:
            download_entry.insert(0, current_download)

        self.body_label(
            right,
            "파일 공유 기능을 통해 받은 파일이 저장되는 폴더입니다.\n"
            "비워두면 프로그램 폴더 안 ReceivedFiles 를 씁니다.",
            size=13, muted=True
        ).pack(pady=(0, 12))

        def browse_download_dir():
            selected = filedialog.askdirectory(title="받은 파일 저장 폴더")
            if selected:
                download_entry.delete(0, tk.END)
                download_entry.insert(0, selected.replace("/", "\\"))

        self.secondary_button(right, "폴더 선택", browse_download_dir, width=150, height=42).pack(pady=(0, 20))

        # ── 오른쪽 — 자동 삭제 ──
        ctk.CTkFrame(right, fg_color=COLOR_BORDER, height=1).pack(fill="x", pady=(0, 16))

        self.body_label(right, "받은 파일 자동 삭제", size=20, bold=True).pack(pady=(0, 4))

        self.body_label(
            right,
            "아래 두 가지가 함께 적용됩니다. 둘 중 먼저 오는 때에 지워집니다.",
            size=13, muted=True
        ).pack(pady=(0, 10))

        delete_frame = ctk.CTkFrame(right, fg_color=COLOR_ENTRY_BG, corner_radius=14)
        delete_frame.pack(pady=(0, 10), ipadx=14, ipady=12)
        delete_frame.grid_anchor("center")

        self.body_label(delete_frame, "보관 시간 (분)", size=14).grid(row=0, column=0, padx=10, pady=4, sticky="e")
        delete_entry = self.make_entry(delete_frame, width=110, height=42, font_size=15)
        delete_entry.grid(row=0, column=1, padx=10, pady=4)
        delete_entry.insert(0, str(get_auto_delete_minutes()))

        self.body_label(
            right,
            "출력 여부와 상관없이, 파일을 다 받은 시점부터 이 시간이 지나면 지웁니다.\n"
            "뽑으려다 그만둔 파일은 이걸로 정리됩니다. 0 을 넣으면 쓰지 않습니다.",
            size=13, muted=True
        ).pack(pady=(0, 10))

        # ── 삭제 시점 고르기 ──
        self.body_label(right, "출력이 끝난 파일은 추가로", size=14, bold=True).pack(pady=(0, 4))

        delete_mode_var = tk.StringVar(value=get_delete_mode())

        mode_box = ctk.CTkFrame(right, fg_color="transparent")
        mode_box.pack(pady=(0, 4))

        ctk.CTkRadioButton(
            mode_box, text="지우지 않는다", variable=delete_mode_var, value="time",
            font=(self.font_family, 14), text_color=COLOR_TEXT,
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
            command=lambda: refresh_delete_ui()
        ).grid(row=0, column=0, padx=(0, 18))

        ctk.CTkRadioButton(
            mode_box, text="바로 지운다", variable=delete_mode_var, value="print",
            font=(self.font_family, 14), text_color=COLOR_TEXT,
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
            command=lambda: refresh_delete_ui()
        ).grid(row=0, column=1)

        delay_row = ctk.CTkFrame(right, fg_color="transparent")
        delay_row.pack(pady=(2, 6))

        ctk.CTkRadioButton(
            delay_row, text="", width=22, variable=delete_mode_var, value="print_delay",
            font=(self.font_family, 14), text_color=COLOR_TEXT,
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
            command=lambda: refresh_delete_ui()
        ).grid(row=0, column=0, padx=(0, 4))

        print_delay_entry = self.make_entry(delay_row, width=68, height=34, font_size=14)
        print_delay_entry.grid(row=0, column=1)
        print_delay_entry.insert(0, str(get_print_delete_minutes()))

        self.body_label(delay_row, "분 뒤에 지운다", size=14).grid(row=0, column=2, padx=(8, 0))

        delete_hint = self.body_label(right, "", size=13, muted=True, wraplength=340)
        delete_hint.pack(pady=(0, 6))

        def describe_delete_rule():
            """두 설정을 합쳐 실제로 언제 지워지는지 한 줄로 보여준다."""
            try:
                hold = int(delete_entry.get().strip())
            except ValueError:
                hold = -1

            try:
                delay = int(print_delay_entry.get().strip())
            except ValueError:
                delay = -1

            mode = delete_mode_var.get()

            hold_text = f"받은 뒤 {hold}분" if hold > 0 else ""

            if mode == "print":
                print_text = "출력 직후"
            elif mode == "print_delay" and delay == 0:
                print_text = "출력 직후"
            elif mode == "print_delay" and delay > 0:
                print_text = f"출력 {delay}분 뒤"
            else:
                print_text = ""

            if hold_text and print_text:
                return f"→ {hold_text} · {print_text} 중 먼저 오는 때에 삭제"

            if hold_text or print_text:
                return f"→ {hold_text or print_text}에 삭제"

            return "→ 자동으로 지우지 않음 (직접 정리해야 합니다)"

        def refresh_delete_ui(event=None):
            mode = delete_mode_var.get()
            print_delay_entry.configure(state="normal" if mode == "print_delay" else "disabled")

            hints = {
                "time": "출력해도 그대로 두고 보관 시간만 적용합니다.\n"
                        "다시 뽑기는 편하지만, 그때까지 다른 학생도 그 파일을 볼 수 있습니다.",
                "print": "출력이 끝나는 즉시 지워져 가장 안전합니다.\n"
                         "용지가 걸려 다시 뽑아야 하면 파일을 다시 보내야 합니다.",
                "print_delay": "출력 직후 잠깐만 남겨 둡니다.\n"
                               "다시 뽑을 여유를 주면서도 오래 남지 않습니다. (권장)",
            }
            delete_hint.configure(text=describe_delete_rule() + "\n" + hints.get(mode, ""))

        delete_entry.bind("<KeyRelease>", refresh_delete_ui)
        print_delay_entry.bind("<KeyRelease>", refresh_delete_ui)
        refresh_delete_ui()

        # ==========================================================
        # 하단 공통
        # ==========================================================
        def save_limits():
            daily_value = daily_entry.get().strip()
            page_limit_value = page_limit_entry.get().strip()
            max_file_value = max_file_entry.get().strip()
            max_copies_value = max_copies_entry.get().strip()
            delete_value = delete_entry.get().strip()
            mode_value = limit_mode_var.get()

            # 0 은 "제한 없음" 을 뜻하므로 허용한다
            if not daily_value.isdigit():
                messagebox.showwarning("입력 오류", "하루 출력 제한(파일 수)은 0 이상의 숫자여야 합니다.")
                return

            if not page_limit_value.isdigit():
                messagebox.showwarning("입력 오류", "하루 출력 제한(용지 장수)은 0 이상의 숫자여야 합니다.")
                return

            if not max_file_value.isdigit():
                messagebox.showwarning("입력 오류", "파일 개수 제한은 0 이상의 숫자여야 합니다.")
                return

            if not max_copies_value.isdigit():
                messagebox.showwarning("입력 오류", "부수 제한은 0 이상의 숫자여야 합니다.")
                return

            # 지금 고른 모드의 하루 제한과 한 번에 선택 가능 개수가 모두 0 이면
            # 사실상 아무 제한이 없다는 뜻이므로 한 번 확인한다
            active_limit = page_limit_value if mode_value == "pages" else daily_value
            if int(active_limit) == 0 and int(max_file_value) == 0:
                if not messagebox.askyesno(
                    "확인",
                    "하루 제한과 한 번에 선택 가능한 파일 수를 모두 0 으로 두면\n"
                    "학생이 제한 없이 출력할 수 있습니다.\n\n"
                    "이대로 저장하시겠습니까?"
                ):
                    return

            if not delete_value.isdigit():
                messagebox.showwarning("입력 오류", "보관 시간은 0 이상의 숫자여야 합니다.")
                return

            delete_mode_value = delete_mode_var.get()
            print_delay_value = print_delay_entry.get().strip()

            if delete_mode_value == "print_delay" and not print_delay_value.isdigit():
                messagebox.showwarning("입력 오류", "출력 후 유예 시간은 0 이상의 숫자여야 합니다.")
                return

            if not print_delay_value.isdigit():
                print_delay_value = str(get_print_delete_minutes())

            hold_minutes = int(delete_value)

            # 어느 쪽으로도 지우지 않게 되면 학생들이 보낸 파일이 계속 쌓인다
            if hold_minutes == 0 and delete_mode_value == "time":
                if not messagebox.askyesno(
                    "확인",
                    "받은 파일을 자동으로 지우지 않게 됩니다.\n"
                    "학생들이 보낸 과제물이 이 컴퓨터에 계속 남으므로\n"
                    "담당자가 직접 정리해야 합니다.\n\n"
                    "이대로 저장하시겠습니까?"
                ):
                    return

            # 출력 후에만 지우면, 보내놓고 뽑지 않은 파일은 그대로 남는다
            if hold_minutes == 0 and delete_mode_value != "time":
                if not messagebox.askyesno(
                    "확인",
                    "보관 시간을 0 으로 두면 출력한 파일만 지워집니다.\n"
                    "학생이 파일을 보내놓고 뽑지 않고 가면 그 파일은 계속 남습니다.\n\n"
                    "이대로 저장하시겠습니까?"
                ):
                    return

            # 유예가 보관 시간보다 길면 보관 시간이 먼저 와서 유예가 무의미해진다
            if delete_mode_value == "print_delay" and hold_minutes > 0 \
                    and int(print_delay_value) >= hold_minutes:
                if not messagebox.askyesno(
                    "확인",
                    f"출력 후 유예({print_delay_value}분)가 보관 시간({delete_value}분)보다 깁니다.\n"
                    "이 경우 보관 시간이 먼저 와서 그때 지워집니다.\n\n"
                    "이대로 저장하시겠습니까?"
                ):
                    return

            dir_value = dir_entry.get().strip()
            if dir_value and not os.path.isdir(dir_value):
                messagebox.showwarning("경로 오류", f"폴더를 찾을 수 없습니다.\n{dir_value}")
                return

            download_value = download_entry.get().strip()
            if download_value:
                try:
                    os.makedirs(download_value, exist_ok=True)
                except Exception as e:
                    messagebox.showwarning("경로 오류", f"폴더를 만들 수 없습니다.\n{e}")
                    return

            set_setting("limit_mode", mode_value)
            set_setting("daily_limit", daily_value)
            set_setting("daily_page_limit", page_limit_value)
            set_setting("max_files_per_job", max_file_value)
            set_setting("max_copies", max_copies_value)
            set_setting("file_open_dir", dir_value)
            set_setting("download_dir", download_value)
            set_setting("auto_delete_minutes", delete_value)
            set_setting("delete_mode", delete_mode_value)
            set_setting("print_delete_minutes", print_delay_value)

            hold_text = f"받은 뒤 {delete_value}분" if hold_minutes > 0 else ""

            if delete_mode_value == "print":
                print_text = "출력 직후"
            elif delete_mode_value == "print_delay":
                print_text = ("출력 직후" if int(print_delay_value) == 0
                              else f"출력 후 {print_delay_value}분")
            else:
                print_text = ""

            if hold_text and print_text:
                delete_text = f"{hold_text} · {print_text} 중 먼저 오는 때에 삭제"
            elif hold_text or print_text:
                delete_text = f"{hold_text or print_text}에 삭제"
            else:
                delete_text = "자동으로 지우지 않음"

            if mode_value == "pages":
                limit_summary = (
                    "하루 출력 제한: 용지 제한 없음" if int(page_limit_value) == 0
                    else f"하루 출력 제한: 용지 {page_limit_value}장 (양면 인쇄 시 2쪽 = 1장)"
                )
            else:
                limit_summary = (
                    "하루 출력 제한: 파일 개수 제한 없음" if int(daily_value) == 0
                    else f"하루 출력 제한: 파일 {daily_value}개"
                )

            max_summary = (
                "제한 없음" if int(max_file_value) == 0 else f"{max_file_value}개"
            )
            copies_summary = (
                "제한 없음" if int(max_copies_value) == 0 else f"{max_copies_value}부"
            )

            messagebox.showinfo(
                "완료",
                f"{limit_summary}\n"
                f"한 번에 선택 가능한 파일 수: {max_summary}\n"
                f"파일 하나당 최대 부수: {copies_summary}\n"
                f"파일 선택 기본 폴더: {dir_value if dir_value else '바탕화면'}\n"
                f"받은 파일 저장 폴더: {get_download_dir()}\n"
                f"받은 파일: {delete_text}"
            )
            self.admin_menu()

        ctk.CTkFrame(card, fg_color=COLOR_BORDER, height=1).pack(fill="x", pady=(12, 10))

        self.primary_button(card, "저장", save_limits, width=200, height=52).pack(pady=(0, 8))
        self.link_button(card, "관리자 메뉴로", self.admin_menu, width=180, height=42).pack()

    def printer_setting_screen(self):
        card = self.build_card()
        self.make_title(card, "프린터 설정")

        current = get_setting("selected_printer", "")

        self.body_label(
            card, f"현재 선택된 프린터:\n{current}", size=15, justify="center"
        ).pack(pady=(0, 18))

        try:
            printers = get_printer_list()
            error_message = ""
        except Exception as e:
            printers = []
            error_message = str(e)

        combo = ctk.CTkComboBox(
            card,
            values=printers if printers else [""],
            width=600,
            height=48,
            corner_radius=12,
            font=(self.font_family, 14),
            dropdown_font=(self.font_family, 13),
            fg_color=COLOR_ENTRY_BG,
            border_color=COLOR_BORDER,
            button_color=COLOR_PRIMARY,
            button_hover_color=COLOR_PRIMARY_HOVER,
            dropdown_fg_color=COLOR_CARD,
            dropdown_hover_color=COLOR_SECONDARY_BG,
            text_color=COLOR_TEXT
        )
        combo.pack(pady=(0, 10))
        combo.set(current)

        self.body_label(
            card,
            "목록에 프린터가 안 보이면 정확한 프린터 이름을 위 칸에 직접 입력해도 됩니다.",
            size=15, muted=True
        ).pack(pady=(0, 6))

        if error_message:
            ctk.CTkLabel(
                card,
                text=f"프린터 목록을 불러오지 못했습니다:\n{error_message}",
                font=(self.font_family, 14),
                text_color=COLOR_DANGER,
                justify="center"
            ).pack(pady=(6, 10))
        elif not printers:
            ctk.CTkLabel(
                card,
                text="설치된 프린터를 찾지 못했습니다. Windows에 프린터가 정상 설치되어 있는지 확인하세요.",
                font=(self.font_family, 14),
                text_color=COLOR_DANGER
            ).pack(pady=(6, 10))

        def refresh_list():
            self.printer_setting_screen()

        self.link_button(card, "목록 새로고침", refresh_list, width=160, height=42).pack(pady=(0, 16))

        def save_printer():
            selected = combo.get().strip()

            if not selected:
                messagebox.showwarning("선택 오류", "프린터를 선택하세요.")
                return

            set_setting("selected_printer", selected)
            messagebox.showinfo("완료", f"프린터를 저장했습니다.\n{selected}")
            self.printer_setting_screen()

        def test_print():
            selected = combo.get().strip()

            if not selected:
                messagebox.showwarning("선택 오류", "프린터를 선택하세요.")
                return

            file_path = filedialog.askopenfilename(
                title="테스트 출력할 파일 선택",
                filetypes=[
                    ("지원 파일", "*.pdf *.doc *.docx *.hwp *.hwpx *.jpg *.jpeg *.png *.ppt *.pptx *.csv *.xls *.xlsx *.txt"),
                    ("모든 파일", "*.*")
                ]
            )

            if not file_path:
                return

            try:
                print_document(file_path, selected)
                admin_log("테스트 출력", f"{os.path.basename(file_path)} → {selected}")
                messagebox.showinfo("전송 완료", "테스트 출력 명령을 보냈습니다.")
            except Exception as e:
                messagebox.showerror("오류", str(e))

        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.pack(pady=(0, 8))

        self.primary_button(btn_frame, "프린터 저장", save_printer, width=180, height=48, font_size=16).grid(row=0, column=0, padx=8)
        self.secondary_button(btn_frame, "테스트 출력", test_print, width=180, height=48, font_size=16).grid(row=0, column=1, padx=8)

        self.link_button(card, "관리자 메뉴로", self.admin_menu, width=180, height=42).pack(pady=(10, 0))

    def log_manage_screen(self):
        card = self.build_card()
        self.make_title(card, "출력 기록 관리")

        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.pack(pady=(0, 16))

        self.secondary_button(btn_frame, "오늘 출력 기록 보기", self.show_today_logs, width=180, height=44, font_size=16).grid(row=0, column=0, padx=6, pady=6)
        self.secondary_button(btn_frame, "사용자별 출력 횟수 보기", self.show_user_counts, width=180, height=44, font_size=16).grid(row=0, column=1, padx=6, pady=6)
        self.secondary_button(btn_frame, "특정 사용자 기록 확인", self.show_specific_user_logs, width=180, height=44, font_size=16).grid(row=0, column=2, padx=6, pady=6)
        self.primary_button(btn_frame, "CSV로 내보내기", self.log_export_screen, width=180, height=44, font_size=16).grid(row=0, column=3, padx=6, pady=6)
        self.danger_button(btn_frame, "전체 기록 초기화", self.clear_all_logs, width=180, height=44, font_size=16).grid(row=0, column=4, padx=6, pady=6)

        self.log_tree = ttk.Treeview(
            card,
            columns=("c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8"),
            show="headings",
            height=15,
            style="App.Treeview"
        )

        headings = ["학생증 코드", "학번", "이름", "파일명/횟수", "용지", "날짜", "시간", "결과"]
        widths = [140, 110, 130, 300, 110, 160, 120, 100]

        for i, heading in enumerate(headings, start=1):
            self.log_tree.heading(f"c{i}", text=heading)
            self.log_tree.column(f"c{i}", width=widths[i - 1], minwidth=widths[i - 1], stretch=False)

        self.log_tree.pack(pady=(0, 16))

        self.link_button(card, "관리자 메뉴로", self.admin_menu, width=180, height=42).pack()

    def log_export_screen(self):
        card = self.build_card()
        self.make_title(card, "출력 기록 CSV 내보내기")

        # ---- 기간 ----
        period_frame = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=16)
        period_frame.pack(pady=(0, 18), ipadx=20, ipady=18)

        self.body_label(period_frame, "기간", size=17, bold=True).grid(row=0, column=0, columnspan=2, pady=(0, 12))

        self.body_label(period_frame, "시작일", size=15, muted=True).grid(row=1, column=0, padx=12, pady=8, sticky="e")
        start_entry = self.make_entry(period_frame, width=220, height=44, font_size=15, justify="left")
        start_entry.grid(row=1, column=1, padx=12, pady=8)

        self.body_label(period_frame, "종료일", size=15, muted=True).grid(row=2, column=0, padx=12, pady=8, sticky="e")
        end_entry = self.make_entry(period_frame, width=220, height=44, font_size=15, justify="left")
        end_entry.grid(row=2, column=1, padx=12, pady=8)

        self.body_label(
            period_frame,
            "예: 2026-03-01 (비워두면 전체 기간)",
            size=13, muted=True
        ).grid(row=3, column=0, columnspan=2, pady=(6, 0))

        # 빠른 기간 선택
        quick_frame = ctk.CTkFrame(card, fg_color="transparent")
        quick_frame.pack(pady=(0, 18))

        def set_period(days):
            end = date.today()
            start = end if days == 0 else end - timedelta(days=days - 1)

            start_entry.delete(0, tk.END)
            start_entry.insert(0, str(start))
            end_entry.delete(0, tk.END)
            end_entry.insert(0, str(end))

        def clear_period():
            start_entry.delete(0, tk.END)
            end_entry.delete(0, tk.END)

        self.secondary_button(quick_frame, "오늘", lambda: set_period(0), width=110, height=42).grid(row=0, column=0, padx=5)
        self.secondary_button(quick_frame, "최근 7일", lambda: set_period(7), width=110, height=42).grid(row=0, column=1, padx=5)
        self.secondary_button(quick_frame, "최근 30일", lambda: set_period(30), width=110, height=42).grid(row=0, column=2, padx=5)
        self.secondary_button(quick_frame, "전체 기간", clear_period, width=110, height=42).grid(row=0, column=3, padx=5)

        # ---- 대상 ----
        target_frame = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=16)
        target_frame.pack(pady=(0, 18), ipadx=20, ipady=18)

        self.body_label(target_frame, "특정 사용자만 (비워두면 전체)", size=15, muted=True).grid(row=0, column=0, padx=12, pady=8, sticky="e")
        user_entry = self.make_entry(target_frame, width=220, height=44, font_size=15, justify="left")
        user_entry.grid(row=0, column=1, padx=12, pady=8)

        self.body_label(
            target_frame,
            "학생증 코드 또는 학번을 입력하세요.",
            size=13, muted=True
        ).grid(row=1, column=0, columnspan=2, pady=(6, 0))

        def collect_conditions(need_user=False):
            """입력값을 검증해서 (시작일, 종료일, 학생증코드)를 돌려준다."""
            try:
                start_date = _normalize_date_input(start_entry.get())
                end_date = _normalize_date_input(end_entry.get())
            except ValueError as e:
                messagebox.showwarning("입력 오류", str(e))
                return None

            if start_date and end_date and start_date > end_date:
                messagebox.showwarning("입력 오류", "시작일이 종료일보다 뒤입니다.")
                return None

            card_code = None
            raw_user = user_entry.get().strip()

            if raw_user:
                # 학번으로 입력해도 찾을 수 있게 get_student를 거친다
                student = get_student(raw_user)

                if not student:
                    messagebox.showerror("오류", "해당하는 사용자를 찾을 수 없습니다.")
                    return None

                card_code = student[0]
            elif need_user:
                messagebox.showwarning("입력 오류", "사용자를 입력하세요.")
                return None

            return start_date, end_date, card_code

        def build_filename(prefix, start_date, end_date):
            if start_date and end_date:
                period = f"_{start_date}_{end_date}"
            elif start_date:
                period = f"_{start_date}이후"
            elif end_date:
                period = f"_{end_date}까지"
            else:
                period = "_전체기간"

            return f"{prefix}{period}.csv"

        def ask_save_path(default_name):
            return filedialog.asksaveasfilename(
                title="CSV 저장 위치 선택",
                defaultextension=".csv",
                initialfile=default_name,
                filetypes=[("CSV 파일", "*.csv")]
            )

        def export_detail(need_user=False):
            conditions = collect_conditions(need_user=need_user)
            if conditions is None:
                return

            start_date, end_date, card_code = conditions

            prefix = f"출력기록_{card_code}" if card_code else "출력기록_전체"
            path = ask_save_path(build_filename(prefix, start_date, end_date))

            if not path:
                return

            try:
                count = export_logs_detail_to_csv(path, start_date, end_date, card_code)

                if count == 0:
                    messagebox.showinfo("완료", "조건에 맞는 기록이 없어 빈 파일이 저장되었습니다.")
                else:
                    admin_log("출력 기록 내보내기", f"{count}건")
                    messagebox.showinfo("완료", f"출력 기록 {count}건을 저장했습니다.")

            except Exception as e:
                messagebox.showerror("내보내기 오류", str(e))

        def export_summary():
            conditions = collect_conditions()
            if conditions is None:
                return

            start_date, end_date, card_code = conditions

            if card_code:
                messagebox.showinfo(
                    "안내",
                    "사용자별 집계는 전체 사용자를 대상으로 합니다.\n"
                    "사용자 칸을 비우고 다시 시도하세요."
                )
                return

            path = ask_save_path(build_filename("사용자별집계", start_date, end_date))

            if not path:
                return

            try:
                count = export_logs_summary_to_csv(path, start_date, end_date)

                if count == 0:
                    messagebox.showinfo("완료", "조건에 맞는 기록이 없어 빈 파일이 저장되었습니다.")
                else:
                    admin_log("사용자별 집계 내보내기", f"{count}명")
                    messagebox.showinfo("완료", f"사용자 {count}명의 집계를 저장했습니다.")

            except Exception as e:
                messagebox.showerror("내보내기 오류", str(e))

        export_frame = ctk.CTkFrame(card, fg_color="transparent")
        export_frame.pack(pady=(0, 16))

        self.primary_button(
            export_frame, "건별 기록 내보내기",
            lambda: export_detail(False), width=210, height=50, font_size=15
        ).grid(row=0, column=0, padx=6)

        self.secondary_button(
            export_frame, "특정 사용자만 내보내기",
            lambda: export_detail(True), width=210, height=50, font_size=15
        ).grid(row=0, column=1, padx=6)

        self.secondary_button(
            export_frame, "사용자별 집계 내보내기",
            export_summary, width=210, height=50, font_size=15
        ).grid(row=0, column=2, padx=6)

        self.body_label(
            card,
            "건별 기록은 출력 한 건마다 한 줄로, 집계는 사용자 한 명마다 한 줄로 저장됩니다.",
            size=14, muted=True
        ).pack(pady=(0, 20))

        self.link_button(card, "출력 기록 관리로", self.log_manage_screen, width=200, height=42).pack()

    def clear_log_tree(self):
        for row in self.log_tree.get_children():
            self.log_tree.delete(row)

    def show_today_logs(self):
        self.clear_log_tree()

        conn = connect_db()
        cur = conn.cursor()
        cur.execute("""
        SELECT p.student_id, COALESCE(s.student_number, ''), p.student_name,
               p.file_name, p.sheet_count, p.duplex, p.print_date, p.print_time, p.result
        FROM print_logs p
        LEFT JOIN students s ON s.student_id = p.student_id
        WHERE p.print_date=?
        ORDER BY p.print_time DESC
        """, (str(date.today()),))
        rows = cur.fetchall()
        conn.close()

        for card_code, number, name, file_name, sheets, duplex, d, t, result in rows:
            sheet_text = f"{sheets}장 ({'양면' if duplex else '단면'})"
            self.log_tree.insert(
                "", "end",
                values=(card_code, number, name, file_name, sheet_text, d, t, result)
            )

    def show_user_counts(self):
        self.clear_log_tree()

        conn = connect_db()
        cur = conn.cursor()
        cur.execute("""
        SELECT s.student_id, COALESCE(s.student_number, ''), s.name,
               COUNT(p.id) AS cnt, COALESCE(SUM(p.sheet_count), 0) AS sheets
        FROM students s
        LEFT JOIN print_logs p
        ON s.student_id = p.student_id
        AND p.print_date = ?
        AND p.result != '실패'
        GROUP BY s.student_id, s.student_number, s.name
        ORDER BY sheets DESC, cnt DESC
        """, (str(date.today()),))
        rows = cur.fetchall()
        conn.close()

        for card_code, student_number, name, cnt, sheets in rows:
            self.log_tree.insert(
                "",
                "end",
                values=(card_code, student_number, name, f"오늘 {cnt}회",
                        f"{sheets}장", str(date.today()), "-", "-")
            )

    def ask_text(self, title, prompt):
        """
        전체화면 topmost 상태에서는 tkinter 기본 입력창(simpledialog)이
        메인 창 뒤에 숨어버려서 프로그램이 멈춘 것처럼 보인다.
        (창은 떠 있지만 보이지 않고, 모달 상태라 조작도 안 된다)
        그래서 항상 위로 뜨는 자체 입력창을 쓴다.
        """
        dialog = ctk.CTkToplevel(self.root)
        dialog.title(title)
        dialog.after(250, lambda: apply_window_icon(dialog))
        dialog.geometry("420x220")
        dialog.configure(fg_color=COLOR_BG)
        dialog.resizable(False, False)

        dialog.attributes("-topmost", True)
        dialog.transient(self.root)
        dialog.lift()

        result = {"value": None}

        ctk.CTkLabel(
            dialog, text=prompt,
            font=(self.font_family, 16), text_color=COLOR_TEXT
        ).pack(pady=(28, 14))

        entry = self.make_entry(dialog, width=300, height=46, font_size=16, justify="left")
        entry.pack(pady=(0, 20))

        def submit():
            result["value"] = entry.get().strip()
            dialog.destroy()

        def cancel():
            result["value"] = None
            dialog.destroy()

        entry.bind("<Return>", lambda e: submit())

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack()

        self.primary_button(btn_frame, "확인", submit, width=120, height=44).grid(row=0, column=0, padx=6)
        self.secondary_button(btn_frame, "취소", cancel, width=120, height=44).grid(row=0, column=1, padx=6)

        dialog.protocol("WM_DELETE_WINDOW", cancel)

        dialog.after(120, lambda: (dialog.grab_set(), entry.focus_force()))
        self.root.wait_window(dialog)

        return result["value"]

    def show_specific_user_logs(self):
        raw_code = self.ask_text("특정 사용자 기록", "학생증 코드를 입력하세요.")

        if not raw_code:
            return

        card_code = normalize_card_code(raw_code)

        # 입력값이 학번이어도 찾을 수 있게 학생 조회를 거친다
        student = get_student(card_code)
        if student:
            card_code = student[0]

        self.clear_log_tree()

        conn = connect_db()
        cur = conn.cursor()
        cur.execute("""
        SELECT p.student_id, COALESCE(s.student_number, ''), p.student_name,
               p.file_name, p.sheet_count, p.duplex, p.print_date, p.print_time, p.result
        FROM print_logs p
        LEFT JOIN students s ON s.student_id = p.student_id
        WHERE p.student_id=?
        ORDER BY p.print_date DESC, p.print_time DESC
        """, (card_code,))
        rows = cur.fetchall()
        conn.close()

        if not rows:
            messagebox.showinfo("조회 결과", "해당 사용자의 출력 기록이 없습니다.")
            return

        for c, number, nm, file_name, sheets, duplex, d, t, result in rows:
            sheet_text = f"{sheets}장 ({'양면' if duplex else '단면'})"
            self.log_tree.insert(
                "", "end",
                values=(c, number, nm, file_name, sheet_text, d, t, result)
            )

    def clear_all_logs(self):
        result = messagebox.askyesno(
            "전체 기록 초기화",
            "정말 전체 출력 기록을 초기화하시겠습니까?\n이 작업은 되돌릴 수 없습니다."
        )

        if not result:
            return

        conn = connect_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM print_logs")
        conn.commit()
        conn.close()

        self.clear_log_tree()
        admin_log("출력 기록 전체 삭제")
        messagebox.showinfo("완료", "전체 출력 기록을 초기화했습니다.")

    def change_admin_password_screen(self):
        card = self.build_card()
        self.make_title(card, "비밀번호 설정")

        # ── 비밀번호 변경 ──
        frame = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=14)
        frame.pack(pady=(0, 16), ipadx=16, ipady=16)

        # grid 로 채운 상자는 ipadx/ipady 여백이 오른쪽·아래쪽에만 남는다.
        # center 로 두면 여백이 사방으로 나뉘어 내용이 가운데 온다.
        frame.grid_anchor("center")

        self.body_label(frame, "현재 비밀번호", size=17).grid(row=0, column=0, padx=10, pady=10)
        old_entry = self.make_entry(frame, width=220, height=46, font_size=16, show="*")
        old_entry.grid(row=0, column=1, padx=10, pady=10)

        self.body_label(frame, "새 비밀번호", size=17).grid(row=1, column=0, padx=10, pady=10)
        new_entry = self.make_entry(frame, width=220, height=46, font_size=16, show="*")
        new_entry.grid(row=1, column=1, padx=10, pady=10)

        self.body_label(frame, "새 비밀번호 확인", size=17).grid(row=2, column=0, padx=10, pady=10)
        confirm_entry = self.make_entry(frame, width=220, height=46, font_size=16, show="*")
        confirm_entry.grid(row=2, column=1, padx=10, pady=10)

        def save_password():
            old_pw = old_entry.get()
            new_pw = new_entry.get()
            confirm_pw = confirm_entry.get()

            real_pw = get_setting("admin_password", DEFAULT_ADMIN_PASSWORD)

            if not verify_password(old_pw, real_pw):
                messagebox.showerror("오류", "현재 비밀번호가 틀렸습니다.")
                return

            if not new_pw:
                messagebox.showwarning("입력 오류", "새 비밀번호를 입력하세요.")
                return

            if new_pw != confirm_pw:
                messagebox.showerror("오류", "새 비밀번호가 서로 다릅니다.")
                return

            # 비밀번호는 그대로 저장하지 않고 해시로 바꿔 저장한다.
            # (DB 파일을 열어봐도 비밀번호를 알 수 없다)
            set_setting("admin_password", hash_password(new_pw))
            messagebox.showinfo("완료", "관리자 비밀번호가 변경되었습니다.")
            self.admin_menu()

        self.primary_button(card, "비밀번호 변경", save_password, width=220, height=48).pack(pady=(0, 24))

        divider = ctk.CTkFrame(card, fg_color=COLOR_BORDER, height=1)
        divider.pack(fill="x", pady=(0, 24))

        # ── 관리자 바코드 코드 설정 ──
        self.body_label(card, "관리자 모드 접근 코드", size=20, bold=True).pack(pady=(0, 8))

        self.body_label(
            card,
            "학생증 바코드 입력란에 이 코드를 입력하면 바로 관리자 화면으로 접속됩니다.\n"
            "비워두면 이 기능이 비활성화됩니다.",
            size=15, muted=True
        ).pack(pady=(0, 14))

        barcode_frame = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=14)
        barcode_frame.pack(pady=(0, 16), ipadx=16, ipady=16)
        barcode_frame.grid_anchor("center")

        current_barcode = get_setting("admin_barcode", "")

        self.body_label(barcode_frame, "관리자 코드", size=17).grid(row=0, column=0, padx=12, pady=10)
        barcode_entry = self.make_entry(barcode_frame, width=260, height=46, font_size=16, justify="left")
        barcode_entry.grid(row=0, column=1, padx=12, pady=10)
        if current_barcode:
            barcode_entry.insert(0, current_barcode)

        def save_barcode():
            new_code = normalize_card_code(barcode_entry.get())

            # 등록된 학생증 코드나 학번과 겹치는지 확인
            if new_code:
                student = get_student(new_code)
                if student:
                    messagebox.showerror(
                        "코드 충돌",
                        "입력한 코드가 이미 등록된 학생증 코드 또는 학번과 일치합니다.\n"
                        "다른 코드를 사용하세요."
                    )
                    return

            set_setting("admin_barcode", new_code)
            if new_code:
                messagebox.showinfo("완료", f"관리자 코드를 저장했습니다.\n\n코드: {new_code}")
            else:
                messagebox.showinfo("완료", "관리자 코드를 삭제했습니다.")
            self.admin_menu()

        self.primary_button(card, "코드 저장", save_barcode, width=220, height=48).pack(pady=(0, 12))
        self.link_button(card, "관리자 메뉴로", self.admin_menu).pack()

    def admin_account_screen(self):
        """
        관리자 방식 선택 · 관리자 명단 · 작업 기록으로 가는 화면.

        단일 방식  : 메인 비밀번호 하나로 들어오고, 기록에는 '관리자' 로 남는다.
        개별 방식  : admin_list.csv 에 적힌 사람마다 비밀번호와 관리자코드를 따로 두고,
                     기록에는 그 사람 이름으로 남는다.
        """
        card = self.build_card()
        self.make_title(card, "관리자 계정 · 작업 기록")

        # 이 화면을 벗어나면 명단 자동 새로고침을 끊는다
        self._admin_list_refresh = None

        columns = ctk.CTkFrame(card, fg_color="transparent")
        columns.pack(pady=(0, 10))

        left = ctk.CTkFrame(columns, fg_color="transparent")
        left.grid(row=0, column=0, padx=(0, 24), sticky="n")

        ctk.CTkFrame(columns, fg_color=COLOR_BORDER, width=1).grid(row=0, column=1, sticky="ns")

        right = ctk.CTkFrame(columns, fg_color="transparent")
        right.grid(row=0, column=2, padx=(24, 0), sticky="n")

        # ── 왼쪽: 관리자 방식 ──
        self.body_label(left, "관리자 방식", size=19, bold=True).pack(pady=(0, 10))

        mode_var = tk.StringVar(value=get_admin_mode())

        for value, text in (("single", "단일 비밀번호 + 기록"), ("individual", "관리자별 비밀번호 + 개인별 기록")):
            ctk.CTkRadioButton(
                left, text=text, variable=mode_var, value=value,
                font=(self.font_family, 15), text_color=COLOR_TEXT,
                fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
                command=lambda: refresh_mode_hint()
            ).pack(anchor="w", pady=4)

        mode_hint = self.body_label(left, "", size=13, muted=True, wraplength=330, justify="left")
        mode_hint.pack(pady=(8, 12), anchor="w")

        def refresh_mode_hint():
            if mode_var.get() == "individual":
                mode_hint.configure(text=(
                    "관리자 명단(admin_list.csv)에 적힌 사람만 들어올 수 있습니다.\n"
                    "각자 자기 비밀번호나 관리자코드(카드 QR)로 들어오고,\n"
                    "기록에는 그 사람 이름으로 남습니다.\n"
                    "명단이 비어 있을 때만 메인 비밀번호가 통합니다."
                ))
            else:
                mode_hint.configure(text=(
                    "메인 비밀번호 하나와 메인 관리자코드로 들어옵니다.\n"
                    "작업은 모두 기록되지만 이름 대신 '관리자' 로 남습니다."
                ))

        refresh_mode_hint()

        def save_mode():
            mode = mode_var.get()

            if mode == "individual":
                sync_admin_csv(force=True)

                if not list_admins():
                    if not messagebox.askyesno(
                        "확인",
                        "관리자 명단이 비어 있습니다.\n\n"
                        "명단을 채우기 전까지는 메인 비밀번호로만 들어올 수 있습니다.\n"
                        "그래도 관리자별 방식으로 바꿀까요?"
                    ):
                        return

            set_setting("admin_mode", mode)
            messagebox.showinfo(
                "완료",
                "관리자 방식을 바꿨습니다.\n\n" +
                ("관리자별 비밀번호 + 개인별 기록" if mode == "individual" else "단일 비밀번호 + 기록")
            )
            self.admin_account_screen()

        self.primary_button(left, "방식 저장", save_mode, width=180, height=46).pack(anchor="w")

        # ── 오른쪽: 관리자 명단 ──
        self.body_label(right, "관리자 명단", size=19, bold=True).pack(pady=(0, 6))
        self.body_label(
            right,
            f"프로그램 폴더의 {ADMIN_CSV_NAME} 에 이름 · 학번 · 직함 · 관리자코드 · 비밀번호를 적고 저장하면\n"
            "몇 초 안에 반영됩니다. 적은 비밀번호와 코드는 곧바로 '(등록됨)' 으로 가려집니다.",
            size=13, muted=True
        ).pack(pady=(0, 8))

        tree = ttk.Treeview(
            right, columns=("number", "name", "title", "code", "pw", "note"),
            show="headings", height=6, style="App.Treeview"
        )
        tree.heading("number", text="학번")
        tree.heading("name", text="이름")
        tree.heading("title", text="직함")
        tree.heading("code", text="관리자코드")
        tree.heading("pw", text="비밀번호")
        tree.heading("note", text="비고")
        tree.column("number", width=70, anchor="center", stretch=False)
        tree.column("name", width=90, stretch=False)
        tree.column("title", width=100, stretch=False)
        tree.column("code", width=95, anchor="center", stretch=False)
        tree.column("pw", width=85, anchor="center", stretch=False)
        tree.column("note", width=120, stretch=False)
        tree.pack(pady=(0, 6))

        problem_label = self.body_label(right, "", size=13, wraplength=500, justify="left")
        problem_label.configure(text_color=COLOR_DANGER)
        problem_label.pack(pady=(0, 6))

        def refresh_admins():
            for row in tree.get_children():
                tree.delete(row)

            profiles = list_admin_profiles()

            for name, pw_hash, code_hash, note in list_admins():
                number, title = profiles.get(name, ("", ""))
                tree.insert("", "end", values=(
                    number,
                    name,
                    title,
                    "등록됨" if code_hash else "-",
                    "등록됨" if pw_hash else "-",
                    note
                ))

            problems = _admin_csv_state.get("problems") or []
            problem_label.configure(text="\n".join("· " + p for p in problems[:4]))

        self._admin_list_refresh = refresh_admins
        refresh_admins()

        def open_admin_csv():
            path = ensure_admin_csv()
            try:
                os.startfile(path)
            except Exception:
                messagebox.showinfo("안내", f"명단 파일 위치\n\n{path}")

        def reload_admin_csv():
            sync_admin_csv(force=True)
            refresh_admins()

        list_buttons = ctk.CTkFrame(right, fg_color="transparent")
        list_buttons.pack()
        self.secondary_button(list_buttons, "명단 파일 열기", open_admin_csv, width=160, height=44).grid(row=0, column=0, padx=5)
        self.secondary_button(list_buttons, "다시 읽기", reload_admin_csv, width=120, height=44).grid(row=0, column=1, padx=5)

        # ── 아래: 작업 기록 ──
        bottom = ctk.CTkFrame(card, fg_color="transparent")
        bottom.pack(pady=(14, 0))

        def leave(target):
            self._admin_list_refresh = None
            target()

        self.primary_button(bottom, "작업 기록 보기", lambda: leave(self.admin_log_screen), width=200, height=48).grid(row=0, column=0, padx=8)
        self.link_button(bottom, "관리자 메뉴로", lambda: leave(self.admin_menu), width=180).grid(row=0, column=1, padx=8)

    def admin_log_screen(self):
        """누가 · 언제 · 무엇을 했는지 보는 화면."""
        card = self.build_card()
        self.make_title(card, "관리자 작업 기록")

        filter_frame = ctk.CTkFrame(card, fg_color="transparent")
        filter_frame.pack(pady=(0, 10))

        self.body_label(filter_frame, "관리자", size=16, bold=True).grid(row=0, column=0, padx=(0, 10))

        actors = ["전체"] + list_admin_log_actors()
        actor_var = tk.StringVar(value="전체")

        ctk.CTkOptionMenu(
            filter_frame, values=actors, variable=actor_var,
            width=220, height=40, font=(self.font_family, 15),
            fg_color=COLOR_SECONDARY_BG, button_color=COLOR_PRIMARY,
            button_hover_color=COLOR_PRIMARY_HOVER, text_color=COLOR_TEXT,
            command=lambda _: refresh()
        ).grid(row=0, column=1)

        tree = ttk.Treeview(
            card, columns=("date", "time", "actor", "action", "detail"),
            show="headings", height=14, style="App.Treeview"
        )
        tree.heading("date", text="날짜")
        tree.heading("time", text="시각")
        tree.heading("actor", text="관리자")
        tree.heading("action", text="작업")
        tree.heading("detail", text="내용")
        tree.column("date", width=110, anchor="center", stretch=False)
        tree.column("time", width=90, anchor="center", stretch=False)
        tree.column("actor", width=150, stretch=False)
        tree.column("action", width=170, stretch=False)
        tree.column("detail", width=440, stretch=False)
        tree.pack(pady=(0, 6))

        count_label = self.body_label(card, "", size=13, muted=True)
        count_label.pack(pady=(0, 10))

        def selected_actor():
            value = actor_var.get()
            return None if value == "전체" else value

        def refresh():
            for row in tree.get_children():
                tree.delete(row)

            rows = fetch_admin_logs(actor=selected_actor(), limit=500)
            for row in rows:
                tree.insert("", "end", values=row)

            count_label.configure(text=f"최근 {len(rows)}건 표시 (전체는 CSV 로 내보내기)")

        refresh()

        def export():
            path = filedialog.asksaveasfilename(
                title="관리자 작업 기록 내보내기",
                defaultextension=".csv",
                initialfile=f"관리자기록_{date.today()}.csv",
                filetypes=[("CSV 파일", "*.csv")]
            )

            if not path:
                return

            try:
                count = export_admin_logs_to_csv(path, actor=selected_actor())
                admin_log("작업 기록 내보내기", f"{count}건 · {selected_actor() or '전체'}")
                messagebox.showinfo("완료", f"기록 {count}건을 저장했습니다.")
                refresh()
            except Exception as e:
                messagebox.showerror("내보내기 오류", str(e))

        buttons = ctk.CTkFrame(card, fg_color="transparent")
        buttons.pack()
        self.primary_button(buttons, "CSV 내보내기", export, width=180, height=46).grid(row=0, column=0, padx=8)
        self.link_button(buttons, "뒤로", self.admin_account_screen, width=140).grid(row=0, column=1, padx=8)

    def shutdown_schedule_screen(self):
        card = self.build_card()
        self.make_title(card, "자동 종료 일정 관리")

        # 지나간 1회 일정은 자동 정리
        cleanup_past_schedules()

        # ── 등록된 일정 목록 ──
        tree = ttk.Treeview(
            card,
            columns=("id", "kind", "when", "time"),
            show="headings",
            height=8,
            style="App.Treeview"
        )

        tree.heading("id", text="번호")
        tree.heading("kind", text="종류")
        tree.heading("when", text="요일 / 날짜")
        tree.heading("time", text="시각")

        tree.column("id", width=70, anchor="center", stretch=False)
        tree.column("kind", width=140, anchor="center", stretch=False)
        tree.column("when", width=320, stretch=False)
        tree.column("time", width=140, anchor="center", stretch=False)

        tree.pack(pady=(0, 10))

        def refresh_list():
            for row in tree.get_children():
                tree.delete(row)

            for schedule_id, kind, days, run_date, run_time in list_shutdown_schedules():
                when_text, time_text = describe_schedule(kind, days, run_date, run_time)
                kind_text = "매주 반복" if kind == "weekly" else "특정 날짜"
                tree.insert(
                    "", "end",
                    values=(schedule_id, kind_text, when_text, time_text)
                )

        # ── 새 일정 추가 ──
        divider = ctk.CTkFrame(card, fg_color=COLOR_BORDER, height=1)
        divider.pack(fill="x", pady=(6, 18))

        self.body_label(card, "새 일정 추가", size=17, bold=True).pack(pady=(0, 12))

        kind_var = tk.StringVar(value="weekly")

        kind_frame = ctk.CTkFrame(card, fg_color="transparent")
        kind_frame.pack(pady=(0, 14))

        add_frame = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=14)
        add_frame.pack(pady=(0, 14), ipadx=16, ipady=16)

        # 요일 선택 영역
        weekly_frame = ctk.CTkFrame(add_frame, fg_color="transparent")

        self.body_label(weekly_frame, "반복할 요일", size=15).pack(pady=(0, 8))

        day_row = ctk.CTkFrame(weekly_frame, fg_color="transparent")
        day_row.pack(pady=(0, 10))

        day_vars = {}

        for i, (code, label) in enumerate(zip(WEEKDAY_CODES, WEEKDAY_LABELS)):
            var = tk.IntVar(value=1 if i < 5 else 0)
            day_vars[code] = var

            ctk.CTkCheckBox(
                day_row,
                text=label,
                variable=var,
                onvalue=1,
                offvalue=0,
                width=58,
                font=(self.font_family, 15),
                text_color=COLOR_TEXT,
                fg_color=COLOR_PRIMARY,
                hover_color=COLOR_PRIMARY_HOVER
            ).grid(row=0, column=i, padx=5)

        quick_row = ctk.CTkFrame(weekly_frame, fg_color="transparent")
        quick_row.pack()

        def select_days(codes):
            for code, var in day_vars.items():
                var.set(1 if code in codes else 0)

        self.secondary_button(quick_row, "매일", lambda: select_days(WEEKDAY_CODES), width=100, height=38, font_size=14).grid(row=0, column=0, padx=4)
        self.secondary_button(quick_row, "평일", lambda: select_days(WEEKDAY_CODES[:5]), width=100, height=38, font_size=14).grid(row=0, column=1, padx=4)
        self.secondary_button(quick_row, "주말", lambda: select_days(WEEKDAY_CODES[5:]), width=100, height=38, font_size=14).grid(row=0, column=2, padx=4)
        self.secondary_button(quick_row, "해제", lambda: select_days([]), width=100, height=38, font_size=14).grid(row=0, column=3, padx=4)

        # 특정 날짜 선택 영역
        once_frame = ctk.CTkFrame(add_frame, fg_color="transparent")

        date_row = ctk.CTkFrame(once_frame, fg_color="transparent")
        date_row.pack(pady=(0, 10))

        self.body_label(date_row, "날짜", size=15).grid(row=0, column=0, padx=10)
        date_entry = self.make_entry(date_row, width=200, height=46, font_size=16)
        date_entry.grid(row=0, column=1, padx=10)
        date_entry.insert(0, str(date.today()))

        self.body_label(
            date_row, "YYYY-MM-DD", size=13, muted=True
        ).grid(row=0, column=2, padx=8)

        date_quick = ctk.CTkFrame(once_frame, fg_color="transparent")
        date_quick.pack()

        def set_date(days_later):
            target = date.today() + timedelta(days=days_later)
            date_entry.delete(0, tk.END)
            date_entry.insert(0, str(target))

        self.secondary_button(date_quick, "오늘", lambda: set_date(0), width=100, height=38, font_size=14).grid(row=0, column=0, padx=4)
        self.secondary_button(date_quick, "내일", lambda: set_date(1), width=100, height=38, font_size=14).grid(row=0, column=1, padx=4)
        self.secondary_button(date_quick, "일주일 뒤", lambda: set_date(7), width=110, height=38, font_size=14).grid(row=0, column=2, padx=4)

        # 종류 전환
        def switch_kind():
            if kind_var.get() == "weekly":
                once_frame.pack_forget()
                weekly_frame.pack(pady=(0, 4))
            else:
                weekly_frame.pack_forget()
                once_frame.pack(pady=(0, 4))

        ctk.CTkRadioButton(
            kind_frame, text="매주 반복", variable=kind_var, value="weekly",
            command=switch_kind,
            font=(self.font_family, 16), text_color=COLOR_TEXT,
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER
        ).grid(row=0, column=0, padx=16)

        ctk.CTkRadioButton(
            kind_frame, text="특정 날짜", variable=kind_var, value="once",
            command=switch_kind,
            font=(self.font_family, 16), text_color=COLOR_TEXT,
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER
        ).grid(row=0, column=1, padx=16)

        switch_kind()

        # 시각 입력 (공통)
        time_row = ctk.CTkFrame(add_frame, fg_color="transparent")
        time_row.pack(pady=(14, 0))

        self.body_label(time_row, "종료 시각", size=15).grid(row=0, column=0, padx=10)
        time_entry = self.make_entry(time_row, width=160, height=46, font_size=16)
        time_entry.grid(row=0, column=1, padx=10)
        time_entry.insert(0, "18:00")

        self.body_label(
            time_row, "24시간 형식 (예: 18:00)", size=13, muted=True
        ).grid(row=0, column=2, padx=8)

        def parse_time(text):
            match = re.fullmatch(r"(\d{1,2}):(\d{2})", text.strip())
            if not match:
                return None

            hour, minute = int(match.group(1)), int(match.group(2))
            if hour > 23 or minute > 59:
                return None

            return f"{hour:02d}:{minute:02d}"

        def add_schedule():
            run_time = parse_time(time_entry.get())

            if not run_time:
                messagebox.showwarning("입력 오류", "시각을 HH:MM 형식으로 입력하세요.\n예: 18:00")
                return

            kind = kind_var.get()

            if kind == "weekly":
                selected = [code for code, var in day_vars.items() if var.get() == 1]

                if not selected:
                    messagebox.showwarning("입력 오류", "요일을 하나 이상 선택하세요.")
                    return

                label = f"{format_day_text(selected)} {run_time}"
                ok, output = add_shutdown_schedule("weekly", selected, None, run_time)

            else:
                date_text = date_entry.get().strip()

                try:
                    target = datetime.strptime(date_text, "%Y-%m-%d").date()
                except ValueError:
                    messagebox.showwarning("입력 오류", "날짜를 YYYY-MM-DD 형식으로 입력하세요.\n예: 2026-08-15")
                    return

                if target < date.today():
                    messagebox.showwarning("입력 오류", "지난 날짜는 등록할 수 없습니다.")
                    return

                label = f"{target.month}월 {target.day}일 {run_time}"
                ok, output = add_shutdown_schedule("once", None, str(target), run_time)

            if ok:
                admin_log("자동 종료 일정 추가", label.replace("\n", " "))
                messagebox.showinfo("완료", f"일정을 추가했습니다.\n\n{label}")
                refresh_list()
            else:
                messagebox.showerror(
                    "등록 실패",
                    f"작업 스케줄러 등록에 실패했습니다.\n\n{output[:300]}\n\n"
                    "권한 확인 창에서 '예'를 눌렀는지 확인해 주세요."
                )

        self.primary_button(card, "일정 추가", add_schedule, width=200, height=48).pack(pady=(0, 18))

        # ── 삭제 버튼 ──
        def delete_selected():
            selected = tree.selection()

            if not selected:
                messagebox.showwarning("선택 오류", "삭제할 일정을 목록에서 선택하세요.")
                return

            values = tree.item(selected[0], "values")
            schedule_id = int(values[0])

            if not messagebox.askyesno("확인", f"'{values[2]} {values[3]}' 일정을 삭제할까요?"):
                return

            ok, output = delete_shutdown_schedule(schedule_id)

            if ok:
                admin_log("자동 종료 일정 삭제", f"{schedule_id}번")
                messagebox.showinfo("완료", "일정을 삭제했습니다.")
            else:
                messagebox.showwarning(
                    "일정 삭제",
                    "목록에서는 지웠지만 작업 스케줄러에서 지우지 못했습니다.\n\n"
                    f"{output[:300]}"
                )

            refresh_list()

        def delete_all():
            if not list_shutdown_schedules():
                messagebox.showinfo("안내", "등록된 일정이 없습니다.")
                return

            if not messagebox.askyesno("확인", "등록된 자동 종료 일정을 모두 삭제할까요?"):
                return

            delete_all_shutdown_schedules()
            admin_log("자동 종료 일정 전체 삭제")
            messagebox.showinfo("완료", "모든 일정을 삭제했습니다.")
            refresh_list()

        del_frame = ctk.CTkFrame(card, fg_color="transparent")
        del_frame.pack(pady=(0, 18))

        self.danger_button(del_frame, "선택 일정 삭제", delete_selected, width=180, height=46).grid(row=0, column=0, padx=8)
        self.secondary_button(del_frame, "전체 삭제", delete_all, width=180, height=46).grid(row=0, column=1, padx=8)

        self.link_button(card, "보호 기능 설정으로", self.kiosk_setting_screen, width=200, height=42).pack()

        refresh_list()

    def run_update_check(self):
        """
        업데이트 확인 버튼을 눌렀을 때 실행된다.
        네트워크 작업이 화면을 멈추게 하지 않도록 별도 스레드에서 확인한다.
        """
        waiting = ctk.CTkToplevel(self.root)
        waiting.title("업데이트 확인")
        waiting.geometry("420x180")
        waiting.configure(fg_color=COLOR_BG)
        waiting.resizable(False, False)
        waiting.attributes("-topmost", True)
        waiting.transient(self.root)
        waiting.lift()

        # 확인 중에는 닫기 버튼을 막는다
        waiting.protocol("WM_DELETE_WINDOW", lambda: None)

        status = ctk.CTkLabel(
            waiting,
            text="새 버전이 있는지 확인하는 중입니다...",
            font=(self.font_family, 15),
            text_color=COLOR_TEXT,
        )
        status.pack(expand=True, padx=20)

        waiting.update()

        def worker():
            result, info = check_for_update()
            self.root.after(0, lambda: finish(result, info))

        def finish(result, info):
            try:
                waiting.destroy()
            except Exception:
                pass

            if result == "오류":
                messagebox.showerror("업데이트 확인 실패", info)
                return

            if result == "최신":
                messagebox.showinfo(
                    "최신 버전",
                    f"현재 최신 버전을 사용 중입니다.\n\n버전  v{APP_VERSION}"
                )
                return

            # 새 버전 있음
            message = (
                f"새 버전이 있습니다.\n\n"
                f"현재 버전   v{APP_VERSION}\n"
                f"새 버전     v{info['version']}\n"
            )

            if info["notes"]:
                notes = info["notes"]
                if len(notes) > 300:
                    notes = notes[:300] + " ..."
                message += f"\n[변경 내용]\n{notes}\n"

            message += (
                "\n업데이트하시겠습니까?\n"
                "설치가 끝나면 프로그램이 자동으로 다시 시작됩니다."
            )

            if messagebox.askyesno("업데이트", message):
                self.run_update_install(info)

        threading.Thread(target=worker, daemon=True).start()

    def run_update_install(self, info):
        """새 버전을 내려받고 교체한다."""
        progress_win = ctk.CTkToplevel(self.root)
        progress_win.title("업데이트")
        progress_win.geometry("440x200")
        progress_win.configure(fg_color=COLOR_BG)
        progress_win.resizable(False, False)
        progress_win.attributes("-topmost", True)
        progress_win.transient(self.root)
        progress_win.lift()
        progress_win.protocol("WM_DELETE_WINDOW", lambda: None)

        label = ctk.CTkLabel(
            progress_win,
            text="새 버전을 내려받는 중입니다...",
            font=(self.font_family, 15),
            text_color=COLOR_TEXT,
        )
        label.pack(pady=(40, 12), padx=20)

        percent_label = ctk.CTkLabel(
            progress_win,
            text="0%",
            font=(self.font_family, 22, "bold"),
            text_color=COLOR_PRIMARY,
        )
        percent_label.pack()

        progress_win.update()

        def on_progress(percent):
            self.root.after(0, lambda: percent_label.configure(text=f"{percent}%"))

        def worker():
            try:
                path = download_update(info["url"], on_progress)
            except Exception as e:
                self.root.after(0, lambda: fail(f"다운로드에 실패했습니다.\n\n{e}"))
                return

            try:
                admin_log("프로그램 업데이트", f"v{APP_VERSION} → v{info.get('version', '?')}")
                apply_update(path)
            except Exception as e:
                self.root.after(0, lambda: fail(str(e)))
                return

            self.root.after(0, done)

        def fail(message):
            try:
                progress_win.destroy()
            except Exception:
                pass
            messagebox.showerror("업데이트 실패", message)

        def done():
            try:
                progress_win.destroy()
            except Exception:
                pass

            messagebox.showinfo(
                "업데이트",
                "프로그램을 종료한 뒤 새 버전으로 다시 시작합니다.\n"
                "잠시만 기다려 주세요."
            )

            self.close_external_site()
            stop_key_blocker()
            self.root.destroy()

        threading.Thread(target=worker, daemon=True).start()

    def kiosk_setting_screen(self):
        """
        보호 기능 설정 화면.

        항목이 많아서 한 줄로 세우면 FHD(1080p) 화면에서 아래가 잘린다.
        그래서 좌우 두 칸으로 나눠 배치한다.
          왼쪽 : 화면 · Windows 잠금
          오른쪽: 자동 종료 · 프로그램 업데이트
        """
        card = self.build_card()
        self.make_title(card, "보호 기능 설정")

        columns = ctk.CTkFrame(card, fg_color="transparent")
        columns.pack(pady=(0, 4))

        left = ctk.CTkFrame(columns, fg_color="transparent")
        left.grid(row=0, column=0, padx=(0, 26), sticky="n")

        sep = ctk.CTkFrame(columns, fg_color=COLOR_BORDER, width=1)
        sep.grid(row=0, column=1, sticky="ns", pady=6)

        right = ctk.CTkFrame(columns, fg_color="transparent")
        right.grid(row=0, column=2, padx=(26, 0), sticky="n")

        # ==========================================================
        # 왼쪽 — 화면
        # ==========================================================
        self.body_label(left, "화면", size=20, bold=True).pack(pady=(0, 10))

        fullscreen_var = tk.IntVar(value=1 if get_setting("fullscreen", "1") == "1" else 0)

        ctk.CTkCheckBox(
            left,
            text="전체화면 사용 (작업표시줄 가림)",
            variable=fullscreen_var,
            onvalue=1,
            offvalue=0,
            font=(self.font_family, 15),
            text_color=COLOR_TEXT,
            fg_color=COLOR_PRIMARY,
            hover_color=COLOR_PRIMARY_HOVER
        ).pack(pady=(0, 16))

        # ── 왼쪽 — Windows 잠금 ──
        ctk.CTkFrame(left, fg_color=COLOR_BORDER, height=1).pack(fill="x", pady=(0, 18))

        self.body_label(left, "Windows 잠금", size=20, bold=True).pack(pady=(0, 10))

        key_block_var = tk.IntVar(value=1 if get_setting("block_system_keys", "0") == "1" else 0)

        ctk.CTkCheckBox(
            left,
            text="윈도우 키 · Alt+Tab · Ctrl+Shift+Esc 차단",
            variable=key_block_var,
            onvalue=1,
            offvalue=0,
            font=(self.font_family, 15),
            text_color=COLOR_TEXT,
            fg_color=COLOR_PRIMARY,
            hover_color=COLOR_PRIMARY_HOVER
        ).pack(pady=(0, 12))

        self.body_label(
            left,
            "저장 즉시 적용되며 재부팅이 필요 없습니다.\n"
            "프로그램을 끄면 자동으로 원래대로 돌아옵니다.",
            size=14, muted=True
        ).pack(pady=(0, 18))

        lock_state = "잠금 적용됨" if is_windows_lockdown_enabled() else "잠금 해제됨"
        self.body_label(left, f"작업 관리자 · 실행창: {lock_state}", size=15).pack(pady=(0, 8))

        self.body_label(
            left,
            "아래 버튼은 PC 전체 설정(레지스트리)을 바꿉니다.\n"
            "설정 담당자 PC에는 적용하지 마세요.",
            size=14, muted=True
        ).pack(pady=(0, 11))

        def apply_lock(enabled):
            action = "적용" if enabled else "해제"

            if not messagebox.askyesno("확인", f"작업 관리자·실행창 잠금을 {action}하시겠습니까?"):
                return

            done, failed = set_windows_lockdown(enabled)

            admin_log(f"Windows 잠금 {action}", ", ".join(done) + (f" / 실패: {', '.join(failed)}" if failed else ""))

            message = f"잠금 {action} 결과\n\n"
            if done:
                message += "성공: " + ", ".join(done) + "\n"
            if failed:
                message += "\n실패:\n" + "\n".join(failed)

            messagebox.showinfo("결과", message)
            self.kiosk_setting_screen()

        lock_frame = ctk.CTkFrame(left, fg_color="transparent")
        lock_frame.pack(pady=(0, 4))

        self.danger_button(lock_frame, "잠금 적용", lambda: apply_lock(True), width=150, height=46).grid(row=0, column=0, padx=6)
        self.secondary_button(lock_frame, "잠금 해제", lambda: apply_lock(False), width=150, height=46).grid(row=0, column=1, padx=6)

        # ── 왼쪽 — 파일 공유 방화벽 ──
        ctk.CTkFrame(left, fg_color=COLOR_BORDER, height=1).pack(fill="x", pady=(16, 14))

        self.body_label(left, "파일 공유 방화벽", size=20, bold=True).pack(pady=(0, 10))

        firewall_state = "허용됨" if is_firewall_rule_present() else "설정 안 됨"
        self.body_label(
            left, f"{FILE_SHARE_PORT}번 포트: {firewall_state}", size=15
        ).pack(pady=(0, 8))

        self.body_label(
            left,
            "학생 기기가 파일을 보내려면 이 포트를 열어야 합니다.\n"
            "누르면 관리자 권한 확인 창이 뜹니다.",
            size=14, muted=True
        ).pack(pady=(0, 11))

        def apply_firewall(allow):
            if allow:
                message = (
                    f"파일 공유용 {FILE_SHARE_PORT}번 포트를\n"
                    "Windows 방화벽에서 허용하시겠습니까?"
                )
            else:
                message = (
                    "파일 공유용 방화벽 허용을 취소하시겠습니까?\n"
                    "학생이 파일을 보낼 수 없게 됩니다."
                )

            if not messagebox.askyesno("확인", message):
                return

            error = add_firewall_rule() if allow else remove_firewall_rule()

            if error:
                messagebox.showerror("실패", error)
            else:
                admin_log("방화벽 설정", f"{FILE_SHARE_PORT}번 포트 {'허용' if allow else '차단'}")
                messagebox.showinfo(
                    "완료",
                    f"방화벽 설정을 변경했습니다.\n\n"
                    f"{FILE_SHARE_PORT}번 포트 {'허용' if allow else '차단'}"
                )

            self.kiosk_setting_screen()

        firewall_frame = ctk.CTkFrame(left, fg_color="transparent")
        firewall_frame.pack(pady=(0, 4))

        self.primary_button(
            firewall_frame, "방화벽 제한 해제",
            lambda: apply_firewall(True), width=170, height=46
        ).grid(row=0, column=0, padx=6)

        self.secondary_button(
            firewall_frame, "허용 취소",
            lambda: apply_firewall(False), width=130, height=46
        ).grid(row=0, column=1, padx=6)


        # ==========================================================
        # 오른쪽 — 자동 종료
        # ==========================================================
        self.body_label(right, "자동 종료", size=20, bold=True).pack(pady=(0, 10))

        schedules = list_shutdown_schedules()

        if schedules:
            summary = f"{len(schedules)}개 일정이 등록되어 있습니다."
        else:
            summary = "등록된 일정이 없습니다."

        self.body_label(right, summary, size=15).pack(pady=(0, 8))

        self.body_label(
            right,
            "요일별로 다른 시각을 지정하거나,\n"
            "특정 날짜만 따로 정할 수 있습니다.",
            size=14, muted=True
        ).pack(pady=(0, 11))

        self.secondary_button(
            right, "자동 종료 일정 관리",
            self.shutdown_schedule_screen,
            width=240, height=48
        ).pack(pady=(0, 16))

        # ── 오른쪽 — 프로그램 업데이트 ──
        ctk.CTkFrame(right, fg_color=COLOR_BORDER, height=1).pack(fill="x", pady=(0, 18))

        self.body_label(right, "프로그램 업데이트", size=20, bold=True).pack(pady=(0, 10))

        self.body_label(right, f"현재 버전  v{APP_VERSION}", size=15).pack(pady=(0, 8))

        self.body_label(
            right,
            "버튼을 눌렀을 때만 새 버전을 확인합니다.\n"
            "평소에는 인터넷에 접속하지 않습니다.",
            size=14, muted=True
        ).pack(pady=(0, 11))

        self.secondary_button(
            right, "업데이트 확인",
            self.run_update_check,
            width=240, height=48
        ).pack(pady=(0, 4))

        # ==========================================================
        # 하단 공통
        # ==========================================================
        def save_kiosk_setting():
            value = "1" if fullscreen_var.get() == 1 else "0"
            set_setting("fullscreen", value)

            self.fullscreen = value == "1"
            self.apply_fullscreen()

            key_value = "1" if key_block_var.get() == 1 else "0"
            set_setting("block_system_keys", key_value)

            key_message = ""
            if key_value == "1":
                if start_key_blocker():
                    key_message = "\n단축키 차단이 적용되었습니다."
                else:
                    reason = get_key_blocker_error() or "알 수 없는 오류"
                    key_message = f"\n단축키 차단 실패: {reason}"
            else:
                stop_key_blocker()
                key_message = "\n단축키 차단을 해제했습니다."

            messagebox.showinfo("완료", "보호 기능 설정을 저장했습니다." + key_message)
            self.admin_menu()

        ctk.CTkFrame(card, fg_color=COLOR_BORDER, height=1).pack(fill="x", pady=(16, 14))

        self.primary_button(card, "저장", save_kiosk_setting, width=200, height=52).pack(pady=(0, 12))
        self.link_button(card, "관리자 메뉴로", self.admin_menu, width=180, height=42).pack()


if __name__ == "__main__":
    if ctk is None:
        _root = tk.Tk()
        _root.withdraw()
        messagebox.showerror(
            "패키지 필요",
            "이 프로그램을 실행하려면 customtkinter 패키지가 필요합니다.\n"
            "명령 프롬프트에서 'pip install customtkinter'를 실행한 뒤 다시 실행하세요."
        )
    else:
        # 모니터 배율(125%, 150% 등)이 걸린 PC에서 화면이 잘리거나
        # 흐릿하게 나오지 않도록 DPI 인식을 켠다.
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        root = ctk.CTk()
        app = PrinterKioskApp(root)

        try:
            root.mainloop()
        finally:
            # 프로그램이 꺼질 때 전체화면 브라우저가 남아 있으면
            # 학생이 그 화면에 갇히므로 반드시 같이 정리한다
            try:
                app.close_external_site()
            except Exception:
                pass

            # 파일 공유용 웹서버도 함께 닫는다
            try:
                stop_upload_server()
            except Exception:
                pass

            stop_key_blocker()
