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
APP_VERSION = "1.0.0"
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
def connect_db():
    return sqlite3.connect(DB_NAME)


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
        "admin_password": DEFAULT_ADMIN_PASSWORD,
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
    conn = connect_db()
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO settings(setting_name, setting_value)
    VALUES(?, ?)
    ON CONFLICT(setting_name) DO UPDATE SET setting_value=excluded.setting_value
    """, (name, str(value)))
    conn.commit()
    conn.close()


def get_daily_limit():
    return int(get_setting("daily_limit", "3"))


def get_max_files_per_job():
    return int(get_setting("max_files_per_job", "3"))


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
    today = str(date.today())
    card_code = normalize_card_code(card_code)

    conn = connect_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT COUNT(*) FROM print_logs WHERE student_id=? AND print_date=? AND result!='실패'",
        (card_code, today)
    )
    count = cur.fetchone()[0]
    conn.close()
    return count


def save_print_log(card_code, name, file_name, printer_name, result):
    now = datetime.now()
    card_code = normalize_card_code(card_code)

    conn = connect_db()
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO print_logs
    (student_id, student_name, file_name, print_date, print_time, printer_name, result)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        card_code,
        name,
        file_name,
        str(date.today()),
        now.strftime("%H:%M:%S"),
        printer_name,
        result
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


def make_temp_pdf_path(original_path):
    out_dir = os.path.join(tempfile.gettempdir(), "school_printer_converted")
    os.makedirs(out_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(original_path))[0]
    safe_name = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in base_name)
    timestamp = int(time.time() * 1000)

    return os.path.join(out_dir, f"{safe_name}_{timestamp}.pdf")


def send_pdf_to_printer(file_path, printer_name):
    sumatra = find_sumatra_pdf()

    if not sumatra:
        raise RuntimeError("SumatraPDF.exe를 찾을 수 없습니다.")

    if not printer_name:
        raise RuntimeError("선택된 프린터가 없습니다. 관리자 화면에서 프린터를 선택하세요.")

    cmd = [
        sumatra,
        "-print-to",
        printer_name,
        "-silent",
        file_path
    ]

    subprocess.Popen(cmd)


def convert_libreoffice_to_pdf(file_path):
    soffice = find_libreoffice()

    if not soffice:
        raise RuntimeError("LibreOffice를 찾을 수 없습니다.")

    out_dir = os.path.join(tempfile.gettempdir(), "school_printer_converted")
    os.makedirs(out_dir, exist_ok=True)

    before_files = set(os.listdir(out_dir))

    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        out_dir,
        file_path
    ]

    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=180
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


def print_document(file_path, printer_name):
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".pdf":
        send_pdf_to_printer(file_path, printer_name)
        return

    if ext in LIBREOFFICE_EXTENSIONS:
        pdf_path = convert_libreoffice_to_pdf(file_path)
        send_pdf_to_printer(pdf_path, printer_name)
        return

    if ext in IMAGE_EXTENSIONS:
        pdf_path = convert_image_to_pdf(file_path)
        send_pdf_to_printer(pdf_path, printer_name)
        return

    if ext in [".hwp", ".hwpx"]:
        print_hwp_by_default_program(file_path, printer_name)
        return

    raise RuntimeError("지원하지 않는 파일 형식입니다.")


def is_supported_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    return ext in SUPPORTED_EXTENSIONS


# -----------------------------
# 출력 미리보기 관련 함수
# -----------------------------
PREVIEW_UNSUPPORTED_EXTENSIONS = [".hwp", ".hwpx"]


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

        out_dir = os.path.join(tempfile.gettempdir(), "school_printer_preview")
        os.makedirs(out_dir, exist_ok=True)

        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        safe_name = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in base_name)
        timestamp = int(time.time() * 1000)
        preview_path = os.path.join(out_dir, f"{safe_name}_{timestamp}_preview.png")

        pix.save(preview_path)
        return preview_path
    finally:
        doc.close()


def get_preview_image_path(file_path):
    """
    출력 전 미리볼 이미지 파일 경로를 반환한다.
    이미지 파일은 그대로, PDF/오피스 문서는 첫 페이지를 이미지로 변환해 반환한다.
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext in IMAGE_EXTENSIONS:
        return file_path

    if ext == ".pdf":
        return render_pdf_first_page_image(file_path)

    if ext in LIBREOFFICE_EXTENSIONS:
        pdf_path = convert_libreoffice_to_pdf(file_path)
        return render_pdf_first_page_image(pdf_path)

    if ext in PREVIEW_UNSUPPORTED_EXTENSIONS:
        raise RuntimeError("한글(HWP) 파일은 미리보기를 지원하지 않습니다.\n출력은 정상적으로 진행됩니다.")

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
           p.file_name, p.printer_name, p.result
    FROM print_logs p
    LEFT JOIN students s ON s.student_id = p.student_id
    {clause}
    ORDER BY p.print_date DESC, p.print_time DESC
    """, params)
    rows = cur.fetchall()
    conn.close()

    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["날짜", "시간", "학생증코드", "학번", "이름", "파일명", "프린터", "결과"])
        writer.writerows(rows)

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
           SUM(CASE WHEN p.result = '실패' THEN 1 ELSE 0 END),
           COUNT(*),
           MIN(p.print_date),
           MAX(p.print_date)
    FROM print_logs p
    LEFT JOIN students s ON s.student_id = p.student_id
    {clause}
    GROUP BY p.student_id
    ORDER BY 4 DESC
    """, params)
    rows = cur.fetchall()
    conn.close()

    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            "학생증코드", "학번", "이름",
            "성공 출력 수", "실패 수", "전체 시도 수",
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
FILE_SHARE_URL = "https://pairdrop.net"


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
# 저수준 키보드 훅 (윈도우 키 / Alt+Tab / Ctrl+Esc 등 실시간 차단)
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
    """schtasks 명령을 실행하고 (성공여부, 출력) 을 돌려준다."""
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


def _shutdown_command():
    return (
        'shutdown.exe /s /f /t 60 '
        '/c "자동 종료 시간입니다. 1분 후 컴퓨터가 꺼집니다."'
    )


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

        return _run_schtasks(base_args + schedule_args)

    # 특정 날짜 1회 실행
    year, month, day = run_date.split("-")
    date_formats = [
        f"{year}-{month}-{day}",
        f"{month}/{day}/{year}",
        f"{day}/{month}/{year}",
        f"{year}/{month}/{day}",
    ]

    last_output = ""

    for date_text in date_formats:
        ok, output = _run_schtasks(
            base_args + ["/SC", "ONCE", "/SD", date_text]
        )
        if ok:
            return True, output
        last_output = output

    return False, last_output


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
    """일정 하나를 지운다 (작업 스케줄러 + DB)."""
    _run_schtasks(["/Delete", "/TN", f"{SHUTDOWN_TASK_PREFIX}{schedule_id}", "/F"])

    conn = connect_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM shutdown_schedules WHERE id=?", (schedule_id,))
    conn.commit()
    conn.close()

    return True


def delete_all_shutdown_schedules():
    """모든 일정을 지운다."""
    for row in list_shutdown_schedules():
        _run_schtasks(["/Delete", "/TN", f"{SHUTDOWN_TASK_PREFIX}{row[0]}", "/F"])

    # 예전 버전에서 만든 작업도 함께 정리
    _run_schtasks(["/Delete", "/TN", "SchoolPrinter_AutoShutdown", "/F"])

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
      1) 프로그램이 완전히 꺼질 때까지 잠시 기다림
      2) 기존 exe 를 새 파일로 교체
      3) 새 버전 실행
      4) 자기 자신 삭제
    """
    if not getattr(sys, "frozen", False):
        raise RuntimeError(
            "exe 로 빌드한 상태에서만 업데이트할 수 있습니다.\n"
            "파이썬 파일로 실행 중일 때는 코드를 직접 받아 주세요."
        )

    current_exe = sys.executable
    backup_exe = current_exe + ".old"

    bat_path = os.path.join(
        tempfile.gettempdir(), "school_printer_update", "apply_update.bat"
    )

    script = f"""@echo off
chcp 65001 > nul
timeout /t 2 /nobreak > nul

:retry
del "{backup_exe}" > nul 2>&1
move "{current_exe}" "{backup_exe}" > nul 2>&1
if errorlevel 1 (
    timeout /t 2 /nobreak > nul
    goto retry
)

move "{new_file_path}" "{current_exe}" > nul 2>&1
if errorlevel 1 (
    move "{backup_exe}" "{current_exe}" > nul 2>&1
    echo 업데이트에 실패했습니다.
    pause
    exit
)

del "{backup_exe}" > nul 2>&1
start "" "{current_exe}"
del "%~f0"
"""

    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(script)

    subprocess.Popen(
        ["cmd", "/c", bat_path],
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


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
        self.fullscreen = get_setting("fullscreen", "1") == "1"
        self.apply_fullscreen()

        self.root.protocol("WM_DELETE_WINDOW", self.block_close)
        self.bind_blocked_shortcuts()
        self.setup_scanner_focus()

        # 윈도우 키 등 시스템 단축키 차단 (설정에 따라)
        if get_setting("block_system_keys", "0") == "1":
            start_key_blocker()

        self.start_screen()

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
        (Alt+F4, Alt+Tab, F11, Ctrl+W, Ctrl+Esc 등)
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
            self.open_external_site(REPORT_FORM_URL)
            return "break"

        if keycode == 83:      # S
            self.open_external_site(FILE_SHARE_URL)
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
        for widget in self.root.winfo_children():
            widget.destroy()

    def build_card(self):
        """
        화면 배경 위에 중앙 정렬된 카드 패널을 만들고, 그 카드를 반환한다.
        각 화면은 이 카드 안에 위젯을 채운다.
        """
        self.clear()

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

        return content

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
            self.root.destroy()

    # -----------------------------
    # 학생용 화면
    # -----------------------------
    def start_screen(self):
        card = self.build_card()
        self.make_title(card, "서인천고등학교 공용 프린터 제어 시스템")

        badge = ctk.CTkFrame(card, fg_color=COLOR_PRIMARY_SOFT, corner_radius=14)
        badge.pack(pady=(0, 28))
        ctk.CTkLabel(
            badge,
            text="학생증 바코드를 스캔하세요",
            font=(self.font_family, 26, "bold"),
            text_color=COLOR_PRIMARY
        ).pack(padx=28, pady=20)

        self.card_entry = self.make_entry(card, width=460, height=60, font_size=28)
        self.card_entry.pack(pady=(0, 24))
        self.card_entry.bind("<Return>", lambda event: self.check_user())

        # 학생이 입력창을 클릭하지 않아도 바코드가 바로 입력되도록
        # 이 화면에서는 항상 입력창에 포커스를 유지한다
        self.current_screen = "start"
        self.card_entry.after(50, self.focus_scan_entry)

        self.primary_button(card, "확인", self.check_user, width=240, height=60, font_size=20).pack(pady=(0, 28))

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
            lambda: self.open_external_site(REPORT_FORM_URL),
            width=150, height=44, font_size=14,
            fg_color=COLOR_BG, hover_color=COLOR_SECONDARY_BG
        ).place(relx=1.0, rely=1.0, anchor="se", x=-28, y=-24)

        self.secondary_button(
            self.outer, "파일 공유",
            lambda: self.open_external_site(FILE_SHARE_URL),
            width=150, height=44, font_size=14,
            fg_color=COLOR_BG, hover_color=COLOR_SECONDARY_BG
        ).place(relx=1.0, rely=1.0, anchor="se", x=-188, y=-24)

    def open_external_site(self, url):
        """
        외부 사이트를 별도 브라우저 창으로 연다.

        --kiosk는 닫기 버튼이 없어서 학생이 브라우저를 못 끄므로 --app을 쓴다.
        --app은 주소창/탭 없이 뜨면서 창 조절·닫기 버튼은 그대로 있다.

        또 이 프로그램이 항상 위(topmost)로 떠 있으면 브라우저가 뒤에 가려져서
        조작이 안 되므로, 브라우저가 떠 있는 동안에는 topmost를 잠시 해제하고
        브라우저가 닫히면 자동으로 원래대로 되돌린다.
        """
        if not url:
            messagebox.showinfo("안내", "아직 주소가 설정되지 않았습니다. 관리자에게 문의하세요.")
            return

        # 이미 브라우저를 띄운 상태면 무시한다.
        # (버튼을 연타하거나 단축키를 여러 번 눌러도 창이 하나만 뜨게)
        if self.external_browser_open:
            return

        # 전용 프로필로 띄우므로 프로세스가 살아 있으면 창도 떠 있는 것이다.
        # 전체화면이 아닐 때는 위 플래그가 설정되지 않으므로 이쪽으로 판정한다.
        previous = getattr(self, "_external_process", None)
        if previous is not None and previous.poll() is None:
            return

        # 창이 실제로 뜨기까지 시간이 걸리는데, 그 사이에 또 누르면
        # 위 검사들을 모두 통과해서 두 번 뜬다. 짧은 시간 동안 잠가 둔다.
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

        # 크롬/엣지는 이미 실행 중인 인스턴스가 있으면 새 프로세스가 창을 넘기고
        # 즉시 종료한다. 그러면 아래 최대화 처리에서 창을 못 찾는다.
        # 전용 프로필 폴더를 지정하면 항상 독립 프로세스로 뜬다.
        profile_dir = os.path.join(tempfile.gettempdir(), "school_printer_browser")

        for path in chrome_paths + edge_paths:
            if os.path.exists(path):
                try:
                    process = subprocess.Popen([
                        path,
                        f"--app={url}",
                        f"--user-data-dir={profile_dir}",
                        "--no-first-run",
                        "--no-default-browser-check",
                    ])
                    break
                except Exception:
                    continue

        if process is None:
            try:
                webbrowser.open(url)
            except Exception as e:
                # 실행에 실패했으면 잠금을 풀어서 다시 시도할 수 있게 한다
                self._last_external_open = 0
                messagebox.showerror("오류", f"사이트를 열 수 없습니다.\n{e}")
                return
        else:
            self._external_process = process
            self.maximize_browser_window(process.pid)

        self.release_topmost_until_return()

    def maximize_browser_window(self, pid):
        """
        방금 띄운 브라우저 창을 최대화한다.

        --window-size 는 창 크기만 지정할 뿐 '최대화 상태'가 아니어서
        제목표시줄의 최대화 버튼이 눌린 것과 다르다.
        창이 뜨는 데 시간이 걸리므로 잠시 기다리며 여러 번 시도한다.
        """
        def worker():
            try:
                import ctypes
                import ctypes.wintypes as wintypes
            except ImportError:
                return

            user32 = ctypes.windll.user32
            SW_MAXIMIZE = 3

            EnumWindowsProc = ctypes.WINFUNCTYPE(
                ctypes.c_bool, wintypes.HWND, wintypes.LPARAM
            )

            found = []

            def callback(hwnd, lparam):
                if not user32.IsWindowVisible(hwnd):
                    return True

                window_pid = wintypes.DWORD()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))

                if window_pid.value == pid:
                    # 제목이 있는 실제 창만 (숨은 보조 창 제외)
                    if user32.GetWindowTextLengthW(hwnd) > 0:
                        found.append(hwnd)
                        return False

                return True

            # 창이 뜰 때까지 최대 6초 대기
            for _ in range(30):
                time.sleep(0.2)
                found.clear()

                try:
                    user32.EnumWindows(EnumWindowsProc(callback), 0)
                except Exception:
                    return

                if found:
                    hwnd = found[0]
                    user32.ShowWindow(hwnd, SW_MAXIMIZE)
                    user32.SetForegroundWindow(hwnd)
                    return

        threading.Thread(target=worker, daemon=True).start()

    def release_topmost_until_return(self):
        """
        외부 브라우저가 떠 있는 동안 '항상 위' 속성을 풀어둔다.

        브라우저 프로세스를 기다리는 방식은 쓸 수 없다.
        크롬/엣지는 이미 실행 중인 인스턴스가 있으면 새 프로세스가 창을 넘기고
        즉시 종료해버려서, process.wait()가 곧바로 끝나며 창이 다시 가려진다.

        대신 이 프로그램 창이 다시 활성화되는 시점(= 브라우저가 닫힌 시점)을
        감지해서 키오스크 상태로 되돌린다.
        """
        if not self.fullscreen:
            return

        self.external_browser_open = True
        self.root.attributes("-topmost", False)
        self.root.lower()

        def check_returned():
            if not self.external_browser_open:
                return

            try:
                # 이 프로그램 창이 다시 활성 창이 되면 브라우저가 닫힌 것으로 본다
                if self.root.focus_displayof() is not None:
                    self.restore_kiosk_focus()
                    return
            except Exception:
                pass

            self.root.after(700, check_returned)

        # 브라우저가 뜨는 데 시간이 걸리므로 잠시 뒤부터 감시 시작
        self.root.after(2500, check_returned)

    def restore_kiosk_focus(self):
        self.external_browser_open = False

        if not self.fullscreen:
            return

        self.root.attributes("-topmost", True)
        self.root.lift()
        self.root.focus_force()

    def reset_scan_entry(self):
        """입력창을 비우고 다시 포커스를 준다 (다음 스캔 대기 상태)."""
        if self.card_entry is None or not self.card_entry.winfo_exists():
            return

        try:
            self.card_entry.delete(0, tk.END)
        except Exception:
            pass

        self.focus_scan_entry()

    def check_user(self):
        raw_code = self.card_entry.get().strip()
        card_code = normalize_card_code(raw_code)

        if not card_code:
            messagebox.showwarning("입력 오류", "학생증 코드를 입력하세요.")
            return

        # 관리자 바코드 코드 감지 (이스터에그)
        admin_code = get_setting("admin_barcode", "")
        if admin_code and normalize_card_code(admin_code) == card_code:
            self.reset_scan_entry()
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

        if not unlimited:
            today_count = get_today_count(card_code)
            daily_limit = get_daily_limit()

            if today_count >= daily_limit:
                messagebox.showerror(
                    "출력 제한",
                    f"{name} 사용자는 오늘 출력 가능 횟수를 초과했습니다.\n"
                    f"오늘 출력 횟수: {today_count}/{daily_limit}"
                )
                self.start_screen()
                return

        self.file_screen(card_code, name, unlimited)

    def file_screen(self, card_code, name, unlimited):
        card = self.build_card()
        self.make_title(card, "파일 선택 화면")

        today_count = get_today_count(card_code)
        daily_limit = get_daily_limit()
        remain = daily_limit - today_count
        max_files = get_max_files_per_job()

        if unlimited:
            limit_text = "무제한 출력 사용자"
            remain_text = "제한 없음"
            max_text = "제한 없음"
        else:
            limit_text = f"오늘 출력 횟수: {today_count}/{daily_limit}"
            remain_text = f"{remain}회"
            max_text = f"최대 {max_files}개"

        info_card = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=16)
        info_card.pack(pady=(0, 28), padx=6, ipadx=24, ipady=22)

        student_number = get_student_number(card_code)

        info_rows = [("학생증 코드", card_code)]

        if student_number:
            info_rows.append(("학번", student_number))

        info_rows += [
            ("이름", name),
            ("오늘 출력 현황", limit_text),
            ("남은 출력 가능 횟수", remain_text),
            ("한 번에 선택 가능 파일 수", max_text),
        ]

        for label_text, value_text in info_rows:
            row = ctk.CTkFrame(info_card, fg_color="transparent")
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
            today_count = get_today_count(card_code)
            daily_limit = get_daily_limit()
            remain = daily_limit - today_count
            max_files = get_max_files_per_job()

            if len(file_paths) > max_files:
                messagebox.showerror(
                    "파일 개수 제한",
                    f"한 번에 최대 {max_files}개까지만 선택할 수 있습니다."
                )
                return

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

        today_count = get_today_count(card_code)
        daily_limit = get_daily_limit()
        after_count = today_count + len(file_paths)

        if unlimited:
            count_text = "무제한 출력 사용자이므로 출력 횟수 제한 없음"
        else:
            count_text = f"출력 후 오늘 출력 횟수: {after_count}/{daily_limit}"

        student_number = get_student_number(card_code)

        if student_number:
            user_text = f"{student_number} {name} ({card_code})"
        else:
            user_text = f"{name} ({card_code})"

        info_card = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=16)
        info_card.pack(pady=(0, 22), padx=6, ipadx=24, ipady=18, fill="x")

        self.body_label(
            info_card,
            f"사용자: {user_text}\n"
            f"선택된 프린터: {printer_name}\n"
            f"선택한 파일 수: {len(file_paths)}개\n"
            f"{count_text}",
            size=16, justify="left"
        ).pack(anchor="w", padx=8)

        columns = ("no", "file_name", "file_type")
        tree = ttk.Treeview(card, columns=columns, show="headings", height=10, style="App.Treeview")
        tree.heading("no", text="번호")
        tree.heading("file_name", text="파일명")
        tree.heading("file_type", text="형식")

        tree.column("no", width=60, anchor="center") 
        tree.column("file_name", width=580)
        tree.column("file_type", width=100, anchor="center")

        tree.pack(pady=(0, 10))

        for idx, path in enumerate(file_paths, start=1):
            file_name = os.path.basename(path)
            ext = os.path.splitext(path)[1].lower()
            tree.insert("", "end", values=(idx, file_name, ext))

        tree.bind("<Double-1>", lambda event: self.show_file_preview(tree, file_paths))

        self.body_label(
            card,
            "목록에서 파일을 선택한 뒤 '미리보기'를 누르거나, 파일을 더블클릭하세요.",
            size=15, muted=True
        ).pack(pady=(0, 22))

        btn_frame = ctk.CTkFrame(card, fg_color="transparent") 
        btn_frame.pack(pady=(0, 8))

        self.secondary_button(
            btn_frame, "미리보기",
            lambda: self.show_file_preview(tree, file_paths),
            width=200, height=54, font_size=16
        ).grid(row=0, column=0, padx=8)

        self.primary_button(
            btn_frame, "출력 시작",
            lambda: self.do_print_files(card_code, name, file_paths),
            width=200, height=54, font_size=16
        ).grid(row=0, column=1, padx=8)

        self.secondary_button(
            btn_frame, "취소", 
            lambda: self.file_screen(card_code, name, unlimited),
            width=200, height=54, font_size=16
        ).grid(row=0, column=2, padx=8)

    def show_file_preview(self, tree, file_paths):
        selected = tree.selection()

        if not selected:
            messagebox.showwarning("선택 오류", "미리볼 파일을 목록에서 선택하세요.")
            return

        values = tree.item(selected[0], "values")
        idx = int(values[0]) - 1

        if idx < 0 or idx >= len(file_paths):
            return

        file_path = file_paths[idx]
        file_name = os.path.basename(file_path)

        if Image is None or ImageTk is None:
            messagebox.showerror("미리보기 오류", "미리보기를 위해 pillow(PIL) 패키지가 필요합니다.")
            return

        preview_win = ctk.CTkToplevel(self.root)
        preview_win.title(f"미리보기 - {file_name}")
        preview_win.after(250, lambda: apply_window_icon(preview_win))
        preview_win.geometry("780x840")
        preview_win.configure(fg_color=COLOR_BG)
        preview_win.transient(self.root)
        preview_win.attributes("-topmost", True)
        preview_win.lift()
        preview_win.after(120, preview_win.grab_set)

        ctk.CTkLabel(
            preview_win, text=file_name,
            font=(self.font_family, 20, "bold"), text_color=COLOR_TEXT
        ).pack(pady=(26, 12))

        status_label = ctk.CTkLabel(
            preview_win,
            text="미리보기를 생성하는 중입니다. 잠시만 기다려 주세요...",
            font=(self.font_family, 14),
            text_color=COLOR_TEXT_MUTED
        )
        status_label.pack(pady=18)

        self.secondary_button(preview_win, "닫기", preview_win.destroy, width=160, height=48).pack(side="bottom", pady=22)

        preview_win.update()

        try:
            image_path = get_preview_image_path(file_path)

            with Image.open(image_path) as img:
                img = img.copy()
            img.thumbnail((700, 680))
            photo = ImageTk.PhotoImage(img)

            status_label.configure(text="")

            img_frame = ctk.CTkFrame(
                preview_win, fg_color=COLOR_CARD, corner_radius=14,
                border_width=1, border_color=COLOR_BORDER
            )
            img_frame.pack(pady=6, padx=20)

            img_label = tk.Label(img_frame, image=photo, bg=COLOR_CARD, borderwidth=0)
            img_label.image = photo
            img_label.pack(padx=8, pady=8)

            note = ""
            ext = os.path.splitext(file_path)[1].lower()
            if ext != ".pdf" and ext not in IMAGE_EXTENSIONS:
                note = "(실제 인쇄 결과와 여백/서식이 다를 수 있습니다. 첫 페이지만 표시됩니다.)"

            if note:
                ctk.CTkLabel(
                    preview_win, text=note,
                    font=(self.font_family, 10), text_color=COLOR_TEXT_MUTED
                ).pack(pady=(8, 0))

        except Exception as e:
            status_label.configure(text=str(e), text_color=COLOR_DANGER, wraplength=680, justify="center")

    def do_print_files(self, card_code, name, file_paths):
        printer_name = get_setting("selected_printer", "")

        if not printer_name:
            messagebox.showerror("프린터 오류", "선택된 프린터가 없습니다. 관리자 화면에서 프린터를 선택하세요.")
            return

        success_count = 0
        fail_count = 0
        errors = []

        for file_path in file_paths:
            file_name = os.path.basename(file_path)

            try:
                print_document(file_path, printer_name)
                save_print_log(card_code, name, file_name, printer_name, "전송완료")
                success_count += 1

            except Exception as e:
                save_print_log(card_code, name, file_name, printer_name, "실패")
                fail_count += 1
                errors.append(f"{file_name}: {e}")

        message = (
            f"출력 전송 완료\n\n"
            f"성공: {success_count}개\n"
            f"실패: {fail_count}개"
        )

        if errors:
            message += "\n\n오류 내용:\n" + "\n".join(errors[:5])

        messagebox.showinfo("출력 결과", message)
        self.start_screen()

    # -----------------------------
    # 관리자 화면
    # -----------------------------
    def admin_login_screen(self):
        card = self.build_card()
        self.make_title(card, "관리자 로그인")

        self.body_label(card, "관리자 비밀번호 입력", size=15).pack(pady=(0, 14))

        password_entry = self.make_entry(card, width=280, height=48, font_size=16, show="*")
        password_entry.pack(pady=(0, 24))
        password_entry.focus()

        def check_password():
            password = password_entry.get()
            real_password = get_setting("admin_password", DEFAULT_ADMIN_PASSWORD)

            if password == real_password:
                self.admin_menu()
            else:
                messagebox.showerror("접근 거부", "관리자 비밀번호가 틀렸습니다.")

        password_entry.bind("<Return>", lambda event: check_password())

        self.primary_button(card, "로그인", check_password, width=240, height=54).pack(pady=(0, 14))
        self.link_button(card, "처음 화면으로", self.start_screen, width=180, height=42).pack()

    def admin_menu(self):
        card = self.build_card()
        self.make_title(card, "관리자 모드")

        buttons = [
            ("이용자 관리", self.user_manage_screen),
            ("출력 제한 설정", self.limit_setting_screen),
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
        input_frame.pack(pady=(0, 12), ipadx=16, ipady=12)

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
        search_frame.pack(pady=(0, 8))

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

        tree.pack(pady=(0, 10))

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
                messagebox.showinfo("완료", "상태를 변경했습니다.")
                refresh_users()

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
                messagebox.showinfo(
                    "불러오기 완료",
                    f"불러온 이용자: {imported}명\n건너뛴 행: {skipped}개"
                )
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
                messagebox.showinfo("내보내기 완료", "이용자 명단을 CSV로 저장했습니다.")
            except Exception as e:
                messagebox.showerror("내보내기 오류", str(e))

        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.pack(pady=(4, 6))

        self.primary_button(btn_frame, "추가/수정 저장", save_user, width=160, height=44, font_size=16).grid(row=0, column=0, padx=6)
        self.danger_button(btn_frame, "비활성화", lambda: set_active(0), width=130, height=44, font_size=16).grid(row=0, column=1, padx=6)
        self.secondary_button(btn_frame, "다시 활성화", lambda: set_active(1), width=130, height=44, font_size=16).grid(row=0, column=2, padx=6)
        self.secondary_button(btn_frame, "새로고침", refresh_users, width=120, height=44, font_size=16).grid(row=0, column=3, padx=6)

        file_btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        file_btn_frame.pack(pady=(0, 8))

        self.secondary_button(file_btn_frame, "CSV/XLSX 불러오기", import_users, width=220, height=42, font_size=16).grid(row=0, column=0, padx=6)
        self.secondary_button(file_btn_frame, "이용자 명단 CSV 내보내기", export_users, width=240, height=42, font_size=16).grid(row=0, column=1, padx=6)

        self.link_button(card, "관리자 메뉴로", self.admin_menu, width=180, height=38).pack()

        refresh_users()

    def limit_setting_screen(self):
        card = self.build_card()
        self.make_title(card, "출력 제한 설정")

        current_daily = get_daily_limit()
        current_max_files = get_max_files_per_job()

        frame = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=16)
        frame.pack(pady=(0, 20), ipadx=20, ipady=20)

        self.body_label(frame, "일반 사용자 하루 출력 제한 횟수", size=14).grid(row=0, column=0, padx=12, pady=12)
        daily_entry = self.make_entry(frame, width=140, height=44, font_size=15)
        daily_entry.grid(row=0, column=1, padx=12, pady=12)
        daily_entry.insert(0, str(current_daily))

        self.body_label(frame, "일반 사용자 한 번에 선택 가능한 파일 수", size=14).grid(row=1, column=0, padx=12, pady=12)
        max_file_entry = self.make_entry(frame, width=140, height=44, font_size=15)
        max_file_entry.grid(row=1, column=1, padx=12, pady=12)
        max_file_entry.insert(0, str(current_max_files))

        self.body_label(
            card, "무제한 출력 사용자는 위 제한을 적용받지 않습니다.", size=14, muted=True
        ).pack(pady=(0, 22))

        divider = ctk.CTkFrame(card, fg_color=COLOR_BORDER, height=1)
        divider.pack(fill="x", pady=(0, 22))

        self.body_label(card, "파일 선택 기본 폴더", size=20, bold=True).pack(pady=(0, 12))

        current_dir = get_setting("file_open_dir", "")
        dir_display = current_dir if current_dir else "설정 안 됨 (바탕화면)"

        self.body_label(
            card, f"현재: {dir_display}", size=15, muted=True
        ).pack(pady=(0, 12))

        dir_entry = self.make_entry(card, width=560, height=46, font_size=15, justify="left")
        dir_entry.pack(pady=(0, 8))
        if current_dir:
            dir_entry.insert(0, current_dir)

        self.body_label(
            card,
            "경로를 직접 입력하거나, 아래 '폴더 선택' 버튼으로 고르세요.\n"
            "비워두면 바탕화면에서 시작합니다.",
            size=15, muted=True
        ).pack(pady=(0, 14))

        def browse_dir():
            selected = filedialog.askdirectory(title="기본 폴더 선택")
            if selected:
                dir_entry.delete(0, tk.END)
                dir_entry.insert(0, selected.replace("/", "\\"))

        self.secondary_button(card, "폴더 선택", browse_dir, width=160, height=44).pack(pady=(0, 22))

        def save_limits():
            daily_value = daily_entry.get().strip()
            max_file_value = max_file_entry.get().strip()

            if not daily_value.isdigit() or int(daily_value) <= 0:
                messagebox.showwarning("입력 오류", "하루 출력 제한은 1 이상의 숫자여야 합니다.")
                return

            if not max_file_value.isdigit() or int(max_file_value) <= 0:
                messagebox.showwarning("입력 오류", "파일 개수 제한은 1 이상의 숫자여야 합니다.")
                return

            dir_value = dir_entry.get().strip()

            if dir_value and not os.path.isdir(dir_value):
                messagebox.showwarning("경로 오류", f"폴더를 찾을 수 없습니다.\n{dir_value}")
                return

            set_setting("daily_limit", daily_value)
            set_setting("max_files_per_job", max_file_value)
            set_setting("file_open_dir", dir_value)

            messagebox.showinfo(
                "완료",
                f"하루 출력 제한: {daily_value}회\n"
                f"한 번에 선택 가능한 파일 수: {max_file_value}개\n"
                f"파일 선택 기본 폴더: {dir_value if dir_value else '바탕화면'}"
            )
            self.admin_menu()

        self.primary_button(card, "저장", save_limits, width=200, height=52).pack(pady=(0, 12))
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
            columns=("c1", "c2", "c3", "c4", "c5", "c6", "c7"),
            show="headings",
            height=15,
            style="App.Treeview"
        )

        headings = ["학생증 코드", "학번", "이름", "파일명/횟수", "날짜", "시간", "결과"]
        widths = [150, 120, 140, 340, 180, 130, 110]

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
               p.file_name, p.print_date, p.print_time, p.result
        FROM print_logs p
        LEFT JOIN students s ON s.student_id = p.student_id
        WHERE p.print_date=?
        ORDER BY p.print_time DESC
        """, (str(date.today()),))
        rows = cur.fetchall()
        conn.close()

        for row in rows:
            self.log_tree.insert("", "end", values=row)

    def show_user_counts(self):
        self.clear_log_tree()

        conn = connect_db()
        cur = conn.cursor()
        cur.execute("""
        SELECT s.student_id, COALESCE(s.student_number, ''), s.name, COUNT(p.id) AS cnt
        FROM students s
        LEFT JOIN print_logs p
        ON s.student_id = p.student_id
        AND p.print_date = ?
        AND p.result != '실패'
        GROUP BY s.student_id, s.student_number, s.name
        ORDER BY cnt DESC
        """, (str(date.today()),))
        rows = cur.fetchall()
        conn.close()

        for card_code, student_number, name, cnt in rows:
            self.log_tree.insert(
                "",
                "end",
                values=(card_code, student_number, name, f"오늘 {cnt}회", str(date.today()), "-", "-")
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
               p.file_name, p.print_date, p.print_time, p.result
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

        for row in rows:
            self.log_tree.insert("", "end", values=row)

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
        messagebox.showinfo("완료", "전체 출력 기록을 초기화했습니다.")

    def change_admin_password_screen(self):
        card = self.build_card()
        self.make_title(card, "비밀번호 설정")

        # ── 비밀번호 변경 ──
        frame = ctk.CTkFrame(card, fg_color=COLOR_ENTRY_BG, corner_radius=14)
        frame.pack(pady=(0, 16), ipadx=16, ipady=16)

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

            if old_pw != real_pw:
                messagebox.showerror("오류", "현재 비밀번호가 틀렸습니다.")
                return

            if not new_pw:
                messagebox.showwarning("입력 오류", "새 비밀번호를 입력하세요.")
                return

            if new_pw != confirm_pw:
                messagebox.showerror("오류", "새 비밀번호가 서로 다릅니다.")
                return

            set_setting("admin_password", new_pw)
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
                messagebox.showinfo("완료", f"일정을 추가했습니다.\n\n{label}")
                refresh_list()
            else:
                messagebox.showerror(
                    "등록 실패",
                    f"작업 스케줄러 등록에 실패했습니다.\n\n{output[:300]}\n\n"
                    "관리자 권한으로 프로그램을 실행했는지 확인하세요."
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

            delete_shutdown_schedule(schedule_id)
            messagebox.showinfo("완료", "일정을 삭제했습니다.")
            refresh_list()

        def delete_all():
            if not list_shutdown_schedules():
                messagebox.showinfo("안내", "등록된 일정이 없습니다.")
                return

            if not messagebox.askyesno("확인", "등록된 자동 종료 일정을 모두 삭제할까요?"):
                return

            delete_all_shutdown_schedules()
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
        columns.pack(pady=(0, 8))

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
        ).pack(pady=(0, 22))

        # ── 왼쪽 — Windows 잠금 ──
        ctk.CTkFrame(left, fg_color=COLOR_BORDER, height=1).pack(fill="x", pady=(0, 18))

        self.body_label(left, "Windows 잠금", size=20, bold=True).pack(pady=(0, 10))

        key_block_var = tk.IntVar(value=1 if get_setting("block_system_keys", "0") == "1" else 0)

        ctk.CTkCheckBox(
            left,
            text="윈도우 키 · Alt+Tab · Ctrl+Esc 차단",
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
        ).pack(pady=(0, 14))

        def apply_lock(enabled):
            action = "적용" if enabled else "해제"

            if not messagebox.askyesno("확인", f"작업 관리자·실행창 잠금을 {action}하시겠습니까?"):
                return

            done, failed = set_windows_lockdown(enabled)

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
        ).pack(pady=(0, 14))

        self.secondary_button(
            right, "자동 종료 일정 관리",
            self.shutdown_schedule_screen,
            width=240, height=48
        ).pack(pady=(0, 22))

        # ── 오른쪽 — 프로그램 업데이트 ──
        ctk.CTkFrame(right, fg_color=COLOR_BORDER, height=1).pack(fill="x", pady=(0, 18))

        self.body_label(right, "프로그램 업데이트", size=20, bold=True).pack(pady=(0, 10))

        self.body_label(right, f"현재 버전  v{APP_VERSION}", size=15).pack(pady=(0, 8))

        self.body_label(
            right,
            "버튼을 눌렀을 때만 새 버전을 확인합니다.\n"
            "평소에는 인터넷에 접속하지 않습니다.",
            size=14, muted=True
        ).pack(pady=(0, 14))

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

        ctk.CTkFrame(card, fg_color=COLOR_BORDER, height=1).pack(fill="x", pady=(22, 18))

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
            stop_key_blocker()
