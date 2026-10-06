# ==============================================================================
# main.py — Orchestrator (ΠΛΗΡΗΣ ΡΟΕΣ)
# ==============================================================================
#
# ΦΑΣΗ Α (μια φορά):
#   1. Selenium → ICISnet scraping → df_final
#   2. Pywinauto → SAP export → ΕΙΣΑΓΩΓΕΣ_database.xlsx
#   3. Queries → FULL_RESULTS.xlsx (opened / done / undone)
#
# PAUSE — Popup editable πίνακας:
#   Βλέπεις: MRN, PROT, PDF, PDF.1, ΚΑΤΑΣΤΑΣΗ
#   Checkbox ανά γραμμή | Double-click → inline edit | OK → εκκίνηση
#
# ΦΑΣΗ Β σε δύο μέρη:
# Β.1 (μόνο ICISNet, SAP κλειστό) για ΟΛΑ τα επιλεγμένα MRN σερί:
#   XML -> xml\<MRN>.xml (αν υπάρχει, δεν ξανακατεβαίνει) -> έλεγχος MRN ->
#   parse + φίλτρα -> κράματα -> ΟΛΑ τα popups (ΚΡΑΜΑ/ΠΡΟΤΙΜΗΣΗ/ΕΝΤΟΛΗ ΑΓΟΡΑΣ).
#   Σφάλμα σε ένα MRN δεν σταματάει τα υπόλοιπα. Browser κλείνει.
# Β.2 (μόνο SAP, ΕΝΑ login, κανένα popup) για κάθε MRN:
#   PDF rename -> ανά Α/Α γέμισμα/Save/attach -> PDF move -> Status DONE.
#   Σφάλμα -> σταματάει το batch (νεκρό session). SAP κλείνει.
#
# ==============================================================================

import os
import sys
import io
import ssl
import time
import shutil
import base64
import warnings
import subprocess
import traceback
import logging
import ctypes
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from io import StringIO
from datetime import datetime
from time import perf_counter
from dateutil.relativedelta import relativedelta

import pyautogui
import pandas as pd
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from pywinauto import Application
from pywinauto.keyboard import send_keys

import win32com.client
import win32event

# Encoding για Power Automate
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", line_buffering=True)

# ==============================================================================
# SETTINGS
# ==============================================================================

os.environ["WDM_SSL_VERIFY"] = "0"
ssl._create_default_https_context = ssl._create_unverified_context
warnings.filterwarnings("ignore")

import keyring

ICISNET_USER  = "YOUR_ICISNET_USER"
ICISNET_PASS  = keyring.get_password("ICISNET", ICISNET_USER)
SAP_USER      = "YOUR_SAP_USER"
SAP_PASS      = keyring.get_password("YOUR_SAP_KEYRING", SAP_USER)

# Κοινό mutex σε ΟΛΑ τα scripts που αγγίζουν SAP — βλ. ΤΙΜΟΛΟΓΙΑ.py.
SAP_MUTEX_NAME = "Global\\SAP_Automation_Lock"
COMPANY_CODE = "YOUR_COMPANY_CODE"

for _name, _user, _pass in [("ICISNET", ICISNET_USER, ICISNET_PASS), ("YOUR_SAP_KEYRING", SAP_USER, SAP_PASS)]:
    if _pass is None:
        raise RuntimeError(
            f"Δεν βρέθηκε password στο Credential Manager για '{_name}'. Τρέξε μία φορά:\n"
            f"    python -m keyring set {_name} {_user}\n"
            "και πληκτρολόγησε το password όταν σου ζητηθεί."
        )

_BASE         = Path(r"C:\Users\YOUR_USERNAME")
DESKTOP       = _BASE / "Desktop"
DOCUMENTS     = _BASE / "Documents"
SAP_FILE_A    = DESKTOP / "LIST_N.sap"
SAP_FILE_B    = DESKTOP / "ZELVMM_IMP_1.sap"
BROKER_XLSX     = DESKTOP / "BROKER" / "MRN.xlsx"
DATABASE_XLSX = DESKTOP / "ΕΙΣΑΓΩΓΕΣ_database.xlsx"
SAP_GUI_DIR   = DOCUMENTS / "SAP" / "SAP GUI"
PDF_SAVE_PATH = DESKTOP / "Αναζήτηση _ Αποτελέσματα Αναζήτησης.pdf"
OUTPUT_EXCEL  = Path(__file__).resolve().parent / "FULL_RESULTS.xlsx"
SAVE_FOLDER   = Path(__file__).resolve().parent / "xml_temp"   # λήψη Chrome (αδειάζει πριν από κάθε λήψη)
XML_DIR       = Path(__file__).resolve().parent / "xml"        # ένα XML ανά MRN — αν υπάρχει, δεν ξανακατεβαίνει
ARCHIVE_BASE    = Path(r"\\YOUR_SERVER\YOUR_SHARE\IMPORT_DECLARATIONS")

KATH_DIR = {
    3:  "ΑΠΑΛΛΑΓΗ ΦΠΑ",
    2:  "ΕΝΕΡΓΗΤΙΚΗ",
    5:  "ΕΛΕΥΘΕΡΑ",
    12: "ΕΝΕΡΓΗΤΙΚΗ - INF",
}


def compute_kath(x16: str, tk_: str) -> int:
    if x16 == "X16":             return 3
    elif tk_ == "5111":          return 12
    elif tk_.startswith("5"):    return 2
    else:                        return 5

_now      = datetime.now()
DATE_TO   = _now.strftime("%d.%m.%Y")
DATE_FROM = (_now.replace(day=1) - relativedelta(months=1)).strftime("01.%m.%Y")

# ⚠️  Άλλαξε αυτό κάθε φορά που αλλάζει η περίοδο 
LIMIT_DATE = pd.Timestamp(2025, 8, 1).date()

SHOW_COLS     = ["MRN", "PROT", "PDF", "PDF.1", "ΚΑΤΑΣΤΑΣΗ"]
EDIT_COLS     = {"PROT", "PDF", "PDF.1"}
ALLOWED_DASMOS = ("76", "72", "81", "2804690", "2710112")

# ── Logging (ίδιο pattern με script ΕΙΣΑΓΩΓΕΣ ΠΕΙΡΑΙΑ) ────────────────────
_SCRIPT_NAME = Path(__file__).stem
_LOG_DIR = Path(__file__).resolve().parent / "logs"
_LOG_DIR.mkdir(exist_ok=True)
_LOG_FILE = _LOG_DIR / f"{_SCRIPT_NAME}_{datetime.now():%Y%m%d_%H%M%S}.log"
logging.basicConfig(
    level=logging.WARNING,  # root: κόβει το θορυβώδες DEBUG τρίτων (selenium/urllib3 κλπ)
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.FileHandler(_LOG_FILE, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(_SCRIPT_NAME)
log.setLevel(logging.DEBUG)  # μόνο το δικό μας logger σε DEBUG
log.info(f"Log file: {_LOG_FILE}")


class _ErrorFlag(logging.Handler):
    def __init__(self):
        super().__init__(level=logging.ERROR)
        self.seen = False

    def emit(self, record):
        self.seen = True


_ERROR_FLAG = _ErrorFlag()
logging.getLogger().addHandler(_ERROR_FLAG)
_SAVED = {"n": 0}


def _finish_log() -> None:
    """Το log μένει μόνο αν έγινε σφάλμα ή σώθηκε έστω μία καταχώρηση — αλλιώς σβήνεται."""
    if _ERROR_FLAG.seen or _SAVED["n"]:
        return
    root = logging.getLogger()
    for h in root.handlers[:]:
        if isinstance(h, logging.FileHandler):
            h.close()
            root.removeHandler(h)
    _LOG_FILE.unlink(missing_ok=True)


def save_screenshot(tag: str) -> Path:
    """Screenshot για οπτικό/audit έλεγχο — καλείται μετά από κάθε SAP entry
    (επιτυχία ή σφάλμα). Φέρνει ΠΡΩΤΑ το SAP window μπροστά (foreground),
    ώστε το screenshot να δείχνει ΠΑΝΤΑ το SAP και όχι ό,τι άλλο τύχει να
    είναι ενεργό στην οθόνη τη δεδομένη στιγμή."""
    path = _LOG_DIR / f"{_SCRIPT_NAME}_{tag}_{datetime.now():%Y%m%d_%H%M%S}.png"
    try:
        Application(backend="uia").connect(
            class_name="SAP_FRONTEND_SESSION", timeout=5
        ).window(class_name="SAP_FRONTEND_SESSION").set_focus()
        time.sleep(0.3)
    except Exception as e:
        log.warning(f"Δεν βρέθηκε/έγινε focus το SAP window για screenshot: {e}")
    try:
        pyautogui.screenshot(str(path))
        log.info(f"Screenshot αποθηκεύτηκε: {path}")
    except Exception as e:
        log.exception(f"Απέτυχε το screenshot: {e}")
    return path

# ==============================================================================
# HELPERS
# ==============================================================================

def make_chrome(download_folder: Path = None):
    """Φτιάχνει Chrome driver. Αν δοθεί download_folder → auto-download mode."""
    opts = Options()
    opts.add_argument("--ignore-certificate-errors")
    opts.add_argument("--start-maximized")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    # Απενεργοποιεί άσχετες προσπάθειες σύνδεσης σε Google υπηρεσίες (GCM,
    # component updater κλπ) που μπλοκάρονται από το εταιρικό δίκτυο και
    # απλά καθυστερούν το άνοιγμα - δεν έχουν καμία σχέση με το ICISnet.
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--disable-component-update")
    opts.add_argument("--disable-sync")
    opts.add_argument("--disable-client-side-phishing-detection")
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-default-browser-check")
    if download_folder:
        opts.add_experimental_option("prefs", {
            "download.default_directory": str(download_folder),
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True,
        })
    else:
        opts.add_argument("--kiosk-printing")
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=opts)


def close_sap():
    for proc in ["saplogon.exe", "saplgpad.exe", "sapgui.exe", "nwbc.exe"]:
        subprocess.call(["taskkill", "/F", "/T", "/IM", proc],
                        stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)


def sap_login(transaction_code: str):
    """
    Login μέσω SAP GUI Scripting (OpenConnection, χωρίς pywinauto) + πλοήγηση
    στο δοσμένο transaction μέσω του πεδίου εντολών (okcd). Επιστρέφει session.
    (Ίδιο με το script ΕΙΣΑΓΩΓΕΣ ΠΕΙΡΑΙΑ — validated live.)
    """
    close_sap(); time.sleep(2)
    subprocess.Popen('start "" saplogon.exe', shell=True); time.sleep(3)

    sap_gui_auto = None
    for _ in range(30):
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            break
        except Exception:
            time.sleep(1)
    if sap_gui_auto is None:
        raise RuntimeError("SAPGUI scripting engine δεν βρέθηκε (timeout).")

    application = sap_gui_auto.GetScriptingEngine
    connection  = application.OpenConnection("YOUR_SAP_CONNECTION", True)
    session     = connection.Children(0)
    time.sleep(1)

    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "YOUR_CLIENT"
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_USER
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = SAP_PASS
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EL"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(3)

    sap_navigate(session, transaction_code)
    return session


def sap_ima_login():
    """
    Άνοιγμα SAP Logon + username/password login. ΔΕΝ μπαίνει καθόλου στο
    transaction — αυτό το κάνει το sap_ima_reenter_transaction(), που
    καλείται ξεχωριστά ΚΑΘΕ φορά ακριβώς πριν από μια καταχώρηση (ώστε το
    XML να είναι ήδη έτοιμο ΠΡΙΝ ξεκινήσει η πλοήγηση SAP, όχι μετά).
    Καλείται ΜΙΑ φορά για όλο το batch. Επιστρέφει session.
    """
    close_sap(); time.sleep(2)
    subprocess.Popen('start "" saplogon.exe', shell=True)
    time.sleep(3)

    sap_gui_auto = None
    for _ in range(30):
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            break
        except Exception:
            time.sleep(1)
    if sap_gui_auto is None:
        raise RuntimeError("SAPGUI scripting engine δεν βρέθηκε (timeout).")

    application = sap_gui_auto.GetScriptingEngine
    connection = application.OpenConnection("YOUR_SAP_CONNECTION", True)
    session = connection.Children(0)
    time.sleep(1)

    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "YOUR_CLIENT"
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_USER
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = SAP_PASS
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EL"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(3)

    return session


def sap_ima_reenter_transaction(session):
    """
    ZELVMM_IMP_1 (okcd) + Επιλογή Πεδίου (ΕΤΑΙΡΙΑ=COMPANY_CODE) + Εκτέλεση +
    Εμφάνιση→Αλλαγή. ΣΤΑΜΑΤΑΕΙ στην οθόνη "...Επισκόπηση/Αλλαγή" — ΔΕΝ
    πατάει "Νέες Καταχωρίσεις" ακόμα. Καλείται ΚΑΘΕ φορά ακριβώς πριν από
    μια καταχώρηση (και την πρώτη, και τις επόμενες μετά το διπλό F3) —
    ΠΟΤΕ πριν είναι έτοιμο το XML του MRN, πάντα μετά (επιβεβαιώθηκε ότι
    το "Επιλογή Πεδίου" εμφανίζεται ΚΑΘΕ φορά, όχι μόνο την πρώτη).
    """
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZELVMM_IMP_1"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    time.sleep(2)

    for idx in [0, 1, 2, 4, 6]:
        session.findById(
            f"wnd[1]/usr/sub:SAPLSVIX:0210/chkMARK_CHECKBOX[{idx},0]"
        ).Selected = True
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(1)

    session.findById("wnd[1]/usr/sub:SAPLSVIX:0100/ctxtD0100_FIELD_TAB-LOWER_LIMIT[6,37]").text = COMPANY_CODE
    session.findById("wnd[1]/usr/sub:SAPLSVIX:0100/ctxtD0100_FIELD_TAB-UPPER_LIMIT[7,37]").text = COMPANY_CODE

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(3)

    session.findById("wnd[0]").maximize()
    time.sleep(2)
    session.findById("wnd[0]/tbar[1]/btn[25]").press()
    time.sleep(2)


def sap_ima_open_new_entry(session):
    """Πατάει "Νέες Καταχωρίσεις" (btn[5]) + maximize. Καλείται ΓΙΑ ΚΑΘΕ
    καταχώρηση, μετά το sap_ima_reenter_transaction()."""
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    time.sleep(2)
    session.findById("wnd[0]").maximize()
    time.sleep(1)


def sap_ima_back_to_overview(session):
    """Διπλό F3 (Πίσω) — γυρνάει ΠΙΣΩ ΜΕΧΡΙ ΤΟ ΣΗΜΕΙΟ ΠΟΥ ΘΑ ΞΑΝΑΓΡΑΨΟΥΜΕ
    ZELVMM_IMP_1 (επιβεβαιώθηκε από τον χρήστη) — ίδιο pattern "διπλό F3"
    με το attach_pdf_gos() του ΕΙΣΑΓΩΓΕΣ ΠΕΙΡΑΙΑ NEW. Μετά από αυτό πρέπει
    να καλείται sap_ima_reenter_transaction(), όχι απευθείας
    sap_ima_open_new_entry()."""
    session.findById("wnd[0]/tbar[0]/btn[3]").press(); time.sleep(1)
    session.findById("wnd[0]/tbar[0]/btn[3]").press(); time.sleep(1)


def sap_navigate(session, transaction_code: str):
    """Πλοήγηση σε transaction μέσω του πεδίου εντολών (okcd)."""
    session.findById("wnd[0]/tbar[0]/okcd").text = transaction_code
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    time.sleep(2)


def setup_keyboard():
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    ctypes.windll.user32.PostMessageW(hwnd, 0x0050, 0, 0x0408)
    if ctypes.windll.user32.GetKeyState(0x14) & 1:
        pyautogui.press("capslock")


def safe_select(wait, element_id, text):
    for _ in range(5):
        try:
            Select(wait.until(EC.element_to_be_clickable((By.ID, element_id))))\
                .select_by_visible_text(text)
            return
        except:
            time.sleep(1)
    raise Exception(f"safe_select failed: {element_id}={text}")


def popup_input(title: str, prompt: str, default: str = "") -> str:
    """Custom (όχι simpledialog) ώστε να μπορούμε να το φέρουμε πάντα μπροστά
    με topmost/lift/focus_force πάνω στο ΙΔΙΟ ορατό window — το simpledialog
    πάνω σε withdrawn root δεν ερχόταν πάντα μπροστά από το SAP."""
    root = tk.Tk()
    root.title(title)
    root.resizable(False, False)
    root.configure(bg="#F5F4F0")
    w, h = 420, 160
    sw = root.winfo_screenwidth(); sh = root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
    root.attributes("-topmost", True)

    result = {"value": None}

    tk.Label(root, text=prompt,
        bg="#F5F4F0", fg="#1A1A1A", font=("Consolas", 10),
        wraplength=380, justify="center"
    ).pack(pady=(20, 10))

    entry = tk.Entry(root, font=("Consolas", 11), justify="center")
    entry.insert(0, default)
    entry.pack(pady=(0, 15), ipady=3, padx=40, fill="x")

    def submit(event=None):
        result["value"] = entry.get()
        root.destroy()

    tk.Button(root, text="OK", font=("Consolas", 10, "bold"),
        bg="#1A1A1A", fg="#FFFFFF", relief="flat", padx=20, pady=6,
        cursor="hand2", command=submit
    ).pack()

    entry.bind("<Return>", submit)
    root.lift(); root.focus_force()
    entry.focus_set()
    root.mainloop()

    return result["value"].strip() if result["value"] else ""

def show_info(msg: str):
    """Εμφανίζει ενημερωτικό banner με OK."""
    root = tk.Tk()
    root.title("Ενημέρωση")
    root.resizable(False, False)
    root.configure(bg="#F5F4F0")
    root.attributes("-topmost", True)
    w, h = 420, 150
    sw = root.winfo_screenwidth(); sh = root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
    root.lift(); root.focus_force()

    tk.Label(root, text=msg,
        bg="#F5F4F0", fg="#1A1A1A", font=("Consolas", 10),
        wraplength=380, justify="center"
    ).pack(pady=(25, 15))

    tk.Button(root, text="OK", font=("Consolas", 10, "bold"),
        bg="#1A1A1A", fg="#FFFFFF", relief="flat", padx=20, pady=6,
        cursor="hand2", command=root.destroy
    ).pack()

    root.mainloop()

# ==============================================================================
# ΦΑΣΗ Α.1 — ICISnet scraping
# ==============================================================================

def phase_a_icisnet() -> pd.DataFrame:
    for attempt in range(1, 4):
        print(f"  ICISnet — απόπειρα {attempt}/3...")
        driver = None
        try:
            driver = make_chrome()
            wait   = WebDriverWait(driver, 45)

            driver.get("https://www1.gsis.gr/icisnet/itrader/common/home.jsf")
            wait.until(EC.element_to_be_clickable((By.NAME, "username"))).send_keys(ICISNET_USER)
            driver.find_element(By.NAME, "password").send_keys(ICISNET_PASS)
            driver.find_element(By.NAME, "btn_login").click()
            wait.until(EC.url_contains("icisnet"))

            wait.until(EC.element_to_be_clickable((By.ID, "iconcontentForm:mainMenu_ics"))).click()
            wait.until(EC.element_to_be_clickable((By.ID, "iconcontentForm:ics_import_declaration"))).click()
            wait.until(EC.element_to_be_clickable((By.ID, "iconcontentForm:menu_ics_import_declaration_search"))).click()

            df_el = wait.until(EC.element_to_be_clickable(
                (By.ID, "contentForm:submission_date_fromInputDate")))
            driver.execute_script(
                "arguments[0].removeAttribute('readonly'); arguments[0].value=arguments[1];",
                df_el, DATE_FROM.replace(".", "-"))
            Select(driver.find_element(By.ID, "contentForm:search_scope")).select_by_value("Trader")

            t_search = perf_counter()
            for sa in range(1, 11):
                try:
                    driver.find_element(By.XPATH, "//input[@value='Αναζήτηση']").click()
                    WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.ID, "contentForm:printResultsReport")))
                    break
                except:
                    if sa == 10: raise
                    time.sleep(1)
            print(f"    [χρόνος] Αναζήτηση -> αποτελέσματα έτοιμα: {fmt_duration(perf_counter() - t_search)}")

            t_popup = perf_counter()
            driver.execute_script(
                "jsfcljs(document.getElementById('contentForm'),{"
                "'contentForm:printResultsReport':'contentForm:printResultsReport',"
                "'showLrn':'true','dispatch':'','movementReferenceLabel':'MRN',"
                "'noLRNcol':'false'},'new');")
            wait.until(lambda d: len(d.window_handles) > 1)
            driver.switch_to.window(driver.window_handles[-1])
            wait.until(EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(),'Αποτελέσματα Αναζήτησης')]")))
            print(f"    [χρόνος] Άνοιγμα popup 'Αποτελέσματα Αναζήτησης': {fmt_duration(perf_counter() - t_popup)}")

            t_pdf = perf_counter()
            try:
                pdf_data = driver.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
                PDF_SAVE_PATH.write_bytes(base64.b64decode(pdf_data["data"]))
                print(f"    PDF saved  |  [χρόνος] Page.printToPDF: {fmt_duration(perf_counter() - t_pdf)}")
            except Exception as e:
                print(f"    PDF save failed: {e}")

            try:
                df_raw = pd.read_html(StringIO(driver.page_source))[0]
            except ValueError:
                driver.quit()
                print("    ICISnet: δεν βρέθηκαν εισαγωγές")
                return pd.DataFrame(columns=["MRN","ΤΥΠΟΣ","ΚΑΤΑΣΤΑΣΗ","LRN","ΗΜ_ΥΠΟΒ","ΗΜ_ΕΝΗΜ","PDF"])

            driver.quit(); driver = None

            # Καθαρισμός & φιλτράρισμα
            df = df_raw.copy()
            df.columns = [str(c).strip() for c in df.columns]
            repl = [
                ("YOUR_BROKER_ID/25/", "ELVYOUR_BROKER_ID/25/"), ("ELVELV", "ELV"),
                ("YOUR_BROKER_ID/25 /", "ELVYOUR_BROKER_ID/25 /"),
                ("YOUR_BROKER_ID2/26/131ELB", "YOUR_BROKER_ID2/26/131ELV"),
                ("CBRM", "CBELVRM"), ("CB78-2023", "ELVCB78-2023"),
            ]
            df["LRN"] = df["LRN"].astype(str)
            for old, new in repl:
                df["LRN"] = df["LRN"].str.replace(old, new, regex=False)
            df["Τύπος Δήλωσης"] = df["Τύπος Δήλωσης"].astype(str).str.replace("-","",regex=False)
            df["Ημ/νία Υποβολής"] = pd.to_datetime(
                df["Ημ/νία Υποβολής"], dayfirst=True, errors="coerce").dt.date
            df["Ημ/νία Ενημέρωσης Κατάστασης"] = pd.to_datetime(
                df["Ημ/νία Ενημέρωσης Κατάστασης"], dayfirst=True, errors="coerce").dt.date
            df = df[~df["LRN"].str.contains("XALELV", na=False)]
            df = df[
                df["LRN"].str.contains(r"ELV|ΕLV", na=False) |
                df["LRN"].str.contains("YOUR_SPECIAL_LRN", na=False) |
                df["MRN"].isin(["YOUR_MRN_1","YOUR_MRN_2"])
            ]
            df = df[
                (df["Ημ/νία Υποβολής"] >= LIMIT_DATE) &
                (df["Ημ/νία Ενημέρωσης Κατάστασης"] >= LIMIT_DATE)
            ]
            df["Κατάσταση_Temp"] = df["Κατάσταση"].astype(str).str.replace(
                r"Εισαγωγή.*","ID29", regex=True)
            cond = (
                df["Κατάσταση_Temp"].str.startswith("ID29", na=False) |
                (df["Κατάσταση"] == "Τακτοποιημένο") |
                df["Κατάσταση"].str.contains("Αποδεκτή", na=False) |
                (df["MRN"] == "YOUR_MRN_3")
            )
            df = df[cond]
            df["Κατάσταση"] = df["Κατάσταση_Temp"]
            df["PDF"] = df["MRN"].astype(str) + " " + df["Τύπος Δήλωσης"].astype(str)
            df = df.rename(columns={
                "Τύπος Δήλωσης":"ΤΥΠΟΣ","Κατάσταση":"ΚΑΤΑΣΤΑΣΗ",
                "Ημ/νία Υποβολής":"ΗΜ_ΥΠΟΒ","Ημ/νία Ενημέρωσης Κατάστασης":"ΗΜ_ΕΝΗΜ"})
            df_final = df[["MRN","ΤΥΠΟΣ","ΚΑΤΑΣΤΑΣΗ","LRN","ΗΜ_ΥΠΟΒ","ΗΜ_ΕΝΗΜ","PDF"]]\
                .sort_values(["ΗΜ_ΥΠΟΒ","MRN"], ascending=[False,True]).reset_index(drop=True)
            print(f"    ICISnet OK — {len(df_final)} εγγραφές")
            return df_final

        except Exception as e:
            if driver:
                try: driver.quit()
                except: pass
            print(f"    Απέτυχε: {e}")
            time.sleep(3)

    raise RuntimeError("ICISnet scraping απέτυχε μετά από 3 απόπειρες.")


# ==============================================================================
# ΦΑΣΗ Α.2 — SAP export
# ==============================================================================

def phase_a_sap_export():
    """
    SAP export μέσω SAP GUI Scripting — login και export χωρίς pywinauto.
    Ίδιο transaction/φίλτρα/export-μενού με το phase_a_sap_export του
    script ΕΙΣΑΓΩΓΕΣ ΠΕΙΡΑΙΑ (ήδη validated live).
    """
    for attempt in range(1, 4):
        print(f"  SAP Export — απόπειρα {attempt}/3...")
        try:
            session = sap_login("ZELVMM_IMP_1_LIST_N")

            session.findById("wnd[0]/usr/ctxtS_ZDATE-LOW").text = DATE_FROM
            session.findById("wnd[0]/usr/ctxtS_ZDATE-HIGH").text = DATE_TO
            session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = COMPANY_CODE
            session.findById("wnd[0]/tbar[1]/btn[8]").press()   # Εκτέλεση
            time.sleep(10)

            # ── Export μέσω μενού Λίστα → Εξαγωγή → Υπολογιστικό φύλλο ──────
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(DESKTOP)
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ΕΙΣΑΓΩΓΕΣ_database.xlsx"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            time.sleep(5)

            subprocess.call(["taskkill","/F","/IM","excel.exe"],
                            stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)
            time.sleep(2); close_sap()
            print("    SAP Export OK"); return

        except Exception as e:
            print(f"    Απέτυχε: {e}"); time.sleep(5)

    raise RuntimeError("SAP Export απέτυχε μετά από 3 απόπειρες.")

# ==============================================================================
# ΦΑΣΗ Α.3 — Queries → FULL_RESULTS.xlsx
# ==============================================================================

def phase_a_queries(df_final: pd.DataFrame) -> pd.DataFrame:
    print("  Queries...")
    df_ger = pd.read_excel(BROKER_XLSX, sheet_name=0)
    df_db  = pd.read_excel(DATABASE_XLSX, sheet_name=0)

    df_ger["MRN"]       = df_ger["MRN"].astype(str)
    df_ger["PROT"]      = df_ger["PROT"].astype(str)
    df_ger["TYPE"]      = df_ger["TYPE"].fillna("").astype(str)
    df_ger["ABC"]       = df_ger["ABC"].fillna("").astype(str)
    df_ger["KATH"]      = df_ger["KATH"].fillna("").astype(str).str.replace(" ","",regex=False)
    df_ger["TYPE_FULL"] = df_ger["TYPE"] + df_ger["ABC"]
    df_ger["PDF"]       = (df_ger["MRN"] + " " + df_ger["TYPE_FULL"]).str.strip()
    df_ger = df_ger.drop_duplicates(subset=["MRN","PROT","TYPE_FULL","PDF"])

    df_db["Καθεστώς Εισαγωγής"] = pd.to_numeric(df_db["Καθεστώς Εισαγωγής"], errors="coerce")
    df_sap_exp = (
        df_db[df_db["Καθεστώς Εισαγωγής"].notna() &
              ~df_db["Καθεστώς Εισαγωγής"].isin([7,4])]
        .rename(columns={"Διασάφηση Εισαγωγής":"MRN"})[["MRN","Ημ/νία δημιουρ."]]
    )
    pdf_files = (
        [f.replace(".pdf","") for f in os.listdir(SAP_GUI_DIR) if "GRIM" in f]
        if SAP_GUI_DIR.exists() else []
    )
    df_gui = pd.DataFrame({"PDF NAME": pdf_files})
    df_gui["PDF NAME"] = df_gui["PDF NAME"].astype(str)

    if df_gui.empty:
        show_info("⚠️  Δεν υπάρχουν PDF στο SAP GUI folder.\nΚαμία διασάφηση προς καταχώρηση.")
        df_gui = pd.DataFrame({"PDF NAME": pd.Series([], dtype=str)})

    df_full = df_db.copy()
    df_full["Τελωνείο Εισόδου"] = df_full["Τελωνείο Εισόδου"].astype(str)
    df_full = df_full[
        (df_full["Τελωνείο Εισόδου"] != "0") &
        df_full["Καθεστώς Εισαγωγής"].notna() &
        (df_full["Καθεστώς Εισαγωγής"] != 8)
    ][["Διασάφηση Εισαγωγής","Διασάφηση Εισόδου"]]
    df_full = df_full[
        ~df_full["Διασάφηση Εισαγωγής"].astype(str).str.startswith("23GRIM",na=False)]

    # opened
    o = df_final.copy()
    o = o.merge(df_sap_exp[["MRN"]], on="MRN", how="left", indicator=True)
    o = o[o["_merge"]=="left_only"].drop(columns=["_merge"])
    o = o[~o["ΤΥΠΟΣ"].astype(str).str.endswith(("X","Y","Χ","Υ"),na=False)]
    o = o[~o["MRN"].astype(str).str.contains("GRIM0304",na=False)]
    o = o[~o["MRN"].isin(["YOUR_MRN_4","YOUR_MRN_5"])]
    o = o.merge(df_ger[["MRN","PROT","KATH","PDF"]], on="MRN", how="left", suffixes=("",".1"))
    o = o[o["PDF.1"].notna()]
    o = o[o["KATH"].astype(str) != "7100"]
    o = o.merge(df_full[["Διασάφηση Εισόδου"]], left_on="MRN",
                right_on="Διασάφηση Εισόδου", how="left")
    o = o[o["Διασάφηση Εισόδου"].isna()].drop(columns=["Διασάφηση Εισόδου"])
    o = o.merge(df_gui[["PDF NAME"]], left_on="PDF", right_on="PDF NAME", how="left")
    o["OK"] = o.apply(
        lambda x: "OK" if (pd.notna(x["PDF.1"]) and pd.notna(x["PDF NAME"])
                           and x["PDF"]==x["PDF.1"]==x["PDF NAME"])
        else (None if pd.isna(x["PDF.1"]) else "CHANGE"), axis=1)
    df_opened = o[["MRN","ΤΥΠΟΣ","ΚΑΤΑΣΤΑΣΗ","LRN","ΗΜ_ΥΠΟΒ","ΗΜ_ΕΝΗΜ","PROT","PDF","PDF.1","OK"]]\
        .sort_values(["OK","ΗΜ_ΥΠΟΒ"], ascending=[False,False])

    # done
    d = df_final.copy()
    d = d.merge(df_sap_exp[["MRN"]], on="MRN", how="left", indicator=True)
    d = d.merge(df_full, left_on="MRN", right_on="Διασάφηση Εισόδου", how="left")
    d["NEW_MRN"] = d.apply(
        lambda x: x["MRN"] if x["_merge"]=="both" else x["Διασάφηση Εισαγωγής"], axis=1)
    df_done = d[d["NEW_MRN"].notna()].drop(
        columns=["_merge","Διασάφηση Εισαγωγής","Διασάφηση Εισόδου","NEW_MRN"], errors="ignore")

    # undone
    u = df_final.copy()
    u = u[~u["ΤΥΠΟΣ"].astype(str).str.endswith(("X","Y","Χ","Υ"),na=False)]
    u = u.merge(df_sap_exp[["MRN"]], on="MRN", how="left", indicator=True)
    u = u.merge(df_full, left_on="MRN", right_on="Διασάφηση Εισόδου", how="left")
    u["NEW_MRN"] = u.apply(
        lambda x: x["MRN"] if x["_merge"]=="both" else x["Διασάφηση Εισαγωγής"], axis=1)
    df_undone = u[u["NEW_MRN"].isna()].drop(
        columns=["_merge","Διασάφηση Εισαγωγής","Διασάφηση Εισόδου","NEW_MRN"], errors="ignore")

    # Export FULL_RESULTS.xlsx
    OUTPUT_EXCEL.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        for sname, df in {"opened":df_opened,"done":df_done,"undone":df_undone}.items():
            df.to_excel(writer, sheet_name=sname, index=False)
            ws = writer.sheets[sname]; rc, cc = df.shape
            for ci, col in enumerate(ws.columns, 1):
                mx = 0; cl = get_column_letter(ci)
                for cell in col:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    try:
                        if cell.value: mx = max(mx, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[cl].width = mx + 3
            if rc > 0:
                ref = f"A1:{get_column_letter(cc)}{rc+1}"
                tab = Table(displayName=sname, ref=ref)
                tab.tableStyleInfo = TableStyleInfo(
                    name="TableStyleMedium7", showFirstColumn=False,
                    showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                ws.add_table(tab)

    print(f"    FULL_RESULTS.xlsx — opened={len(df_opened)} | done={len(df_done)} | undone={len(df_undone)}")
    return df_opened

# ==============================================================================
# POPUP — Editable approval table
# ==============================================================================

# Στήλες που βλέπεις στο popup
SHOW_COLS = ["MRN", "PROT", "PDF", "PDF.1", "ΚΑΤΑΣΤΑΣΗ",
             "ΚΡΑΜΑ_1", "ΒΑΡΟΣ_1", "ΚΡΑΜΑ_2", "ΒΑΡΟΣ_2", "ΚΡΑΜΑ_3", "ΒΑΡΟΣ_3"]

# Στήλες που μπορείς να επεξεργαστείς
EDIT_COLS = {"PROT", "PDF", "PDF.1",
             "ΚΡΑΜΑ_1", "ΒΑΡΟΣ_1", "ΚΡΑΜΑ_2", "ΒΑΡΟΣ_2", "ΚΡΑΜΑ_3", "ΒΑΡΟΣ_3"}


class ApprovalPopup:
    """
    Πρώτο popup — pending MRNs από το FULL_RESULTS.
    Για 0832 + ΔΑΣΜ_ΚΛ 76012080/76012030 συμπληρώνεις ΚΡΑΜΑ/ΒΑΡΟΣ.
    Double-click σε editable στήλη → inline edit.
    Επιστρέφει list[dict] ή None αν ακυρωθεί.
    """

    def __init__(self, rows: list):
        self.rows         = [r.copy() for r in rows]
        self.result       = None
        self._edit_widget = None
        self.root = tk.Tk()
        self.root.title("Έλεγχος Εισαγωγών — Επιλογή & Επεξεργασία")
        self.root.resizable(True, True)
        w, h = 1400, 520
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.root.configure(bg="#F5F4F0")
        self._build()
        self.root.lift()  # ← νέο
        self.root.focus_force()  # ← νέο
        self.root.attributes("-topmost", True)  # ← νέο

    def _build(self):
        r = self.root

        # Header
        hdr = tk.Frame(r, bg="#1A1A1A", height=52)
        hdr.pack(fill="x")
        tk.Label(hdr,
            text="  Pending Εισαγωγές — έλεγξε, τροποποίησε αν χρειαστεί, πάτα OK",
            bg="#1A1A1A", fg="#FFFFFF", font=("Consolas", 11), anchor="w"
        ).pack(side="left", padx=8, pady=14)
        tk.Label(hdr, text=f"{len(self.rows)} MRNs",
            bg="#1A1A1A", fg="#8A8A82", font=("Consolas", 10)
        ).pack(side="right", padx=16)

        # Hint
        hint = tk.Frame(r, bg="#F5F4F0")
        hint.pack(fill="x", padx=14, pady=(8, 2))
        tk.Label(hint,
            text="  Double-click για επεξεργασία   |   "
                 "Για 0832: συμπλήρωσε ΚΡΑΜΑ/ΒΑΡΟΣ (αν 1 κράμα αφησε ΒΑΡΟΣ κενό)",
            bg="#F5F4F0", fg="#6B6B65", font=("Consolas", 9), anchor="w"
        ).pack(side="left")

        # Treeview
        frm = tk.Frame(r, bg="#F5F4F0")
        frm.pack(fill="both", expand=True, padx=14, pady=(4, 6))

        style = ttk.Style(); style.theme_use("clam")
        style.configure("T.Treeview",
            background="#FFFFFF", foreground="#1A1A1A", rowheight=28,
            fieldbackground="#FFFFFF", font=("Consolas", 10), borderwidth=0)
        style.configure("T.Treeview.Heading",
            background="#E8E6DF", foreground="#3A3A36",
            font=("Consolas", 10, "bold"), relief="flat")
        style.map("T.Treeview", background=[("selected", "#D4E8FF")])

        vsb = ttk.Scrollbar(frm, orient="vertical")
        hsb = ttk.Scrollbar(frm, orient="horizontal")
        all_cols = ["_check"] + SHOW_COLS
        self.tree = ttk.Treeview(frm, columns=all_cols, show="headings",
            style="T.Treeview", yscrollcommand=vsb.set, xscrollcommand=hsb.set,
            selectmode="browse")
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        col_w = {
            "_check": 40, "MRN": 200, "PROT": 70, "PDF": 180, "PDF.1": 180,
            "ΚΑΤΑΣΤΑΣΗ": 90, "ΚΡΑΜΑ_1": 70, "ΒΑΡΟΣ_1": 70,
            "ΚΡΑΜΑ_2": 70, "ΒΑΡΟΣ_2": 70, "ΚΡΑΜΑ_3": 70, "ΒΑΡΟΣ_3": 70,
        }
        col_l = {
            "_check": "ok", "MRN": "MRN", "PROT": "PROT",
            "PDF": "PDF (target)", "PDF.1": "PDF.1 (source)", "ΚΑΤΑΣΤΑΣΗ": "ΚΑΤΑΣΤΑΣΗ",
            "ΚΡΑΜΑ_1": "ΚΡ_1", "ΒΑΡΟΣ_1": "ΒΑΡ_1",
            "ΚΡΑΜΑ_2": "ΚΡ_2", "ΒΑΡΟΣ_2": "ΒΑΡ_2",
            "ΚΡΑΜΑ_3": "ΚΡ_3", "ΒΑΡΟΣ_3": "ΒΑΡ_3",
        }
        for c in all_cols:
            self.tree.heading(c, text=col_l[c])
            self.tree.column(c, width=col_w[c],
                anchor="center" if c == "_check" else "w",
                stretch=(c in ("MRN", "PDF", "PDF.1")))

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frm.grid_rowconfigure(0, weight=1); frm.grid_columnconfigure(0, weight=1)

        self._cvars = []; self._iids = []
        for i, row in enumerate(self.rows):
            var = tk.BooleanVar(value=True); self._cvars.append(var)
            vals = ["[X]"] + [str(row.get(c, "")) for c in SHOW_COLS]
            iid = self.tree.insert("", "end", iid=str(i), values=vals,
                                   tags=("odd" if i % 2 else "even",))
            self._iids.append(iid)
        self.tree.tag_configure("odd",  background="#FAFAF8")
        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.bind("<Button-1>",        self._on_click)
        self.tree.bind("<Double-Button-1>", self._on_dbl)

        # Footer
        foot = tk.Frame(r, bg="#F5F4F0")
        foot.pack(fill="x", padx=14, pady=(0, 12))
        tk.Button(foot, text="Επιλογή Όλων", font=("Consolas", 9),
            bg="#E8E6DF", fg="#1A1A1A", relief="flat", padx=10, pady=5,
            cursor="hand2", command=self._all
        ).pack(side="left", padx=(0, 6))
        tk.Button(foot, text="Αποεπιλογή Όλων", font=("Consolas", 9),
            bg="#E8E6DF", fg="#1A1A1A", relief="flat", padx=10, pady=5,
            cursor="hand2", command=self._none
        ).pack(side="left")
        self._cnt = tk.Label(foot, text="", bg="#F5F4F0", fg="#6B6B65", font=("Consolas", 9))
        self._cnt.pack(side="left", padx=14); self._upd()
        tk.Button(foot, text="Ακύρωση", font=("Consolas", 10),
            bg="#E8E6DF", fg="#6B6B65", relief="flat", padx=14, pady=7,
            cursor="hand2", command=self._cancel
        ).pack(side="right", padx=(6, 0))
        tk.Button(foot, text="OK — Εκκίνηση", font=("Consolas", 10, "bold"),
            bg="#1A1A1A", fg="#FFFFFF", relief="flat", padx=14, pady=7,
            cursor="hand2", activebackground="#333", activeforeground="#FFF",
            command=self._ok
        ).pack(side="right")

    def _on_click(self, event):
        if (self.tree.identify_column(event.x) == "#1" and
                self.tree.identify_region(event.x, event.y) == "cell"):
            iid = self.tree.identify_row(event.y)
            if not iid: return
            self._commit()
            idx = self._iids.index(iid)
            nv  = not self._cvars[idx].get(); self._cvars[idx].set(nv)
            v   = list(self.tree.item(iid, "values")); v[0] = "[X]" if nv else "[ ]"
            self.tree.item(iid, values=v); self._upd()

    def _on_dbl(self, event):
        col_id = self.tree.identify_column(event.x)
        iid    = self.tree.identify_row(event.y)
        if not iid: return
        ci = int(col_id.replace("#", "")) - 1
        ac = ["_check"] + SHOW_COLS
        if ci < 0 or ci >= len(ac): return
        cn = ac[ci]
        if cn not in EDIT_COLS: return
        self._commit()
        bbox = self.tree.bbox(iid, column=col_id)
        if not bbox: return
        x, y, w, h = bbox
        vals = list(self.tree.item(iid, "values"))
        e = tk.Entry(self.tree, font=("Consolas", 10), bg="#FFFDE7",
            fg="#1A1A1A", relief="solid", bd=1, insertbackground="#1A1A1A")
        e.place(x=x, y=y, width=w, height=h)
        e.insert(0, vals[ci]); e.select_range(0, "end"); e.focus_set()
        def commit(ev=None):
            vals[ci] = e.get().strip()
            self.tree.item(iid, values=vals)
            self.rows[self._iids.index(iid)][cn] = vals[ci]
            e.destroy(); self._edit_widget = None
        e.bind("<Return>", commit); e.bind("<Tab>", commit)
        e.bind("<Escape>", lambda ev: e.destroy()); e.bind("<FocusOut>", commit)
        self._edit_widget = e

    def _commit(self):
        if self._edit_widget:
            try: self._edit_widget.event_generate("<FocusOut>")
            except: pass

    def _upd(self):
        n = sum(v.get() for v in self._cvars)
        self._cnt.config(text=f"{n} / {len(self.rows)} επιλεγμένα")

    def _all(self):
        for i, v in enumerate(self._cvars):
            v.set(True)
            vs = list(self.tree.item(self._iids[i], "values")); vs[0] = "[X]"
            self.tree.item(self._iids[i], values=vs)
        self._upd()

    def _none(self):
        for i, v in enumerate(self._cvars):
            v.set(False)
            vs = list(self.tree.item(self._iids[i], "values")); vs[0] = "[ ]"
            self.tree.item(self._iids[i], values=vs)
        self._upd()

    def _ok(self):
        self._commit()
        sel = [self.rows[i] for i, v in enumerate(self._cvars) if v.get()]
        if not sel:
            messagebox.showwarning("Καμία επιλογή",
                "Τικάρισε τουλάχιστον ένα MRN.", parent=self.root)
            return
        self.result = sel; self.root.destroy()

    def _cancel(self):
        if messagebox.askyesno("Ακύρωση", "Σίγουρα θέλεις να ακυρώσεις;", parent=self.root):
            self.root.destroy()

    def run(self) -> list | None:
        self.root.mainloop(); return self.result

# ==============================================================================
# POPUP — Items popup (για 0832 με 2+ είδη)
# ==============================================================================

class ItemsPopup:
    """
    Εμφανίζεται ΜΟΝΟ για 0832 + ΔΑΣΜ_ΚΛ 76012080/76012030 με 2+ είδη.
    Δείχνει τα είδη από το XML και ζητά ΚΡΑΜΑ_1/ΒΑΡΟΣ_1 κλπ ανά είδος.
    Επιστρέφει dict {aa: {ΚΡΑΜΑ_1, ΒΑΡΟΣ_1, ΚΡΑΜΑ_2, ΒΑΡΟΣ_2, ΚΡΑΜΑ_3, ΒΑΡΟΣ_3}}
    ή None αν ακυρωθεί.
    """

    ITEM_SHOW_COLS = ["Α/Α", "ΔΑΣΜ_ΚΛ", "ΒΑΡΟΣ",
                      "ΚΡΑΜΑ_1", "ΒΑΡΟΣ_1", "ΚΡΑΜΑ_2", "ΒΑΡΟΣ_2", "ΚΡΑΜΑ_3", "ΒΑΡΟΣ_3"]
    ITEM_EDIT_COLS = {"ΚΡΑΜΑ_1", "ΒΑΡΟΣ_1", "ΚΡΑΜΑ_2", "ΒΑΡΟΣ_2", "ΚΡΑΜΑ_3", "ΒΑΡΟΣ_3"}

    def __init__(self, df_result: pd.DataFrame, mrn: str):
        self.df      = df_result.copy()
        self.mrn     = mrn
        self.result  = None
        self._edit_widget = None
        self._commit_fn = None

        # Φτιάχνουμε dict με τιμές ανά Α/Α
        self.data = {}
        for _, row in self.df.iterrows():
            aa = str(row["Α/Α"])
            self.data[aa] = {
                "Α/Α":     aa,
                "ΔΑΣΜ_ΚΛ": str(row["ΔΑΣΜ_ΚΛ"]),
                "ΒΑΡΟΣ":   str(row["ΒΑΡΟΣ"]),
                "ΚΡΑΜΑ_1": "", "ΒΑΡΟΣ_1": "",
                "ΚΡΑΜΑ_2": "", "ΒΑΡΟΣ_2": "",
                "ΚΡΑΜΑ_3": "", "ΒΑΡΟΣ_3": "",
            }

        self.root = tk.Tk()
        self.root.title(f"Κράματα ανά Είδος — MRN: {mrn}")
        self.root.resizable(True, True)
        w, h = 1200, 600
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.root.configure(bg="#F5F4F0")
        self._build()
        self.root.lift()  # ← νέο
        self.root.focus_force()  # ← νέο
        self.root.attributes("-topmost", True)  # ← νέο

    def _build(self):
        r = self.root

        # Header
        hdr = tk.Frame(r, bg="#1A1A1A", height=52)
        hdr.pack(fill="x")
        tk.Label(hdr,
            text=f"  Ορισμός Κραμάτων ανά Είδος — MRN: {self.mrn}",
            bg="#1A1A1A", fg="#FFFFFF", font=("Consolas", 11), anchor="w"
        ).pack(side="left", padx=8, pady=14)

        # Hint
        hint = tk.Frame(r, bg="#F5F4F0")
        hint.pack(fill="x", padx=14, pady=(8, 2))
        tk.Label(hint,
            text="  Double-click για επεξεργασία   |   "
                 "Αν 1 κράμα: γράψε μόνο ΚΡΑΜΑ_1, άφησε ΒΑΡΟΣ_1 κενό   |   "
                 "Αν 2-3 κράματα: συμπλήρωσε ΚΡΑΜΑ + ΒΑΡΟΣ (άθροισμα = συνολικό βάρος είδους)",
            bg="#F5F4F0", fg="#6B6B65", font=("Consolas", 9), anchor="w"
        ).pack(side="left")

        # Treeview
        frm = tk.Frame(r, bg="#F5F4F0")
        frm.pack(fill="both", expand=True, padx=14, pady=(4, 6))

        style = ttk.Style(); style.theme_use("clam")
        style.configure("T2.Treeview",
            background="#FFFFFF", foreground="#1A1A1A", rowheight=28,
            fieldbackground="#FFFFFF", font=("Consolas", 10), borderwidth=0)
        style.configure("T2.Treeview.Heading",
            background="#E8E6DF", foreground="#3A3A36",
            font=("Consolas", 10, "bold"), relief="flat")
        style.map("T2.Treeview", background=[("selected", "#D4E8FF")])

        vsb = ttk.Scrollbar(frm, orient="vertical")
        hsb = ttk.Scrollbar(frm, orient="horizontal")
        self.tree = ttk.Treeview(frm, columns=self.ITEM_SHOW_COLS, show="headings",
            style="T2.Treeview", yscrollcommand=vsb.set, xscrollcommand=hsb.set,
            selectmode="browse")
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        col_w = {
            "Α/Α": 50, "ΔΑΣΜ_ΚΛ": 100, "ΒΑΡΟΣ": 90,
            "ΚΡΑΜΑ_1": 70, "ΒΑΡΟΣ_1": 80,
            "ΚΡΑΜΑ_2": 70, "ΒΑΡΟΣ_2": 80,
            "ΚΡΑΜΑ_3": 70, "ΒΑΡΟΣ_3": 80,
        }
        for c in self.ITEM_SHOW_COLS:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=col_w[c], anchor="w",
                stretch=(c in ("ΔΑΣΜ_ΚΛ", "ΒΑΡΟΣ")))

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frm.grid_rowconfigure(0, weight=1); frm.grid_columnconfigure(0, weight=1)

        self._iids = []
        for aa, d in self.data.items():
            vals = [d[c] for c in self.ITEM_SHOW_COLS]
            iid = self.tree.insert("", "end", iid=aa, values=vals)
            self._iids.append(iid)

        self.tree.bind("<Double-Button-1>", self._on_dbl)

        # Footer
        foot = tk.Frame(r, bg="#F5F4F0")
        foot.pack(fill="x", padx=14, pady=(0, 12))
        tk.Button(foot, text="Ακύρωση", font=("Consolas", 10),
            bg="#E8E6DF", fg="#6B6B65", relief="flat", padx=14, pady=7,
            cursor="hand2", command=self._cancel
        ).pack(side="right", padx=(6, 0))
        tk.Button(foot, text="OK — Συνέχεια", font=("Consolas", 10, "bold"),
            bg="#1A1A1A", fg="#FFFFFF", relief="flat", padx=14, pady=7,
            cursor="hand2", activebackground="#333", activeforeground="#FFF",
            command=self._ok
        ).pack(side="right")

    def _on_dbl(self, event):
        col_id = self.tree.identify_column(event.x)
        iid    = self.tree.identify_row(event.y)
        if not iid: return
        ci = int(col_id.replace("#", "")) - 1
        if ci < 0 or ci >= len(self.ITEM_SHOW_COLS): return
        cn = self.ITEM_SHOW_COLS[ci]
        if cn not in self.ITEM_EDIT_COLS: return
        self._commit()
        bbox = self.tree.bbox(iid, column=col_id)
        if not bbox: return
        x, y, w, h = bbox
        vals = list(self.tree.item(iid, "values"))
        e = tk.Entry(self.tree, font=("Consolas", 10), bg="#FFFDE7",
            fg="#1A1A1A", relief="solid", bd=1, insertbackground="#1A1A1A")
        e.place(x=x, y=y, width=w, height=h)
        e.insert(0, vals[ci]); e.select_range(0, "end"); e.focus_set()
        committed = {"done": False}
        def commit(ev=None):
            # Guard: αποτρέπει διπλή εκτέλεση (π.χ. FocusOut μετά από ήδη-συγχρονισμένο commit)
            if committed["done"]:
                return
            committed["done"] = True
            vals[ci] = e.get().strip()
            self.tree.item(iid, values=vals)
            self.data[iid][cn] = vals[ci]
            try: e.destroy()
            except tk.TclError: pass
            if self._edit_widget is e:
                self._edit_widget = None
                self._commit_fn = None
        e.bind("<Return>", commit); e.bind("<Tab>", commit)
        e.bind("<Escape>", lambda ev: e.destroy()); e.bind("<FocusOut>", commit)
        self._edit_widget = e
        self._commit_fn = commit

    def _commit(self):
        # Συγχρονισμένο commit του τρέχοντος ανοιχτού κελιού - ΟΧΙ μέσω
        # event_generate (που είναι ασύγχρονο by default και προκαλούσε
        # απώλεια τιμών σε γρήγορα διαδοχικά double-click).
        if self._commit_fn:
            try: self._commit_fn()
            except tk.TclError: pass
        self._edit_widget = None
        self._commit_fn = None

    def _ok(self):
        self._commit()
        # Validation ανά είδος
        errors = []
        for aa, d in self.data.items():
            kr1 = d["ΚΡΑΜΑ_1"].strip()
            if not kr1:
                errors.append(f"Α/Α {aa}: ΚΡΑΜΑ_1 είναι κενό!")
                continue

            # Αν έχει ΚΡΑΜΑ_2, πρέπει να έχει και ΒΑΡΟΣ_1 + ΒΑΡΟΣ_2
            kr2 = d["ΚΡΑΜΑ_2"].strip()
            kr3 = d["ΚΡΑΜΑ_3"].strip()
            has_multiple = bool(kr2)

            if has_multiple:
                # Validation αθροίσματος
                try:
                    # Παίρνουμε το συνολικό βάρος (αντικατάσταση κόμμα→τελεία)
                    total = float(d["ΒΑΡΟΣ"].replace(",", "."))
                    b1 = float(d["ΒΑΡΟΣ_1"].replace(",", ".")) if d["ΒΑΡΟΣ_1"].strip() else 0
                    b2 = float(d["ΒΑΡΟΣ_2"].replace(",", ".")) if d["ΒΑΡΟΣ_2"].strip() else 0
                    b3 = float(d["ΒΑΡΟΣ_3"].replace(",", ".")) if d["ΒΑΡΟΣ_3"].strip() else 0
                    s  = b1 + b2 + (b3 if kr3 else 0)
                    if abs(s - total) > 0.01:
                        errors.append(
                            f"Α/Α {aa}: άθροισμα βαρών ({s}) ≠ συνολικό βάρος ({total})!")
                except ValueError:
                    errors.append(f"Α/Α {aa}: μη έγκυρες τιμές βάρους!")

        if errors:
            messagebox.showerror("Σφάλματα",
                "\n".join(errors), parent=self.root)
            return

        self.result = self.data
        self.root.destroy()

    def _cancel(self):
        if messagebox.askyesno("Ακύρωση", "Σίγουρα θέλεις να ακυρώσεις;", parent=self.root):
            self.root.destroy()

    def run(self) -> dict | None:
        self.root.mainloop(); return self.result

# ==============================================================================
# HELPER — Expand 0832 rows με κράματα
# ==============================================================================

KRAMMA_DASMOS_KL = {"76012080", "76012030", "76011010"}

def is_0832_kramma(row) -> bool:
    """Ελέγχει αν η γραμμή χρειάζεται χειροκίνητο κράμα (0832 + συγκεκριμένη ΔΑΣΜ_ΚΛ)."""
    return (str(row.get("ΤΕΛΩΝ", "")) == "0832" and
            str(row.get("ΔΑΣΜ_ΚΛ", "")) in KRAMMA_DASMOS_KL)


def expand_krammata(df_result: pd.DataFrame, krammata: dict) -> pd.DataFrame:
    """
    Σπάει κάθε είδος σε 1-3 γραμμές ανάλογα με τα κράματα.

    krammata: dict με keys = str(Α/Α), values = dict με ΚΡΑΜΑ_1/ΒΑΡΟΣ_1 κλπ
              ή απευθείας από το 1ο popup (ένα dict για όλα τα είδη)

    Κανόνες:
    - Αν ΚΡΑΜΑ_1 μόνο (ΒΑΡΟΣ_1 κενό) → 1 γραμμή με το συνολικό βάρος
    - Αν ΚΡΑΜΑ_1 + ΒΑΡΟΣ_1 + ΚΡΑΜΑ_2 + ΒΑΡΟΣ_2 → 2 γραμμές με τα αντίστοιχα βάρη
    - Αν και ΚΡΑΜΑ_3 → 3 γραμμές
    """
    new_rows = []
    for _, row in df_result.iterrows():
        aa  = str(row["Α/Α"])
        # Αν δεν είναι 0832 kramma → αφήνουμε ως έχει
        if not is_0832_kramma(row):
            new_rows.append(row.to_dict())
            continue

        d = krammata.get(aa, {})
        kr1 = d.get("ΚΡΑΜΑ_1", "").strip()
        vr1 = d.get("ΒΑΡΟΣ_1", "").strip()
        kr2 = d.get("ΚΡΑΜΑ_2", "").strip()
        vr2 = d.get("ΒΑΡΟΣ_2", "").strip()
        kr3 = d.get("ΚΡΑΜΑ_3", "").strip()
        vr3 = d.get("ΒΑΡΟΣ_3", "").strip()

        if not kr1:
            # Δεν έχει κράμα → αφήνουμε ως έχει (θα ρωτήσει popup_input μέσα στο sap_entry)
            new_rows.append(row.to_dict())
            continue

        # Γραμμή 1
        r1 = row.to_dict()
        r1["ΚΡΑΜΑ"] = kr1
        # Αν ΒΑΡΟΣ_1 κενό → παίρνει το συνολικό βάρος (μόνο 1 κράμα)
        r1["ΒΑΡΟΣ"] = vr1 if vr1 else row["ΒΑΡΟΣ"]
        new_rows.append(r1)

        # Γραμμή 2 (αν υπάρχει ΚΡΑΜΑ_2)
        if kr2:
            r2 = row.to_dict()
            r2["ΚΡΑΜΑ"] = kr2
            r2["ΒΑΡΟΣ"] = vr2
            new_rows.append(r2)

        # Γραμμή 3 (αν υπάρχει ΚΡΑΜΑ_3)
        if kr3:
            r3 = row.to_dict()
            r3["ΚΡΑΜΑ"] = kr3
            r3["ΒΑΡΟΣ"] = vr3
            new_rows.append(r3)

    return pd.DataFrame(new_rows).reset_index(drop=True)

# ==============================================================================
# ΦΑΣΗ Β.1 — Download XML
# ==============================================================================

def phase_b_download_xml(mrn: str, driver):
    SAVE_FOLDER.mkdir(parents=True, exist_ok=True)
    for f in SAVE_FOLDER.glob("*.xml"):
        try: f.unlink()
        except: pass

    # Έλεγξε αν ο browser είναι ζωντανός
    try:
        _ = driver.current_url
    except:
        print("    Browser έκλεισε — επανασύνδεση ICISnet...")
        driver = make_chrome(download_folder=SAVE_FOLDER)
        wait   = WebDriverWait(driver, 30)
        driver.get("https://www1.gsis.gr/icisnet/itrader/common/home.jsf")
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(ICISNET_USER)
        driver.find_element(By.NAME, "password").send_keys(ICISNET_PASS)
        driver.find_element(By.NAME, "btn_login").click()
        time.sleep(3)

    # Επιστροφή στο menu — ← ΕΔΩ
    driver.get("https://www1.gsis.gr/icisnet/itrader/common/home.jsf")
    driver.maximize_window()
    time.sleep(3)

    wait = WebDriverWait(driver, 30)
    print(f"    ICISnet download XML για MRN={mrn}...")

    clicked = False
    for _ in range(5):
        try:
            driver.execute_script("arguments[0].click();",
                driver.find_element(By.ID, "tablehidecontentForm:actions_message_search"))
            clicked = True; break
        except: time.sleep(1)
    if not clicked:
        for _ in range(5):
            try:
                driver.execute_script("arguments[0].click();",
                    driver.find_element(By.ID, "contentForm:actions_message_search"))
                time.sleep(2)
                driver.execute_script("arguments[0].click();",
                    driver.find_element(By.ID, "tablehidecontentForm:actions_message_search"))
                clicked = True; break
            except: time.sleep(1)
    if not clicked: raise Exception("Menu navigation failed")

    wait.until(EC.presence_of_element_located((By.ID, "contentForm:mrn-arc-nid")))
    inp = wait.until(EC.element_to_be_clickable((By.ID, "contentForm:mrn-arc-nid")))
    inp.clear(); inp.send_keys(mrn)
    safe_select(wait, "contentForm:domain",       "Εισαγωγές"); time.sleep(2)
    safe_select(wait, "contentForm:message_type", "ID29")
    safe_select(wait, "contentForm:search_scope", "Συναλλασσόμενος")

    driver.execute_script("arguments[0].click();",
        wait.until(EC.element_to_be_clickable((By.ID, "contentForm:submitFormButton"))))
    wait.until(EC.presence_of_element_located((By.ID, "contentForm:search_results")))

    driver.execute_script("arguments[0].click();",
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "ID29"))))

    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//*[contains(text(),'Ενέργειες')]"))).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//*[contains(text(),'Αποθήκευση ως XML')]"))).click()

    for _ in range(20):
        time.sleep(1)
        xmls = list(SAVE_FOLDER.glob("*.xml"))
        if xmls:
            latest = max(xmls, key=lambda f: f.stat().st_ctime)
            XML_DIR.mkdir(parents=True, exist_ok=True)
            dest = XML_DIR / f"{mrn}.xml"
            dest.unlink(missing_ok=True)
            shutil.move(str(latest), str(dest))
            print(f"    XML saved: {dest.name}")
            driver.minimize_window()
            return driver

    raise Exception("XML download timed out")

# ==============================================================================
# ΦΑΣΗ Β.2 — Parse XML
# ==============================================================================

def phase_b_parse_xml(xml_path: Path) -> pd.DataFrame:
    import xml.etree.ElementTree as ET
    tree = ET.parse(xml_path); root = tree.getroot()

    def fmt(d): return d.strftime("%d.%m.%Y")
    def to_comma(val):
        if val is None or val == "": return ""
        try:
            f = float(val)
            s = str(int(f)) if f == int(f) else f"{f:.2f}"
            return s.replace(".", ",")
        except: return str(val).replace(".", ",")

    def to_comma_rate(val):
        """Για ισοτιμία — max 5 δεκαδικά."""
        if val is None or val == "": return ""
        try:
            f = float(val)
            s = str(int(f)) if f == int(f) else f"{f:.5f}".rstrip("0")
            return s.replace(".", ",")
        except:
            return str(val).replace(".", ",")

    hea = root.find("HEAHEA")
    if hea is None: raise ValueError("HEAHEA tag not found in XML")

    mrn         = hea.findtext("DocNumHEA5", "")
    typ_dec     = hea.findtext("TypOfDecHEA24", "")
    typ_dec_b12 = hea.findtext("TypOfDecBx12HEA651", "")
    cou_dis     = hea.findtext("CouOfDisCodHEA55", "")
    dec_date    = datetime.strptime(hea.findtext("DecDatHEA383", ""), "%Y%m%d").date()
    typos       = typ_dec + typ_dec_b12

    hl  = dec_date + relativedelta(months=32) - relativedelta(days=1)
    hl2 = hl + relativedelta(months=-26)
    hp  = hl2 + relativedelta(months=1)

    telwneio = root.find("IMPCUSOFF").findtext("RefNumIMPCUSOFF", "")[-4:]
    prom     = root.find("TRACONCO1").findtext("NamCO17", "").strip()
    oroi     = root.find("DELTER").findtext("IncCodTDL1", "")
    tradat   = root.find("TRADAT")
    nom_isot = tradat.findtext("CurTRD1", "")
    isot     = tradat.findtext("ExcRatTRD1", "1")

    rows = []
    for item in root.findall("GOOITEGDS"):
        aa       = item.findtext("IteNumGDS7", "")
        desc     = item.findtext("GooDesGDS23", "")
        net_mas  = item.findtext("NetMasGDS48", "")
        pro_req  = item.findtext("ProReqGDI1", "")
        pre_pro  = item.findtext("PreProGDI1", "")
        nom_stat = item.findtext("StaValCurGDI1", "")
        cou_ori  = item.findtext("CouOfOriGDI1", "")
        pro_pri  = item.findtext("ProPri4002", "")
        sta_val  = item.findtext("StaValAmoGDI1", "")

        kramma = ""
        if "SLABS " in desc:
            kramma = desc.split("SLABS ", 1)[1].split(" ", 1)[0][:4]

        tel_kath = pro_req + pre_pro
        rel      = item.find("REL800")
        rel_cod  = rel.findtext("RelRelCod02", "") if rel is not None else ""

        if rel_cod == "X16":              kath = 3
        elif tel_kath.startswith("51"):   kath = 2
        else:                             kath = 5

        xwra   = cou_dis.replace("XS", "RS") if cou_ori == "EU" else cou_ori
        ent_ag = ""
        for doc in item.findall("PRODOCDC2"):
            di = doc.findtext("DocInfDC1008", "")
            if di: ent_ag = di.replace("/", "000").replace("-", "000"); break

        comcod    = item.find("COMCODGODITM")
        dasmos_kl = comcod.findtext("ComNomCMD1", "") if comcod is not None else ""

        taxes = [(float(t.findtext("RatOfTaxCTX1", "0")),
                  float(t.findtext("AmoOfTaxTCL1", "0")))
                 for t in item.findall("CALTAXGOD")]
        if taxes:
            rates = [t[0] for t in taxes]; amos = [t[1] for t in taxes]
            sd = min(rates); da = min(amos); sf = max(rates); fa = max(amos)
            df_val = None if da == fa else da; sf_fin = sf if sf else 24.0
        else:
            sd = da = fa = None; df_val = None; sf_fin = 24.0

        rows.append({
            "ΤΕΛΩΝ": telwneio, "ΤΥΠΟΣ": typos, "MRN": mrn, "ΗΜΕΡ": fmt(dec_date),
            "ΔΑΣΜ_ΚΛ": dasmos_kl, "Α/Α": int(aa) if aa.isdigit() else aa,
            "ΧΩΡΑ": xwra, "ΚΡΑΜΑ": kramma, "ΚΑΘ": kath, "ΤΕΛ_ΚΑΘ": tel_kath,
            "ΗΜΕΡ_ΛΗΞΗΣ": fmt(hl), "ΗΜΕΡ_ΛΗΞΗΣ2": fmt(hl2),
            "ΒΑΡΟΣ": to_comma(net_mas), "X16": rel_cod, "ΝΟΜ_ΣΤΑΤ": nom_stat,
            "ΣΤΑΤ_ΑΞΙΑ": f"{float(sta_val):.2f}".replace(".", ",") if sta_val else "",
            "ΠΡΟΜ": prom, "ΟΡΟΙ": oroi,
            "ΤΙΜΗ": f"{float(pro_pri):.2f}".replace(".", ",") if pro_pri else "",
            "ΝΟΜ_ΙΣΟΤ": nom_isot, "ΙΣΟΤ": to_comma_rate(isot),
            "ΣΥΝΤ_ΔΑΣΜ": to_comma(sd), "ΔΑΣΜ": to_comma(df_val),
            "ΣΥΝΤ_ΦΠΑ": to_comma(sf_fin), "ΦΠΑ": to_comma(fa),
            "ΕΝΤ_ΑΓ": ent_ag, "ΗΜΕΡ_ΠΡΟΘ": fmt(hp),
        })

    cols = ["ΤΕΛΩΝ","ΤΥΠΟΣ","MRN","ΗΜΕΡ","ΔΑΣΜ_ΚΛ","Α/Α","ΧΩΡΑ","ΚΡΑΜΑ","ΚΑΘ","ΤΕΛ_ΚΑΘ",
            "ΗΜΕΡ_ΛΗΞΗΣ","ΗΜΕΡ_ΛΗΞΗΣ2","ΒΑΡΟΣ","X16","ΝΟΜ_ΣΤΑΤ","ΣΤΑΤ_ΑΞΙΑ","ΠΡΟΜ","ΟΡΟΙ",
            "ΤΙΜΗ","ΝΟΜ_ΙΣΟΤ","ΙΣΟΤ","ΣΥΝΤ_ΔΑΣΜ","ΔΑΣΜ","ΣΥΝΤ_ΦΠΑ","ΦΠΑ","ΕΝΤ_ΑΓ","ΗΜΕΡ_ΠΡΟΘ"]
    return pd.DataFrame(rows)[cols]

# ==============================================================================
# ΦΑΣΗ Β.3 — SAP entry
# ==============================================================================

def get_kramma(dasmos_kl: str, synt_dasm: str):
    if dasmos_kl.startswith("760200"):
        return "SCRAP-ΦΥΡΑ" if synt_dasm == "24" else "SCRAP"
    m = {
        "76011090": "1XXX", "76012080": "AL", "76012030": "AL", "76011010": "AL",
        "81041100": "MG",   "81110011": "MN",
        "72052900": "FE",   "72069000": "FE", "72029980": "FE",
    }
    return m.get(dasmos_kl)


def focus_sap():
    try:
        Application(backend="uia")\
            .connect(title_re=".*Νέες.*Καταχωρίσεων.*", timeout=5)\
            .window(title_re=".*Νέες.*Καταχωρίσεων.*").set_focus()
        time.sleep(0.3)
    except:
        try:
            Application(backend="uia")\
                .connect(class_name="SAP_FRONTEND_SESSION", timeout=5)\
                .window(class_name="SAP_FRONTEND_SESSION").set_focus()
            time.sleep(0.3)
        except: pass

SAP_WIN_TITLE = "Νέες Καταχωρίσεις: Λεπτομέρειες Προστιθέμενων Καταχωρίσεων"


def get_sap_win(timeout=10):
    app = Application(backend="uia").connect(
        class_name="SAP_FRONTEND_SESSION", title=SAP_WIN_TITLE, timeout=timeout)
    return app.window(class_name="SAP_FRONTEND_SESSION", title=SAP_WIN_TITLE)


ZIMP_BASE = "/app/con[0]/ses[0]/wnd[0]/usr/"


def sap_entry(session, row, prot: str):
    """
    SAP GUI Scripting έκδοση — γεμίζει όλα τα πεδία της οθόνης "Νέες
    Καταχωρίσεις" μέσω findById(...).text (όχι send_keys/TAB, άρα η σειρά
    δεν έχει σημασία — κάθε πεδίο γράφεται απευθείας με το ID του).

    ΣΤΑΜΑΤΑΕΙ ΑΚΡΙΒΩΣ ΠΡΙΝ ΤΟ SAVE: δεν πατάει ποτέ btn[11] (Save) ούτε F11.
    Το session πρέπει να είναι ήδη στην οθόνη καταχώρησης (μετά το "Νέες
    Καταχωρίσεις").

    Ισοδυναμεί με το παλιό sap_entry() του IMA-IMC.py, field-by-field
    (βλ. ima_mapping.txt) — καμία τιμή/κλάδος δεν έχει αλλάξει λογική.
    """
    B = ZIMP_BASE

    tl  = str(row["ΤΕΛΩΝ"]);  ty  = str(row["ΤΥΠΟΣ"]); mn  = str(row["MRN"])
    im  = str(row["ΗΜΕΡ"]);   dk  = str(row["ΔΑΣΜ_ΚΛ"]); aa = str(row["Α/Α"])
    xw  = str(row["ΧΩΡΑ"]);   x16 = str(row["X16"]); tk_ = str(row["ΤΕΛ_ΚΑΘ"])
    il  = str(row["ΗΜΕΡ_ΛΗΞΗΣ"]); il2 = str(row["ΗΜΕΡ_ΛΗΞΗΣ2"])
    vr  = str(row["ΒΑΡΟΣ"]);  sa  = str(row["ΣΤΑΤ_ΑΞΙΑ"]); ns = str(row["ΝΟΜ_ΣΤΑΤ"])
    pr  = str(row["ΠΡΟΜ"]);   or_ = str(row["ΟΡΟΙ"]); ti = str(row["ΤΙΜΗ"])
    ni  = str(row["ΝΟΜ_ΙΣΟΤ"]); is_ = str(row["ΙΣΟΤ"]); sd = str(row["ΣΥΝΤ_ΔΑΣΜ"])
    da  = str(row["ΔΑΣΜ"]);   sf  = str(row["ΣΥΝΤ_ΦΠΑ"]); fp = str(row["ΦΠΑ"])
    ea  = str(row["ΕΝΤ_ΑΓ"]).strip(); ip = str(row["ΗΜΕΡ_ΠΡΟΘ"])

    kath = compute_kath(x16, tk_)

    res = resolve_inputs(row, prot)
    kr1, vr1 = res["ΚΡΑΜΑ_1"], res["ΒΑΡΟΣ_1"]
    kr2, vr2 = res["ΚΡΑΜΑ_2"], res["ΒΑΡΟΣ_2"]
    kr3, vr3 = res["ΚΡΑΜΑ_3"], res["ΒΑΡΟΣ_3"]
    prot, ea = res["PROT"], res["ΕΝΤ_ΑΓ"]

    log.info(
        f"    Α/Α={aa} | ΚΑΘ={kath} | ΚΡΑΜΑ_1={kr1} | ΒΑΡΟΣ_1={vr1} | ΤΕΛΩΝ={tl} | "
        f"ΣΥΝΤ_ΔΑΣΜ={sd} | ΔΑΣΜ={da} | ΣΥΝΤ_ΦΠΑ={sf} | ΦΠΑ={fp} | ΕΝΤ_ΑΓ(ea)={ea!r} | "
        f"ΠΡΟΜ={pr} | ΟΡΟΙ={or_} | ΤΙΜΗ={ti} | ΒΑΡΟΣ={vr} | ΣΤΑΤ_ΑΞΙΑ={sa}"
    )
    _sap_entry_fields(session, B, row, kath, kr1, vr1, kr2, vr2, kr3, vr3, prot, ea)


def resolve_inputs(row, prot: str) -> dict:
    """Όλα τα popups μιας καταχώρησης (ΚΡΑΜΑ / ΠΡΟΤΙΜΗΣΗ / ΕΝΤΟΛΗ ΑΓΟΡΑΣ).
    Καλείται στο Μέρος 1 (πριν ανοίξει το SAP) και οι απαντήσεις γράφονται
    στη γραμμή — στο Μέρος 2 βρίσκει τις τιμές συμπληρωμένες και δεν ρωτάει ξανά."""
    mn = str(row["MRN"]); dk = str(row["ΔΑΣΜ_ΚΛ"]); aa = str(row["Α/Α"])
    sd = str(row["ΣΥΝΤ_ΔΑΣΜ"]); ea = str(row["ΕΝΤ_ΑΓ"]).strip()

    # ── ΚΡΑΜΑΤΑ: έως 3 ζεύγη ΚΡΑΜΑ/ΒΑΡΟΣ σε ΜΙΑ καταχώρηση ──────────────
    # Αν το row έχει ήδη ΚΡΑΜΑ_1/ΒΑΡΟΣ_1 (από το popup για 0832, βλ.
    # prepare_mrn) τα χρησιμοποιούμε ως έχουν. Αλλιώς (γνωστό ΔΑΣΜ_ΚΛ) ένα
    # μόνο κράμα από get_kramma() στη θέση 1, με όλο το βάρος.
    kr1 = str(row.get("ΚΡΑΜΑ_1", "")).strip()
    vr1 = str(row.get("ΒΑΡΟΣ_1", "")).strip()
    kr2 = str(row.get("ΚΡΑΜΑ_2", "")).strip()
    vr2 = str(row.get("ΒΑΡΟΣ_2", "")).strip()
    kr3 = str(row.get("ΚΡΑΜΑ_3", "")).strip()
    vr3 = str(row.get("ΒΑΡΟΣ_3", "")).strip()

    if not kr1:
        # Προτεραιότητα στο ΚΡΑΜΑ που εξάγεται αυτόματα από το XML (περιγραφή
        # "SLABS <κωδικός>..." — βλ. parse_xml) — ίδια σειρά προτεραιότητας με
        # τον παλιό κώδικα ("popup έχει προτεραιότητα πάντα"), εδώ για τη θέση 1.
        xml_kr = str(row.get("ΚΡΑΜΑ", "")).strip()
        if xml_kr and xml_kr not in ("nan", ""):
            kr1 = xml_kr
        elif is_0832_kramma(row):
            kr1 = popup_input("ΚΡΑΜΑ", f"ΤΕΛΩΝ=0832 | {dk}\nΕισάγετε ΚΡΑΜΑ (1ο):")
        else:
            kr1 = get_kramma(dk, sd)
            if kr1 is None:
                kr1 = popup_input("ΚΡΑΜΑ", f"Εισάγετε ΚΡΑΜΑ για {dk}:")
        # ΒΑΡΟΣ_1 μένει κενό (όπως ο παλιός κώδικας) όταν υπάρχει μόνο ένα
        # κράμα — το συνολικό βάρος καλύπτεται ήδη από το ΠΟΣΟΤΗΤΑ ΕΙΣΑΓΩΓΗΣ
        # (vr, πεδίο MENGE). Το ΒΑΡΟΣ_1 γεμίζει μόνο όταν υπάρχουν 2+ κράματα
        # (βλ. expand_krammata: "ΚΡΑΜΑ_1 μόνο → ΒΑΡΟΣ_1 κενό, συνολικό βάρος").

    if not prot or prot in ("", "nan"):
        prot = popup_input("ΠΡΟΤΙΜΗΣΗ", f"MRN:{mn}|Α/Α:{aa}\nΕισάγετε ΠΡΟΤΙΜΗΣΗ:")
        if not prot: raise ValueError("ΠΡΟΤΙΜΗΣΗ κενή")

    if sd != "24" and ea in ("", "nan", "None"):
        raw = popup_input("ΕΝΤΟΛΗ ΑΓΟΡΑΣ", f"MRN:{mn}\n5 ψηφία ΕΝΤ_ΑΓ (ή NO):")
        if raw:
            ea = "NO" if raw.upper() == "NO" else f"41000{raw.strip()}"

    # Το πεδίο EBELN στο SAP δέχεται ΑΚΡΙΒΩΣ 10 χαρακτήρες — οτιδήποτε άλλο
    # σκάει στο session.findById(...).text = ea με "Property can not be set"
    # (confirmed live 2026-10-01, MRN YOUR_MRN_6: ea είχε 13 χαρακτήρες
    # αντί για 10). Αντί να σκάσει, ζητάμε διόρθωση από τον χρήστη.
    while sd != "24" and ea.upper() != "NO" and len(ea) != 10:
        raw = popup_input(
            "ΕΝΤΟΛΗ ΑΓΟΡΑΣ — ΛΑΘΟΣ ΜΗΚΟΣ",
            f"MRN:{mn}\nΗ τιμή '{ea}' έχει {len(ea)} χαρακτήρες (πρέπει ΑΚΡΙΒΩΣ 10).\n"
            f"Διόρθωσε (ή γράψε NO):",
            default=ea,
        )
        if not raw:
            raise ValueError(f"ΕΝΤΟΛΗ ΑΓΟΡΑΣ μη έγκυρη (ακυρώθηκε): {ea!r}")
        ea = "NO" if raw.upper() == "NO" else raw.strip()

    return {"ΚΡΑΜΑ_1": kr1, "ΒΑΡΟΣ_1": vr1, "ΚΡΑΜΑ_2": kr2, "ΒΑΡΟΣ_2": vr2,
            "ΚΡΑΜΑ_3": kr3, "ΒΑΡΟΣ_3": vr3, "PROT": prot, "ΕΝΤ_ΑΓ": ea}


def _sap_entry_fields(session, B, row, kath, kr1, vr1, kr2, vr2, kr3, vr3, prot, ea):
    """Γέμισμα πεδίων (αμετάβλητη λογική από την προηγούμενη έκδοση του sap_entry)."""
    tl  = str(row["ΤΕΛΩΝ"]);  ty  = str(row["ΤΥΠΟΣ"]); mn  = str(row["MRN"])
    im  = str(row["ΗΜΕΡ"]);   dk  = str(row["ΔΑΣΜ_ΚΛ"]); aa = str(row["Α/Α"])
    xw  = str(row["ΧΩΡΑ"]);   x16 = str(row["X16"]); tk_ = str(row["ΤΕΛ_ΚΑΘ"])
    il  = str(row["ΗΜΕΡ_ΛΗΞΗΣ"]); il2 = str(row["ΗΜΕΡ_ΛΗΞΗΣ2"])
    vr  = str(row["ΒΑΡΟΣ"]);  sa  = str(row["ΣΤΑΤ_ΑΞΙΑ"]); ns = str(row["ΝΟΜ_ΣΤΑΤ"])
    pr  = str(row["ΠΡΟΜ"]);   or_ = str(row["ΟΡΟΙ"]); ti = str(row["ΤΙΜΗ"])
    ni  = str(row["ΝΟΜ_ΙΣΟΤ"]); is_ = str(row["ΙΣΟΤ"]); sd = str(row["ΣΥΝΤ_ΔΑΣΜ"])
    da  = str(row["ΔΑΣΜ"]);   sf  = str(row["ΣΥΝΤ_ΦΠΑ"]); fp = str(row["ΦΠΑ"])
    ip  = str(row["ΗΜΕΡ_ΠΡΟΘ"])

    # ── Βασικά πεδία κεφαλίδας ────────────────────────────────────────────
    session.findById(B + "ctxtZIMP_1-ZOLLA").text = tl
    session.findById(B + "txtZIMP_1-ZOLLA_TYPE_B").text = ty
    session.findById(B + "ctxtZIMP_1-ZIMPT").text = mn
    session.findById(B + "ctxtZIMP_1-ZDATE").text = im
    if tl != "0832":
        session.findById(B + "txtZIMP_1-ZOLLA_B").text = tl
        session.findById(B + "txtZIMP_1-ZIMPT_B").text = mn
        session.findById(B + "ctxtZIMP_1-ZDATE_B").text = im
    session.findById(B + "ctxtZIMP_1-STAWN").text = dk
    session.findById(B + "ctxtZIMP_1-BUKRS").text = COMPANY_CODE
    session.findById(B + "txtZIMP_1-ZAA").text = aa
    session.findById(B + "ctxtZIMP_1-ZLAND").text = xw

    # ── Καθεστώς ──────────────────────────────────────────────────────────
    session.findById(B + "ctxtZIMP_1-Z_IMP").text = str(kath)
    session.findById(B + "txtZIMP_1-Z_IMP_TEL").text = tk_

    is_special = (kath == 3 and tk_ == "6121" and pr == "SUPPLIER_X LTD" and dk == "76012080")

    if is_special:
        session.findById(B + "ctxtZIMP_1-ZDUED").text = il
        session.findById(B + "txtZIMP_1-ZDOCDYN").text = "YOUR_DOC_REF_1"
        session.findById(B + "txtZIMP_1-MENGE").text = vr
        session.findById(B + "ctxtZIMP_1-MEINS").text = "KG"
        session.findById(B + "txtZIMP_1-VALUE").text = sa
        session.findById(B + "txtZIMP_1-ZWRBTR_EUR_42").text = sa
        session.findById(B + "ctxtZIMP_1-WAERS").text = ns

        pr = "SUPPLIER_X - ΠΑΘΗΤΙΚΗ ΤΕΛ."
        session.findById(B + "txtZIMP_1-ZLFINN").text = pr
        session.findById(B + "txtZIMP_1-ZPSAPOF").text = "YOUR_PERMIT_REF"
        session.findById(B + "ctxtZIMP_1-ZQM_ALLOY").text = "AL"

        session.findById(B + "txtZIMP_1-ZPREFERENCE").text = prot
        session.findById(B + "txtZIMP_1-ZDELIV_TERMS").text = or_
        session.findById(B + "txtZIMP_1-ZWRBTR_42").text = ti
        session.findById(B + "txtZIMP_1-ZWAERS_42").text = ni
        session.findById(B + "txtZIMP_1-ZKURSF_42").text = is_
        session.findById(B + "txtZIMP_1-VALUE").text = sa
        session.findById(B + "txtZIMP_1-ZWRBTR_EUR_42").text = sa
        if sd != "24":
            session.findById(B + "txtZIMP_1-ZKBERT_DUTIES_42").text = sd
            session.findById(B + "txtZIMP_1-ZWRBTR_DUTIES_42").text = da
        session.findById(B + "txtZIMP_1-ZKBERT_TAXES_42").text = sf
        session.findById(B + "txtZIMP_1-ZWRBTR_TAXES_42").text = fp
        if sd != "24" and ea.upper() != "NO":
            session.findById(B + "ctxtZIMP_1-EBELN").text = ea

        session.findById(B + "chkZIMP_1-ZCHECK").Selected = True
        print("    [SUPPLIER_X] Στοιχεία γεμίστηκαν — ΣΤΑΜΑΤΗΣΕ πριν το Save.")
        return

    # ── Μη-SUPPLIER_X μονοπάτι ─────────────────────────────────────────────────
    if kath == 3:
        session.findById(B + "ctxtZIMP_1-ZDUED").text = il
    elif kath in (5, 12):
        pass
    else:  # kath == 2
        session.findById(B + "ctxtZIMP_1-ZDUED").text = il2
        session.findById(B + "txtZIMP_1-ZDOCDYN").text = "23GR000001IP00242"

    session.findById(B + "txtZIMP_1-MENGE").text = vr
    session.findById(B + "ctxtZIMP_1-MEINS").text = "KG"
    session.findById(B + "txtZIMP_1-VALUE").text = sa
    session.findById(B + "txtZIMP_1-ZWRBTR_EUR_42").text = sa
    session.findById(B + "ctxtZIMP_1-WAERS").text = ns

    if sd == "24":
        pr = f"ΦΥΡΑ - ΕΚΚΑΘΑΡΙΣΗ {im}"

    is_c = ty.upper().endswith("C")

    # Hardcoded reference documents (βλ. ima_mapping.txt) — ταυτόσημο με τον
    # παλιό κώδικα, ΣΥΜΠΕΡΙΛΑΜΒΑΝΟΜΕΝΟΥ του edge case: αν is_c ΚΑΙ ΚΑΘ==12,
    # ο παλιός κώδικας δεν έγραφε ΤΙΠΟΤΑ εδώ (ούτε καν ΠΡΟΜ/ZLFINN) — το ίδιο κρατάμε.
    if kath == 2:
        session.findById(B + "txtZIMP_1-ZLICENSE_C").text = "YOUR_DOC_REF_2"
        session.findById(B + "txtZIMP_1-ZGRN_WARRANTY").text = "YOUR_DOC_REF_3"
        session.findById(B + "txtZIMP_1-ZLICENSE_S").text = "18GR000001SASP00058"
        session.findById(B + "txtZIMP_1-ZLFINN").text = pr
        session.findById(B + "ctxtZIMP_1-ZDATANT").text = ip
        session.findById(B + "txtZIMP_1-ZPSAPOF").text = "10790/13-06-2003"
    elif kath == 3:
        if is_c:
            session.findById(B + "txtZIMP_1-ZLICENSE_C").text = "18GR000001CGU1CT000004"
            session.findById(B + "txtZIMP_1-ZGRN_WARRANTY").text = "YOUR_DOC_REF_4"
            session.findById(B + "txtZIMP_1-ZLICENSE_S").text = "18GR000001SASP00057"
        session.findById(B + "txtZIMP_1-ZLFINN").text = pr
        session.findById(B + "txtZIMP_1-ZPSAPOF").text = "YOUR_PERMIT_REF"
    elif kath == 5:
        if is_c:
            session.findById(B + "txtZIMP_1-ZLICENSE_C").text = "18GR000001CGU1CT000004"
            session.findById(B + "txtZIMP_1-ZGRN_WARRANTY").text = "YOUR_DOC_REF_4"
            session.findById(B + "txtZIMP_1-ZLICENSE_S").text = "18GR000001SASP00057"
        session.findById(B + "txtZIMP_1-ZLFINN").text = pr
        session.findById(B + "txtZIMP_1-ZPSAPOF").text = "."
    elif kath == 12:
        if not is_c:
            session.findById(B + "txtZIMP_1-ZLFINN").text = pr
            session.findById(B + "txtZIMP_1-ZPSAPOF").text = "."
        # is_c και ΚΑΘ==12: όπως ο παλιός κώδικας — δεν γράφεται τίποτα εδώ.

    # ── Κράματα (έως 3 ζεύγη) ────────────────────────────────────────────
    session.findById(B + "ctxtZIMP_1-ZQM_ALLOY").text = kr1
    session.findById(B + "txtZIMP_1-ZQM_ALLOY_MENGE").text = vr1
    if kr2:
        session.findById(B + "ctxtZIMP_1-ZQM_ALLOY_2").text = kr2
        session.findById(B + "txtZIMP_1-ZQM_ALLOY_2_MENGE").text = vr2
    if kr3:
        session.findById(B + "ctxtZIMP_1-ZQM_ALLOY_3").text = kr3
        session.findById(B + "txtZIMP_1-ZQM_ALLOY_3_MENGE").text = vr3

    # ── Προμηθευτής / Όροι / Τιμή / Ισοτιμία ────────────────────────────
    session.findById(B + "txtZIMP_1-ZPREFERENCE").text = prot
    session.findById(B + "txtZIMP_1-ZDELIV_TERMS").text = or_
    session.findById(B + "txtZIMP_1-ZWRBTR_42").text = ti
    session.findById(B + "txtZIMP_1-ZWAERS_42").text = ni
    session.findById(B + "txtZIMP_1-ZKURSF_42").text = is_
    session.findById(B + "txtZIMP_1-VALUE").text = sa
    session.findById(B + "txtZIMP_1-ZWRBTR_EUR_42").text = sa

    # ── Δασμός / ΦΠΑ ─────────────────────────────────────────────────────
    if sd == "24":
        session.findById(B + "txtZIMP_1-ZKBERT_TAXES_42").text = sf
        session.findById(B + "txtZIMP_1-ZWRBTR_TAXES_42").text = fp
    else:
        session.findById(B + "txtZIMP_1-ZKBERT_DUTIES_42").text = sd
        session.findById(B + "txtZIMP_1-ZWRBTR_DUTIES_42").text = da
        session.findById(B + "txtZIMP_1-ZKBERT_TAXES_42").text = sf
        session.findById(B + "txtZIMP_1-ZWRBTR_TAXES_42").text = fp
        if ea.upper() != "NO":
            session.findById(B + "ctxtZIMP_1-EBELN").text = ea

    # ── Checkbox επιλογής γραμμής (πριν το Save) ────────────────────────
    session.findById(B + "chkZIMP_1-ZCHECK").Selected = True

    print("    Στοιχεία γεμίστηκαν — ΣΤΑΜΑΤΗΣΕ πριν το Save (δεν πατήθηκε τίποτα).")

# ==============================================================================
# ΦΑΣΗ Β.4 — SAP attach PDF
# ==============================================================================

def sap_attach(session, pdf_path: Path):
    """
    GOS attach μέσω SAP GUI Scripting (Generic Object Services) — ίδιο
    pattern με το ήδη validated attach_pdf_gos() του ΕΙΣΑΓΩΓΕΣ ΠΕΙΡΑΙΑ
    script, adapted από recording του χρήστη στο IMA/C. Επιβεβαιώνει μέσω
    status bar πριν συνεχίσει.
    """
    session.findById("wnd[0]/titl/shellcont/shell").pressContextButton("%GOS_TOOLBOX")
    session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem("%GOS_PCATTA_CREA")
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(pdf_path.parent)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = pdf_path.name
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    status_text = ""
    for _ in range(20):
        status_text = session.findById("wnd[0]/sbar/pane[0]").text
        if "Προσάρτηση δημιουργήθηκε με επιτυχία" in status_text:
            break
        time.sleep(0.5)
    else:
        raise RuntimeError(f"Δεν επιβεβαιώθηκε το attach. Status bar: '{status_text}'")

    print(f"    Attached OK: {pdf_path.name}")

# ==============================================================================
# PROCESS ONE MRN
# ==============================================================================

def prepare_mrn(row: dict, driver) -> tuple:
    """
    Μέρος 1 (μόνο ICISNet, το SAP είναι κλειστό): XML (από τον φάκελο xml
    αν υπάρχει, αλλιώς λήψη) -> έλεγχος MRN -> parse -> φίλτρα -> κράματα ->
    ΟΛΑ τα popups (ΚΡΑΜΑ / ΠΡΟΤΙΜΗΣΗ / ΕΝΤΟΛΗ ΑΓΟΡΑΣ) για κάθε Α/Α.
    Επιστρέφει (job | None, status, driver) — status: 'ok' / 'skip' / 'fail'.
    Το driver μπορεί να είναι None (ανοίγει μόνο αν χρειαστεί λήψη).
    """
    mrn  = str(row.get("MRN",  "")).strip()
    prot = str(row.get("PROT", "")).strip()
    pdf  = str(row.get("PDF",  "")).strip()
    pdf1 = str(row.get("PDF.1","")).strip()

    prot = "" if prot in ("nan","None") else prot
    pdf  = "" if pdf  in ("nan","None") else pdf
    pdf1 = "" if pdf1 in ("nan","None") else pdf1
    if prot and prot.replace(".", "").isdigit():
        prot = str(int(float(prot)))

    # Κράματα από το 1ο popup (για 0832)
    krammata_from_popup = {
        "ΚΡΑΜΑ_1": str(row.get("ΚΡΑΜΑ_1", "")).strip(),
        "ΒΑΡΟΣ_1":  str(row.get("ΒΑΡΟΣ_1",  "")).strip(),
        "ΚΡΑΜΑ_2": str(row.get("ΚΡΑΜΑ_2", "")).strip(),
        "ΒΑΡΟΣ_2":  str(row.get("ΒΑΡΟΣ_2",  "")).strip(),
        "ΚΡΑΜΑ_3": str(row.get("ΚΡΑΜΑ_3", "")).strip(),
        "ΒΑΡΟΣ_3":  str(row.get("ΒΑΡΟΣ_3",  "")).strip(),
    }

    log.info(f"\n{'='*60}\n  MRN : {mrn}  |  PROT: {prot}\n  PDF : {pdf1} -> {pdf}\n{'='*60}")

    SAVE_FOLDER.mkdir(parents=True, exist_ok=True)
    xml_path = XML_DIR / f"{mrn}.xml"

    # Β.1 XML — από τον φάκελο αν υπάρχει, αλλιώς λήψη
    t_step = perf_counter()
    if xml_path.exists():
        log.info(f"  XML από τον φάκελο: {xml_path.name}")
    else:
        try:
            driver = phase_b_download_xml(mrn, driver)
        except Exception as e:
            log.exception(f"  Download XML failed: {e}"); save_screenshot(f"{mrn}_download_ERROR")
            return None, "fail", driver
        log.info(f"  [χρόνος] Download XML: {fmt_duration(perf_counter() - t_step)}")

    # Β.2 Parse XML + έλεγχος ότι το XML είναι όντως αυτού του MRN
    t_step = perf_counter()
    try:
        df_result = phase_b_parse_xml(xml_path)
        xml_mrn = str(df_result["MRN"].iloc[0]).strip() if len(df_result) else ""
        if xml_mrn != mrn:
            xml_path.unlink(missing_ok=True)
            raise ValueError(f"λάθος αρχείο — το XML είναι του {xml_mrn or '?'} (διαγράφηκε)")
    except Exception as e:
        log.exception(f"  Parse XML failed: {e}"); save_screenshot(f"{mrn}_parse_ERROR")
        return None, "fail", driver
    log.info(f"  [χρόνος] Parse XML: {fmt_duration(perf_counter() - t_step)}")

    # Φίλτρα
    tel_kath_val  = str(df_result["ΤΕΛ_ΚΑΘ"].iloc[0])
    dasmos_kl_val = str(df_result["ΔΑΣΜ_ΚΛ"].iloc[0])
    if tel_kath_val.startswith("71"):
        log.info(f"  ΤΕΛ_ΚΑΘ={tel_kath_val} — παραλείπεται"); return None, "skip", driver
    if not any(dasmos_kl_val.startswith(p) for p in ALLOWED_DASMOS):
        log.info(f"  ΔΑΣΜ_ΚΛ={dasmos_kl_val} — εκτός επιτρεπόμενων"); return None, "skip", driver

    # ── Κράματα (0832) — γεμίζουν ΜΕΣΑ στην ίδια γραμμή, όχι split ──────────
    needs_kramma = any(is_0832_kramma(r) for _, r in df_result.iterrows())
    if needs_kramma:
        num_eidi = len(df_result)
        if num_eidi == 1:
            aa = str(df_result.iloc[0]["Α/Α"])
            krammata = {aa: krammata_from_popup}
        else:
            log.info(f"  0832 με {num_eidi} είδη → ItemsPopup...")
            items_result = ItemsPopup(df_result, mrn).run()
            if items_result is None:
                log.error(f"  ItemsPopup ακυρώθηκε ({mrn})."); return None, "fail", driver
            krammata = items_result

        for i, r in df_result.iterrows():
            aa = str(r["Α/Α"])
            d = krammata.get(aa, {})
            for k in ("ΚΡΑΜΑ_1", "ΒΑΡΟΣ_1", "ΚΡΑΜΑ_2", "ΒΑΡΟΣ_2", "ΚΡΑΜΑ_3", "ΒΑΡΟΣ_3"):
                df_result.at[i, k] = d.get(k, "")

    log.info(f"\n{df_result[['Α/Α','ΔΑΣΜ_ΚΛ','ΒΑΡΟΣ','ΤΙΜΗ','ΦΠΑ']].to_string(index=False)}\n")

    # Όλα τα popups ΤΩΡΑ (πριν το SAP) — οι απαντήσεις μένουν στη γραμμή
    try:
        for i, r in df_result.iterrows():
            res = resolve_inputs(r, prot)
            for k, v in res.items():
                df_result.loc[i, "_PROT" if k == "PROT" else k] = v
    except Exception as e:
        log.exception(f"  Popups ({mrn}): {e}")
        return None, "fail", driver

    return {"mrn": mrn, "pdf": pdf, "pdf1": pdf1, "df_result": df_result}, "ok", driver


def enter_mrn(job: dict, session) -> bool:
    """
    Μέρος 2 (μόνο SAP, κανένα popup): PDF rename -> για κάθε Α/Α
    γέμισμα (με τις τιμές του Μέρους 1) -> Save -> attach -> διπλό F3 ->
    PDF move μία φορά στο τέλος. Σε σφάλμα: screenshot, close_sap, False.
    """
    mrn, pdf, pdf1, df_result = job["mrn"], job["pdf"], job["pdf1"], job["df_result"]
    log.info(f"\n{'='*60}\n  SAP — MRN : {mrn}\n{'='*60}")

    # Β.3 PDF rename
    dst_pdf = SAP_GUI_DIR / f"{pdf}.pdf"
    if pdf and pdf1:
        src_pdf = SAP_GUI_DIR / f"{pdf1}.pdf"
        if dst_pdf.exists():
            log.info(f"  PDF already renamed")
        elif src_pdf.exists():
            src_pdf.rename(dst_pdf); log.info(f"  PDF renamed: {pdf1} -> {pdf}")
        else:
            log.warning(f"  PDF not found: {src_pdf}"); dst_pdf = None
    elif not pdf:
        log.warning("  Δεν βρέθηκε PDF filename — παραλείπεται το attach.")
        dst_pdf = None

    # Β.4 SAP loop — κάθε γραμμή (Α/Α) = 1 ξεχωριστή καταχώρηση, ΧΩΡΙΣ Save.
    # Το session μένει ανοιχτό/logged-in σε όλο το batch (ΔΕΝ κλείνει ανά
    # καταχώρηση) — μεταξύ καταχωρήσεων γυρνάμε πίσω με διπλό F3.
    total = len(df_result)
    for idx, (i, row_item) in enumerate(df_result.iterrows(), 1):
        log.info(f"\n  Καταχώρηση {idx}/{total} | Α/Α={row_item['Α/Α']}")
        t_step = perf_counter()
        try:
            # Το XML/data είναι ήδη έτοιμα (Β.1/Β.2 παραπάνω) — η πλοήγηση SAP
            # (Επιλογή Πεδίου κλπ) γίνεται ΤΩΡΑ, ακριβώς πριν χρειαστεί.
            sap_ima_reenter_transaction(session)
            sap_ima_open_new_entry(session)
            ima_row = df_result.loc[i]
            sap_entry(session, ima_row, str(ima_row["_PROT"]))
        except Exception as e:
            log.exception(f"  SAP entry ΣΦΑΛΜΑ (Α/Α={row_item['Α/Α']}): {e}")
            save_screenshot(f"{mrn}_{idx:02d}_ERROR")
            close_sap(); return False
        log.info(f"  [χρόνος] SAP Entry: {fmt_duration(perf_counter() - t_step)}")

        t_step = perf_counter()
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        status_text = ""
        for _ in range(20):
            status_text = session.findById("wnd[0]/sbar/pane[0]").text
            if "Δεδομένα αποθηκεύτηκαν" in status_text:
                break
            time.sleep(0.5)
        else:
            log.exception(f"  Δεν επιβεβαιώθηκε το Save. Status bar: '{status_text}'")
            save_screenshot(f"{mrn}_{idx:02d}_SAVE_NOT_CONFIRMED")
            close_sap(); return False
        _SAVED["n"] += 1
        log.info(f"  Save επιβεβαιώθηκε (Α/Α={row_item['Α/Α']}): {status_text}  |  "
                 f"[χρόνος] Save: {fmt_duration(perf_counter() - t_step)}")

        t_step = perf_counter()
        try:
            if dst_pdf and dst_pdf.exists():
                sap_attach(session, dst_pdf)
        except Exception as e:
            log.exception(f"  Attach ΣΦΑΛΜΑ (Α/Α={row_item['Α/Α']}): {e}")
            save_screenshot(f"{mrn}_{idx:02d}_ATTACH_ERROR")
            close_sap(); return False
        log.info(f"  [χρόνος] Attach: {fmt_duration(perf_counter() - t_step)}")

        sap_ima_back_to_overview(session)
        # Η επόμενη επανάληψη (ή το επόμενο MRN) θα κάνει sap_ima_reenter_transaction()
        # ΜΕΤΑ που θα είναι έτοιμο το δικό της XML — όχι τώρα.

    # Β.5 PDF Move — ΜΙΑ φορά για ΟΛΗ τη δήλωση, ΜΕΤΑ από ΟΛΑ τα Α/Α (ίδιο
    # pattern με το ΕΙΣΑΓΩΓΕΣ ΠΕΙΡΑΙΑ NEW — ποτέ Move ενδιάμεσα, γιατί όλα τα
    # Α/Α μοιράζονται το ίδιο PDF).
    if dst_pdf and dst_pdf.exists():
        first_row = df_result.iloc[0]
        kath = compute_kath(str(first_row["X16"]), str(first_row["ΤΕΛ_ΚΑΘ"]))
        dest_subdir = KATH_DIR.get(kath)
        if dest_subdir:
            dest = ARCHIVE_BASE / dest_subdir / dst_pdf.name
            if dest.exists():
                log.info(f"  PDF ήδη υπάρχει στο network folder — δεν το αγγίζω: {dest}")
            else:
                try:
                    shutil.move(str(dst_pdf), str(dest))
                    log.info(f"  PDF moved -> {dest}")
                except Exception as e:
                    log.exception(f"  Move failed: {e}")
        else:
            log.warning(f"  Άγνωστο ΚΑΘ={kath} — δεν έγινε Move.")

    log.info("  Ολοκληρώθηκαν όλες οι καταχωρήσεις του MRN — ΜΕ Save, ΜΕ PDF Move.")
    return True

# ==============================================================================
# MAIN
# ==============================================================================

def ask_start_mode() -> int:
    root = tk.Tk()
    root.title("Εκκίνηση")
    root.resizable(False, False)
    root.configure(bg="#F5F4F0")
    root.attributes("-topmost", True)
    w, h = 520, 220
    sw = root.winfo_screenwidth(); sh = root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    choice = tk.IntVar(value=0)

    tk.Label(root, text="Από πού να ξεκινήσω;",
        bg="#F5F4F0", fg="#1A1A1A", font=("Consolas", 11, "bold")
    ).pack(pady=(20, 12))

    for val, label, bg, fg in [
        (1, "1 — Πλήρης ροή  (ICISnet + SAP Export + Queries + Popup)", "#1A1A1A", "#FFFFFF"),
        (2, "2 — Queries από αποθηκευμένο PDF + Database", "#2A2A2A", "#FFFFFF"),
        (3, "3 — Από Popup  (έχω ήδη FULL_RESULTS)", "#E8E6DF", "#1A1A1A"),
    ]:
        tk.Button(root, text=label,
            font=("Consolas", 10), bg=bg, fg=fg,
            relief="flat", padx=12, pady=7, cursor="hand2",
            command=lambda v=val: [choice.set(v), root.destroy()]
        ).pack(fill="x", padx=30, pady=(0, 5))

    root.lift(); root.focus_force()
    root.mainloop()
    return choice.get()


def read_icisnet_from_pdf() -> pd.DataFrame:
    import pdfplumber
    rows = []
    with pdfplumber.open(PDF_SAVE_PATH) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                parts = line.split()
                if len(parts) < 3:
                    continue
                if parts[0] in ("LRN", "Αποτελέσματα", "MRN"):
                    continue
                lrn = parts[0]
                mrn = parts[1] if len(parts) > 1 and parts[1].startswith("26GR") else ""
                typos = next((x for x in parts if x.startswith("IM-")), "")
                if not typos:
                    continue
                rows.append({"LRN": lrn, "MRN": mrn, "Τύπος Δήλωσης": typos})

    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["MRN","ΤΥΠΟΣ","ΚΑΤΑΣΤΑΣΗ","LRN","ΗΜ_ΥΠΟΒ","ΗΜ_ΕΝΗΜ","PDF"])

    repl = [
        ("YOUR_BROKER_ID/25/", "ELVYOUR_BROKER_ID/25/"), ("ELVELV", "ELV"),
        ("YOUR_BROKER_ID2/26/131ELB", "YOUR_BROKER_ID2/26/131ELV"),
    ]
    df["LRN"] = df["LRN"].astype(str)
    for old, new in repl:
        df["LRN"] = df["LRN"].str.replace(old, new, regex=False)

    df["Τύπος Δήλωσης"] = df["Τύπος Δήλωσης"].str.replace("-", "", regex=False)
    df["ΚΑΤΑΣΤΑΣΗ"] = "ID29"
    df["ΗΜ_ΥΠΟΒ"] = pd.NaT
    df["ΗΜ_ΕΝΗΜ"] = pd.NaT
    df["PDF"] = df["MRN"].astype(str) + " " + df["Τύπος Δήλωσης"].astype(str)
    df = df.rename(columns={"Τύπος Δήλωσης": "ΤΥΠΟΣ"})
    df = df[~df["LRN"].str.contains("XALELV", na=False)]
    df = df[df["LRN"].str.contains(r"ELV|ΕLV", na=False)]
    df = df[df["MRN"] != ""]

    return df[["MRN","ΤΥΠΟΣ","ΚΑΤΑΣΤΑΣΗ","LRN","ΗΜ_ΥΠΟΒ","ΗΜ_ΕΝΗΜ","PDF"]].reset_index(drop=True)


def fmt_duration(seconds: float) -> str:
    m, s = divmod(int(round(seconds)), 60)
    return f"{m}λ {s}δευτ" if m else f"{s}δευτ"


def main():
    T0 = perf_counter()
    setup_keyboard()

    mode = ask_start_mode()
    if mode == 0:
        return

    if mode == 1:
        print("\n── ΦΑΣΗ Α: ΕΙΣΑΓΩΓΕΣ ──")
        print("\n[1/3] ICISnet scraping...")
        df_final = phase_a_icisnet()
        print("\n[2/3] SAP Export...")
        phase_a_sap_export()
        print("\n[3/3] Queries -> FULL_RESULTS.xlsx...")
        df_opened = phase_a_queries(df_final)

    elif mode == 2:
        print("\n[Queries] από αποθηκευμένο PDF...")
        df_final = read_icisnet_from_pdf()
        print(f"  PDF: {len(df_final)} εγγραφές")
        df_opened = phase_a_queries(df_final)

    else:
        print("\n── Φόρτωση FULL_RESULTS.xlsx ──")
        df_opened = pd.read_excel(OUTPUT_EXCEL, sheet_name="opened")
        df_opened.columns = df_opened.columns.str.strip()

    if df_opened.empty:
        show_info("⚠️  Δεν υπάρχουν διασαφίσεις προς καταχώρηση.")
        return


    # ── POPUP ─────────────────────────────────────────────────────────────────
    print("\n── POPUP: Αναμονή χρήστη ──")

    rows = []
    for _, r in df_opened.iterrows():
        d = {}
        for col in SHOW_COLS:
            val = r.get(col, "")
            if pd.isna(val): val = ""
            if col == "PROT" and str(val).endswith(".0"):
                val = str(int(float(val)))
            d[col] = str(val).strip() if str(val).strip() != "nan" else ""
        for extra in df_opened.columns:
            if extra not in d:
                v = r.get(extra, "")
                d[extra] = "" if pd.isna(v) else v
        rows.append(d)
  
    selected = ApprovalPopup(rows).run()

    if selected is None:
        print("  Ακυρώθηκε.")
        return

    print(f"  {len(selected)} MRNs επιλέχθηκαν — εκκίνηση Φάσης Β...")

    df_excel = pd.read_excel(OUTPUT_EXCEL, sheet_name="opened")
    if "Status" not in df_excel.columns:
        df_excel["Status"] = ""

    def mark_done(mrn):
        try:
            df_excel.loc[df_excel["MRN"].astype(str) == str(mrn), "Status"] = "DONE"
            with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl",
                                mode="a", if_sheet_exists="replace") as w:
                df_excel.to_excel(w, sheet_name="opened", index=False)
            print(f"  Status -> DONE")
        except Exception as e:
            print(f"  Status update failed: {e}")

    ok_n = 0; fail_n = 0; skip_n = 0

    # ── ΦΑΣΗ Β.1: XML + popups για ΟΛΑ τα MRN (το SAP είναι κλειστό) ────────
    # Ο browser ανοίγει μόνο αν χρειαστεί λήψη (phase_b_download_xml κάνει
    # login όταν το driver δεν είναι ζωντανό). Σφάλμα σε ένα MRN εδώ δεν
    # σταματάει τα υπόλοιπα — δεν έχει αγγιχτεί ακόμα το SAP.
    print("\n── ΦΑΣΗ Β.1: XML + popups ──")
    driver = None
    jobs = []
    for i, row in enumerate(selected, 1):
        mrn = row.get("MRN", "???")
        print(f"\n[{i}/{len(selected)}] {mrn}")
        job, status, driver = prepare_mrn(row, driver)
        if status == "ok":
            jobs.append(job)
        elif status == "skip":
            skip_n += 1
            mark_done(mrn)
        else:
            fail_n += 1
            log.error(f"  FAIL (XML/popups) στο MRN={mrn} — δεν θα περαστεί στο SAP.")
    if driver is not None:
        try:
            driver.quit()
            print("  Browser closed")
        except Exception:
            pass

    # ── ΦΑΣΗ Β.2: SAP — ένα login, όλες οι καταχωρήσεις, κανένα popup ───────
    if jobs:
        print(f"\n── ΦΑΣΗ Β.2: SAP καταχωρήσεις ({len(jobs)} MRN) ──")
        print("  SAP login (μία φορά για όλο το batch)...")
        session = sap_ima_login()
        for i, job in enumerate(jobs, 1):
            print(f"\n[{i}/{len(jobs)}] {job['mrn']}")
            if enter_mrn(job, session):
                ok_n += 1
                mark_done(job["mrn"])
            else:
                fail_n += 1
                # enter_mrn() έχει ήδη κάνει close_sap() — με νεκρό session το
                # επόμενο MRN θα έσκαγε κι αυτό (confirmed live 2026-10-01).
                log.error(f"  FAIL στο MRN={job['mrn']} — σταματάω το batch (το SAP session έκλεισε ήδη).")
                break
        close_sap()
        print("  SAP closed")
    else:
        print("\n  Κανένα MRN για SAP.")

    elapsed = perf_counter() - T0
    total_n = ok_n + fail_n
    avg = elapsed / total_n if total_n else 0
    print(f"\n{'='*60}")
    print(f"  ΟΛΟΚΛΗΡΩΘΗΚΕ  |  OK:{ok_n}  FAIL:{fail_n}  ΠΑΡΑΛΕΙΨΗ:{skip_n}  |  "
          f"Σύνολο: {fmt_duration(elapsed)}  |  AVG: {fmt_duration(avg)}/MRN")
    print(f"{'='*60}")


if __name__ == "__main__":
    mutex = win32event.CreateMutex(None, False, SAP_MUTEX_NAME)
    print("Αναμονή SAP mutex (αν τρέχει άλλο script στο SAP)...")
    win32event.WaitForSingleObject(mutex, win32event.INFINITE)
    try:
        main()
    except Exception as e:
        log.exception(f"ΣΦΑΛΜΑ: {e}")
        raise
    finally:
        win32event.ReleaseMutex(mutex)
        _finish_log()
