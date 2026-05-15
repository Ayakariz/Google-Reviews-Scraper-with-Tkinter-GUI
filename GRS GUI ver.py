import csv
import os
import time
import re
import sys
import json
import socket
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from datetime import date, timedelta, datetime, timezone
from pathlib import Path

# ── Optional dependencies ─────────────────────────────────────────────────────
try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    from dateutil.relativedelta import relativedelta
    HAS_DATEUTIL = True
except ImportError:
    HAS_DATEUTIL = False

try:
    from dotenv import load_dotenv, set_key
    HAS_DOTENV = True
except ImportError:
    HAS_DOTENV = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException

try:
    from webdriver_manager.chrome import ChromeDriverManager
    USE_MANAGER = True
except ImportError:
    USE_MANAGER = False

# ── CONFIG ────────────────────────────────────────────────────────────────────
SCROLL_PAUSE = 2.5
ENV_FILE     = Path(".env")          # .env lives next to the script
ENV_KEY_NAME = "GOOGLE_PLACES_API_KEY"
# ─────────────────────────────────────────────────────────────────────────────

# ═════════════════════════════════════════════════════════════════════════════
#  INTERNET CHECK
# ═════════════════════════════════════════════════════════════════════════════

def is_online(host="8.8.8.8", port=53, timeout=3) -> bool:
    """
    Try to open a TCP socket to Google's DNS.
    Returns True if internet is reachable, False otherwise.
    """
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except (socket.error, OSError):
        return False


# ═════════════════════════════════════════════════════════════════════════════
#  .ENV HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def load_api_key_from_env() -> str:
    """Read GOOGLE_PLACES_API_KEY from .env file if it exists."""
    if not HAS_DOTENV:
        return ""
    if ENV_FILE.exists():
        load_dotenv(ENV_FILE, override=True)
    return os.getenv(ENV_KEY_NAME, "")


def save_api_key_to_env(api_key: str):
    """
    Write/update GOOGLE_PLACES_API_KEY in .env file.
    Creates the file if it doesn't exist.
    """
    if not HAS_DOTENV:
        return
    if not ENV_FILE.exists():
        ENV_FILE.touch()
    set_key(str(ENV_FILE), ENV_KEY_NAME, api_key)


# ═════════════════════════════════════════════════════════════════════════════
#  DATE CONVERSION
# ═════════════════════════════════════════════════════════════════════════════

def relative_to_date(relative: str) -> str:
    """Convert Google's relative timestamp to DD-MM-YYYY."""
    today = date.today()
    s = relative.lower().strip()
    num_match = re.search(r"(\d+)", s)
    n = int(num_match.group(1)) if num_match else 1

    if any(k in s for k in ["jam", "hour", "menit", "minute", "detik", "second"]):
        result = today
    elif any(k in s for k in ["hari", "day"]):
        result = today - timedelta(days=n)
    elif any(k in s for k in ["minggu", "week"]):
        result = today - timedelta(weeks=n)
    elif any(k in s for k in ["bulan", "month"]):
        result = (today - relativedelta(months=n)) if HAS_DATEUTIL else (today - timedelta(days=n * 30))
    elif any(k in s for k in ["tahun", "year"]):
        result = (today - relativedelta(years=n)) if HAS_DATEUTIL else (today - timedelta(days=n * 365))
    else:
        return relative

    return result.strftime("%d-%m-%Y")


# ═════════════════════════════════════════════════════════════════════════════
#  PLACE ID SEARCH
# ═════════════════════════════════════════════════════════════════════════════

def search_place_id(query: str, api_key: str) -> list:
    """
    Call Google Places API (New) Text Search.
    Raises RuntimeError with a user-friendly Indonesian message on any failure.
    """
    if not HAS_REQUESTS:
        raise RuntimeError(
            "Library 'requests' tidak ditemukan.\n"
            "Jalankan: py -m pip install requests"
        )

    if not is_online():
        raise RuntimeError(
            "❌ Tidak ada koneksi internet.\n\n"
            "Fitur Cari Place ID membutuhkan koneksi internet\n"
            "untuk mengakses Google Places API.\n\n"
            "Jika Anda sudah tahu Place ID-nya, langsung masukkan\n"
            "di tab Scraper."
        )

    url = "https://places.googleapis.com/v1/places:searchText"
    headers = {
        "Content-Type"    : "application/json",
        "X-Goog-Api-Key"  : api_key,
        "X-Goog-FieldMask": "places.id,places.displayName,places.formattedAddress,places.rating,places.userRatingCount",
    }

    try:
        resp = requests.post(url, json={"textQuery": query}, headers=headers, timeout=10)
    except requests.exceptions.ConnectionError:
        raise RuntimeError(
            "❌ Koneksi gagal.\n\n"
            "Tidak dapat terhubung ke server Google.\n"
            "Periksa koneksi internet Anda dan coba lagi."
        )
    except requests.exceptions.Timeout:
        raise RuntimeError(
            "❌ Koneksi timeout.\n\n"
            "Server Google tidak merespons dalam 10 detik.\n"
            "Coba lagi beberapa saat kemudian."
        )

    if resp.status_code == 400:
        raise RuntimeError(
            "❌ API Key tidak valid atau request salah.\n"
            "Periksa kembali API Key Anda."
        )
    if resp.status_code == 403:
        raise RuntimeError(
            "❌ API Key tidak punya akses ke Places API.\n"
            "Aktifkan 'Places API (New)' di Google Cloud Console."
        )
    if resp.status_code != 200:
        raise RuntimeError(
            f"❌ Error dari Google API: {resp.status_code}\n{resp.text}"
        )

    places = resp.json().get("places", [])
    return [
        {
            "name"   : p.get("displayName", {}).get("text", "—"),
            "address": p.get("formattedAddress", "—"),
            "id"     : p.get("id", ""),
            "rating" : p.get("rating", "N/A"), # Ambil Rating
            "reviews": p.get("userRatingCount", 0), # Ambil Jumlah Review
        }
        for p in places
    ]


# ═════════════════════════════════════════════════════════════════════════════
#  SELENIUM SCRAPER
# ═════════════════════════════════════════════════════════════════════════════

def build_driver() -> webdriver.Chrome:
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    # Aktifkan CDP performance logging untuk capture network response
    opts.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    if USE_MANAGER:
        try:
            if not is_online():
                raise RuntimeError(
                    "❌ Tidak ada koneksi internet.\n\n"
                    "webdriver-manager perlu mengunduh ChromeDriver.\n\n"
                    "Solusi:\n"
                    "1. Sambungkan ke internet lalu coba lagi, ATAU\n"
                    "2. Unduh ChromeDriver secara manual dari:\n"
                    "   https://chromedriver.chromium.org/downloads\n"
                    "   lalu letakkan chromedriver.exe di folder yang sama\n"
                    "   dengan script ini."
                )
            service = Service(ChromeDriverManager().install())
            driver  = webdriver.Chrome(service=service, options=opts)
        except RuntimeError:
            raise
        except Exception:
            # Fallback: try using chromedriver.exe in the same folder
            local_driver = Path(__file__).parent / "chromedriver.exe"
            if local_driver.exists():
                service = Service(str(local_driver))
                driver  = webdriver.Chrome(service=service, options=opts)
            else:
                raise RuntimeError(
                    "❌ ChromeDriver tidak ditemukan.\n\n"
                    "Unduh ChromeDriver dari:\n"
                    "https://chromedriver.chromium.org/downloads\n\n"
                    "Pastikan versinya sesuai dengan versi Chrome Anda,\n"
                    "lalu letakkan chromedriver.exe di folder yang sama\n"
                    "dengan script ini."
                )
    else:
        local_driver = Path(__file__).parent / "chromedriver.exe"
        if local_driver.exists():
            service = Service(str(local_driver))
            driver  = webdriver.Chrome(service=service, options=opts)
        else:
            driver = webdriver.Chrome(options=opts)

    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"}
    )
    enable_cdp_network(driver)
    return driver



def enable_cdp_network(driver):
    """Aktifkan CDP Network logging untuk capture response Google Maps."""
    driver.execute_cdp_cmd("Network.enable", {})
    print("[CDP] Network interception aktif — siap capture timestamp presisi.")


WIB_TZ = timezone(timedelta(hours=7))

# Rentang sanity-check timestamp UNIX (detik): 2010-01-01 .. 2030-01-01.
# Microsecond timestamps Google Maps berada di rentang ini setelah dibagi 1e6.
_TS_MIN_SEC = 1262304000
_TS_MAX_SEC = 1893456000


def _format_us_timestamp(ts_us: int) -> str:
    """Convert microsecond UNIX timestamp -> 'DD-MM-YYYY HH:MM' di WIB."""
    ts_sec = ts_us / 1_000_000
    dt = datetime.fromtimestamp(ts_sec, tz=timezone.utc).astimezone(WIB_TZ)
    return dt.strftime("%d-%m-%Y %H:%M")


def _format_utc_outscraper(ts_us: int) -> str:
    """Convert microsecond UNIX timestamp -> 'MM/DD/YYYY HH:MM:SS' di UTC,
    sesuai format kolom `review_datetime_utc` Outscraper."""
    ts_sec = ts_us / 1_000_000
    dt = datetime.fromtimestamp(ts_sec, tz=timezone.utc)
    return dt.strftime("%m/%d/%Y %H:%M:%S")


def _ts_us_to_seconds(ts_us: int) -> int:
    """Convert microsecond UNIX timestamp -> integer detik (kolom
    `review_timestamp` Outscraper)."""
    return int(ts_us // 1_000_000)


def _looks_like_review_id(s) -> bool:
    """
    Encoded review ID (yang sama dengan attribute data-review-id di DOM) berformat
    base64-url-safe panjang, contohnya:
      "Ci90QU1RQUNvZENodHl1Rj1vTzpCRVdFRnZhMUp4Y0ZVMVliWmtRM1Z..."
      "ChdDSUhNMG9nS0VJQ0FnSUNad0lha..."
    Heuristik: string panjang >= 20, hanya karakter [A-Za-z0-9_-], biasanya dimulai
    dengan "Ch" atau "Ci".
    """
    if not isinstance(s, str) or len(s) < 20:
        return False
    if not re.fullmatch(r"[A-Za-z0-9_\-]+", s):
        return False
    return True


def _extract_id_ts_pairs(obj, out: dict, depth: int = 0):
    """
    Telusuri response JSON Google Maps; setiap kali menemukan review entry, simpan
    pemetaan {review_id_encoded: "DD-MM-YYYY HH:MM"} di `out`.

    Bentuk review entry yang dikenali (lihat frame 19 video user):
      entry[0] = "<encoded_review_id>"                  ← cocok dgn data-review-id
      entry[1] = ["0x...:0x...", null, null, <ts_us>, <ts_us>, ...]
      entry[3] = [null, <edit_ts_us>, ..., "a day ago"] (opsional, jika di-edit)
    """
    if depth > 12 or not isinstance(obj, list):
        return

    # Pola 1: entry langsung dari array index [7] response pc?authuser=
    if (
        len(obj) >= 2
        and _looks_like_review_id(obj[0])
        and isinstance(obj[1], list)
        and len(obj[1]) >= 4
        and isinstance(obj[1][0], str)
        and obj[1][0].startswith("0x")
    ):
        ts_us = None
        # Cari int 16-digit di obj[1] (post timestamp). Posisi observasi: [3]/[4].
        for v in obj[1]:
            if isinstance(v, int) and len(str(v)) == 16:
                sec = v / 1_000_000
                if _TS_MIN_SEC < sec < _TS_MAX_SEC:
                    ts_us = v
                    break
        if ts_us is not None:
            # Simpan microsecond mentah supaya pemanggil bisa derive multiple
            # representasi (DD-MM-YYYY WIB, UNIX seconds, MM/DD/YYYY UTC, dst.)
            out[obj[0]] = ts_us
            return  # entry diserap, jangan rekursi ke dalam supaya tidak overwrite

    # Pola 2 (fallback): nested — beberapa response membungkus entry di list lain.
    for item in obj:
        if isinstance(item, list):
            _extract_id_ts_pairs(item, out, depth + 1)


def extract_timestamps_from_logs(driver) -> dict:
    """
    Baca CDP performance logs dari request Google Maps yang membawa data review,
    lalu kembalikan {review_id_encoded: ts_us_int} — microsecond UNIX timestamp
    mentah. Caller boleh format ke berbagai representasi:
      _format_us_timestamp(us)    -> 'DD-MM-YYYY HH:MM' (WIB)
      _format_utc_outscraper(us)  -> 'MM/DD/YYYY HH:MM:SS' (UTC, format Outscraper)
      _ts_us_to_seconds(us)       -> int detik (kolom review_timestamp Outscraper)

    review_id_encoded sama persis dgn attribute `data-review-id` di DOM.

    Catatan: fungsi ini memanggil driver.get_log("performance") yang akan
    mengosongkan buffer log Chrome — jadi panggil sekali per scroll iteration.
    """
    mapping: dict = {}

    try:
        logs = driver.get_log("performance")
    except Exception:
        return mapping

    # Kumpulkan requestId dari endpoint yang membawa data review.
    relevant_request_ids = {}
    for entry in logs:
        try:
            msg = json.loads(entry["message"])["message"]
            if msg.get("method") != "Network.responseReceived":
                continue
            url = msg.get("params", {}).get("response", {}).get("url", "")
            if (
                "pc?authuser" in url
                or "listugcposts" in url
                or "listentitiesreviews" in url
                or ("/maps/rpc/" in url and "google.com" in url)
            ):
                req_id = msg.get("params", {}).get("requestId")
                if req_id:
                    relevant_request_ids[req_id] = url
        except Exception:
            continue

    for req_id in relevant_request_ids:
        try:
            result = driver.execute_cdp_cmd(
                "Network.getResponseBody", {"requestId": req_id}
            )
        except Exception:
            continue
        body = result.get("body", "")
        if not body:
            continue

        # Strategi 1: parse JSON, walk struktur untuk extract pasangan id↔timestamp.
        try:
            clean_body = re.sub(r"^\)\]\}['\"]?\s*", "", body.strip())
            data = json.loads(clean_body)
            _extract_id_ts_pairs(data, mapping)
        except (json.JSONDecodeError, ValueError):
            # Strategi 2 (fallback regex): cari pasangan ["<encoded_id>", ...,
            # "0x...", ..., <ts_us>] secara tekstual. Tidak seakurat strategi 1
            # tapi bisa menyelamatkan respons yang malformed.
            for m in re.finditer(
                r'"(C[hi][A-Za-z0-9_\-]{18,})"[^]]{0,400}?"0x[0-9a-fA-F:]+x?[0-9a-fA-F]*"[^]]{0,200}?(1[6-9]\d{14})',
                body,
            ):
                rid, ts_us = m.group(1), int(m.group(2))
                sec = ts_us / 1_000_000
                if _TS_MIN_SEC < sec < _TS_MAX_SEC and rid not in mapping:
                    mapping[rid] = ts_us

    if mapping:
        print(f"[CDP] +{len(mapping)} timestamp presisi terdeteksi pada batch ini.")
    return mapping
def clean_username(name: str) -> str:
    return re.sub(r"^(Foto|Photo of|Photo)\s+", "", name, flags=re.IGNORECASE).strip()


def google_login(driver, stop_check=None):
    print("\n" + "="*60)
    print("  STEP 1: GOOGLE LOGIN")
    print("="*60)

    if not is_online():
        raise RuntimeError(
            "❌ Tidak ada koneksi internet.\n\n"
            "Scraper membutuhkan internet untuk membuka Google Maps."
        )

    print("[*] Membuka halaman login Google...")
    driver.get("https://accounts.google.com/")
    time.sleep(3)
    print("[!] Silakan login ke akun Google Anda di browser.")
    print("[!] Script akan otomatis lanjut setelah login berhasil (max 120 detik)...")

    for _ in range(120):
        if stop_check and stop_check():
            print("[!] Proses login dihentikan oleh pengguna.")
            return
        if "accounts.google.com" not in driver.current_url:
            print("[+] Login terdeteksi! Melanjutkan proses...")
            time.sleep(2)
            return
        time.sleep(1)

    print("[!] Timeout login.")
    messagebox.showinfo(
        "Login Timeout",
        "Waktu tunggu login habis.\n\nJika Anda sudah berhasil login, klik OK untuk melanjutkan."
    )
    time.sleep(2)


def open_reviews_tab(driver, place_id):
    place_url = f"https://www.google.com/maps/place/?q=place_id:{place_id}"
    print("\n" + "="*60)
    print("  STEP 2: MEMBUKA GOOGLE MAPS PLACE")
    print("="*60)
    print(f"[*] Menuju ke: {place_url}")
    driver.get(place_url)
    time.sleep(5)

    selectors = [
        "//button[contains(@aria-label,'Reviews') or contains(@aria-label,'Ulasan')]",
        "//button[@data-tab-index='1']",
        "//div[@role='tab'][.//span[contains(text(),'Review') or contains(text(),'Ulasan')]]",
        "//button[.//div[contains(text(),'Reviews') or contains(text(),'Ulasan')]]",
    ]
    clicked = False
    for sel in selectors:
        try:
            el = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, sel)))
            el.click()
            clicked = True
            print("[+] Tab Ulasan berhasil diklik otomatis.")
            break
        except TimeoutException:
            pass

    if not clicked:
        print("[!] Gagal mengklik tab Ulasan secara otomatis.")
        messagebox.showinfo(
            "Tindakan Diperlukan",
            "Gagal membuka tab ulasan secara otomatis.\n\n"
            "Silakan klik tab 'Ulasan' / 'Reviews' secara manual, lalu klik OK."
        )
    time.sleep(3)


def sort_by_newest(driver):
    print("[*] Mencoba mengurutkan ulasan ke Terbaru...")
    try:
        sort_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH,
                "//button[contains(@aria-label,'Urutkan') or contains(@aria-label,'Sort') or @data-value='Sort']"))
        )
        sort_btn.click()
        time.sleep(2)
        newest_opt = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.XPATH,
                "//*[@role='menuitemradio' and (@data-index='1' or contains(.,'Newest') or contains(.,'Terbaru'))]"))
        )
        newest_opt.click()
        print("[+] Berhasil diurutkan berdasarkan Terbaru.")
        time.sleep(3)
    except Exception as e:
        print(f"[!] Tidak bisa mengurutkan otomatis ({e}). Melanjutkan dengan urutan default.")


def parse_stars(review_el) -> str:
    try:
        star_span = review_el.find_element(By.XPATH, ".//span[@role='img' and @aria-label]")
        label = star_span.get_attribute("aria-label")
        match = re.search(r"(\d+(?:\.\d+)?)", label)
        return match.group(1) if match else ""
    except NoSuchElementException:
        return ""


def parse_text(review_el) -> str:
    selectors = [".//span[@class='wiI7pd']", ".//div[@class='MyEned']//span", ".//span[contains(@class, 'review-full-text')]"]
    
    for sel in selectors:
        try:
            txt = review_el.find_element(By.XPATH, sel).text.strip()
            if txt and not txt.endswith("..."):
                return re.sub(r'\s+', ' ', txt).strip()
        except NoSuchElementException:
            continue
    return ""


def parse_timestamp(review_el) -> str:
    for sel in [".//span[@class='rsqaWe']", ".//span[contains(@class,'dehysf')]",
                ".//span[contains(@class,'review-snippet')]", ".//span[@class='xRkPPb']"]:
        try:
            ts = review_el.find_element(By.XPATH, sel).text.strip()
            if ts:
                return ts
        except NoSuchElementException:
            pass
    return ""


def parse_username(review_el) -> str:
    try:
        btn = review_el.find_element(By.XPATH, ".//button[contains(@aria-label,' ') and @data-href]")
        label = btn.get_attribute("aria-label")
        if label and label.strip():
            return clean_username(label)
    except NoSuchElementException:
        pass

    for cls in ["d4r55", "al6Kxe", "RfnDt", "bHjbKc", "NfpBac", "Jtu6Td"]:
        try:
            name = review_el.find_element(By.XPATH, f".//div[@class='{cls}']").text.strip()
            if name:
                return clean_username(name)
        except NoSuchElementException:
            pass

    try:
        link = review_el.find_element(
            By.XPATH, ".//a[contains(@href,'maps/contrib') or contains(@href,'/contrib/')]")
        name = link.text.strip()
        if name:
            return clean_username(name)
        child = link.find_element(By.XPATH, ".//*[string-length(text()) > 0]")
        name = child.text.strip()
        if name:
            return clean_username(name)
    except NoSuchElementException:
        pass

    try:
        for div in review_el.find_elements(By.XPATH, ".//div[string-length(text()) > 2]")[:5]:
            t = div.text.strip()
            if t and not re.search(r"\d", t) and len(t) < 60:
                return t
    except Exception:
        pass

    return ""


def scrape_reviews(driver, max_scrolls, stop_check=None, app=None) -> list:
    print("\n" + "="*60)
    print("  STEP 3: SCRAPING REVIEWS")
    print("="*60)

    seen_ids    = set()
    all_reviews = []
    last_count  = 0
    stale_iters = 0

    # Mapping {data-review-id -> "DD-MM-YYYY HH:MM" (WIB)} dari response
    # pc?authuser=. Diisi inkremental tiap iterasi scroll. Reviews yang sudah
    # tersimpan dgn fallback relative-date akan di-upgrade ke timestamp presisi
    # begitu ID-nya muncul di mapping.
    precise_ts: dict = {}

    print(f"[*] Melakukan scroll hingga {max_scrolls} kali...\n")

    for scroll_num in range(max_scrolls):
        if stop_check and stop_check():
            print(f"\n[!] Scraping dihentikan pada scroll ke-{scroll_num+1}.")
            break

        # Cukup satu XPath yang mencakup semua variasi tombol expand
        try:
            expand_xpath = "//button[contains(@aria-label, 'more') or contains(@aria-label, 'Lainnya') or @class='w8B6Ac']"
            for btn in driver.find_elements(By.XPATH, expand_xpath):
                driver.execute_script("arguments[0].click();", btn)
            time.sleep(0.5)
        except:
            pass

        # Drain CDP performance log dulu — supaya request pc?authuser= dari
        # scroll sebelumnya ter-capture sebelum buffer di-overwrite.
        try:
            new_ts = extract_timestamps_from_logs(driver)
            if new_ts:
                precise_ts.update(new_ts)
        except Exception as e:
            print(f"[CDP] gagal ambil timestamp presisi: {e}")

        for card in driver.find_elements(By.XPATH, "//div[@data-review-id]"):
            rid = card.get_attribute("data-review-id")
            if rid in seen_ids:
                continue
            seen_ids.add(rid)
            raw_ts = parse_timestamp(card)
            ts_us = precise_ts.get(rid)
            all_reviews.append({
                "_rid"               : rid,
                "username"           : parse_username(card),
                "stars"              : parse_stars(card),
                "date"               : _format_us_timestamp(ts_us) if ts_us else (relative_to_date(raw_ts) if raw_ts else ""),
                "review_timestamp"   : _ts_us_to_seconds(ts_us) if ts_us else "",
                "review_datetime_utc": _format_utc_outscraper(ts_us) if ts_us else "",
                "review_text"        : parse_text(card),
            })

        # Upgrade pass: review yg tadinya cuma punya relative-date (atau kosong)
        # akan diisi/diganti dgn timestamp presisi begitu ID-nya muncul di
        # `precise_ts` di batch berikutnya.
        for r in all_reviews:
            ts_us = precise_ts.get(r["_rid"])
            if not ts_us:
                continue
            r["date"]                = _format_us_timestamp(ts_us)
            r["review_timestamp"]    = _ts_us_to_seconds(ts_us)
            r["review_datetime_utc"] = _format_utc_outscraper(ts_us)

        count = len(all_reviews)

        if app:
            pct = int((scroll_num + 1) / max_scrolls * 100)
            app.root.after(0, lambda p=pct: app.progress.config(value=p))

        print(f"  Scroll {scroll_num+1:>3} | {count} ulasan terkumpul...")

        cards = driver.find_elements(By.XPATH, "//div[@data-review-id]")
        if cards:
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", cards[-1])
            except Exception:
                pass

        for _ in range(int(SCROLL_PAUSE * 10)):
            if stop_check and stop_check():
                break
            time.sleep(0.1)

        if count == last_count:
            stale_iters += 1
            if stale_iters >= 3:
                print(f"\n[+] Tidak ada ulasan baru setelah {stale_iters} scroll — selesai.")
                break
        else:
            stale_iters = 0
        last_count = count

    # Final drain — response dari scroll terakhir kadang baru masuk buffer
    # setelah loop berhenti.
    try:
        new_ts = extract_timestamps_from_logs(driver)
        if new_ts:
            precise_ts.update(new_ts)
            for r in all_reviews:
                ts_us = precise_ts.get(r["_rid"])
                if ts_us:
                    r["date"]                = _format_us_timestamp(ts_us)
                    r["review_timestamp"]    = _ts_us_to_seconds(ts_us)
                    r["review_datetime_utc"] = _format_utc_outscraper(ts_us)
    except Exception:
        pass

    # Statistik akurasi.
    precise_count = sum(
        1 for r in all_reviews if isinstance(r.get("review_timestamp"), int)
    )
    print(
        f"\n[+] Total ulasan terkumpul: {len(all_reviews)} "
        f"({precise_count} dgn timestamp presisi, "
        f"{len(all_reviews) - precise_count} fallback ke tanggal relatif)"
    )
    return all_reviews


CSV_FIELDNAMES = [
    "username", "stars", "date",
    "review_timestamp", "review_datetime_utc",
    "review_text",
]


def save_csv(reviews: list, path: str):
    if not reviews:
        print("[!] Tidak ada ulasan untuk disimpan.")
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDNAMES)
        writer.writeheader()
        # Buang field internal (_rid) sebelum tulis.
        for r in reviews:
            writer.writerow({k: r.get(k, "") for k in CSV_FIELDNAMES})
    print(f"[+] Berhasil disimpan (CSV): {len(reviews)} ulasan -> {path}")


def save_xlsx(reviews: list, path: str):
    """Simpan reviews ke file .xlsx. Memerlukan paket `openpyxl`."""
    if not reviews:
        print("[!] Tidak ada ulasan untuk disimpan.")
        return
    if not HAS_OPENPYXL:
        raise RuntimeError(
            "Paket 'openpyxl' tidak ditemukan.\n"
            "Jalankan: py -m pip install openpyxl"
        )
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reviews"
    ws.append(CSV_FIELDNAMES)
    for r in reviews:
        row = []
        for k in CSV_FIELDNAMES:
            v = r.get(k, "")
            # Excel happiest with str/int/float; coerce empty stays empty.
            row.append(v)
        ws.append(row)
    # Auto-size kolom (heuristik kecil — cukup untuk readability).
    for col_idx, name in enumerate(CSV_FIELDNAMES, start=1):
        max_len = max(
            [len(str(r.get(name, ""))) for r in reviews] + [len(name)]
        )
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = min(max_len + 2, 60)
    wb.save(path)
    print(f"[+] Berhasil disimpan (XLSX): {len(reviews)} ulasan -> {path}")


def save_reviews(reviews: list, base_path: str, fmt: str):
    """
    Simpan reviews dgn format yg dipilih user.
      fmt = "csv"  -> hanya .csv
      fmt = "xlsx" -> hanya .xlsx
      fmt = "both" -> kedua format (file dgn nama dasar yg sama).
    `base_path` adalah path lengkap pilihan user; ekstensinya akan disesuaikan.
    Return: list path file yg berhasil ditulis.
    """
    base, _ = os.path.splitext(base_path)
    written = []
    if fmt in ("csv", "both"):
        p = base + ".csv"
        save_csv(reviews, p)
        written.append(p)
    if fmt in ("xlsx", "both"):
        p = base + ".xlsx"
        save_xlsx(reviews, p)
        written.append(p)
    return written


# ═════════════════════════════════════════════════════════════════════════════
#  GUI
# ═════════════════════════════════════════════════════════════════════════════

class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, s):
        self.widget.insert(tk.END, s)
        self.widget.see(tk.END)

    def flush(self):
        pass


class ScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Maps Review Scraper")
        self.root.geometry("640x660")
        self.root.minsize(520, 500)
        self.root.resizable(True, True)          # ← resizable window

        self.stop_requested  = False
        self.original_stdout = sys.stdout

        # ── Notebook ─────────────────────────────────────────────────────────
        notebook = ttk.Notebook(root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_search  = ttk.Frame(notebook, padding=15)
        self.tab_scraper = ttk.Frame(notebook, padding=15)
        notebook.add(self.tab_scraper, text="⚙️  Scraper")
        notebook.add(self.tab_search,  text="🔍  Cari Place ID")
        
        self._build_search_tab()
        self._build_scraper_tab()

        # Offline warning banner (shown at bottom of window)
        self.banner_var = tk.StringVar()
        self.banner = tk.Label(
            root, textvariable=self.banner_var,
            bg="#fff3cd", fg="#856404",
            font=("Arial", 9), anchor="w", padx=10
        )
        # Check internet on startup (non-blocking)
        threading.Thread(target=self._check_online_banner, daemon=True).start()

        sys.stdout = TextRedirector(self.log_box)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── OFFLINE BANNER ────────────────────────────────────────────────────────

    def _check_online_banner(self):
        online = is_online()
        def _update():
            if not online:
                self.banner_var.set(
                    "⚠️  Tidak ada koneksi internet — fitur Cari Place ID dan Scraper tidak akan berfungsi."
                )
                self.banner.pack(fill=tk.X, side=tk.BOTTOM, before=self.root.winfo_children()[0])
            else:
                self.banner.pack_forget()
        self.root.after(0, _update)

    # ── TAB 1: PLACE ID SEARCH ────────────────────────────────────────────────

    def _build_search_tab(self):
        tab = self.tab_search
        # Mengatur berat kolom agar Nama (0) lebih lebar dari Kota (1)
        tab.columnconfigure(0, weight=3) 
        tab.columnconfigure(1, weight=1)
        tab.columnconfigure(2, weight=0) # Tombol tidak perlu melebar

        # API Key (Internal)
        self._api_key = load_api_key_from_env()

        # --- BARIS 2: LABEL ---
        # sticky="sw" memastikan teks menempel ke kiri bawah (tepat di atas entry)
        ttk.Label(tab, text="Nama Tempat:").grid(row=2, column=0, sticky="sw")
        ttk.Label(tab, text="Kota (Opsional):").grid(row=2, column=1, sticky="sw", padx=(5, 0))

        # --- BARIS 3: ENTRY & BUTTON ---
        self.entry_search_query = ttk.Entry(tab)
        self.entry_search_query.grid(row=3, column=0, sticky="ew", padx=(0, 5))
        self.entry_search_query.bind("<Return>", lambda e: self._do_search())

        self.entry_city = ttk.Entry(tab)
        self.entry_city.grid(row=3, column=1, sticky="ew", padx=(5, 5))
        self.entry_city.insert(0, "Jakarta")

        ttk.Button(tab, text="🔍 Cari", command=self._do_search).grid(row=3, column=2)

        # --- MENU KLIK KANAN (CONTEXT MENU) ---
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label=" 🌎Buka di Google Maps", command=self._open_in_maps)
        
        # --- RESULTS TREEVIEW ---
        ttk.Label(tab, text="Hasil Pencarian:").grid(row=4, column=0, sticky="w", pady=(10, 0))
        
        tree_frame = ttk.Frame(tab)
        tree_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(2, 6))
        tab.rowconfigure(5, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        cols = ("name", "rating", "reviews", "address", "place_id")
        self.result_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=8)

        self.result_tree["displaycolumns"] = ("name", "rating", "reviews", "address")
        self.result_tree.heading("name",     text="Nama Tempat")
        self.result_tree.heading("rating",   text="⭐")
        self.result_tree.heading("reviews",  text="Ulasan")
        self.result_tree.heading("address",  text="Alamat")
        # self.result_tree.heading("place_id", text="Place ID")
        
        # Konfigurasi Kolom: (Header, Lebar, Anchor, Stretch)
        col_settings = {
            "name":    ("Nama Tempat", 200, "w", True),
            "rating":  ("⭐", 50, "center", False),
            "reviews": ("Ulasan", 80, "center", False),
            "address": ("Alamat", 400, "w", True)
        }

        for col, (txt, w, anc, strc) in col_settings.items():
            self.result_tree.heading(col, text=txt, 
                command=lambda c=col: self._treeview_sort_column(self.result_tree, c, False))
            self.result_tree.column(col, width=w, minwidth=w//2, stretch=strc, anchor=anc)

        # Tambahkan binding klik kanan di sini
        self.result_tree.bind("<Button-3>", self._show_context_menu)
        self.result_tree.bind("<Double-1>", self._on_result_select)

        # Scrollbar dan elemen lainnya tetap sama seperti sebelumnya...
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=vsb.set)
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        # Hints & link
        ttk.Label(tab, text="💡 Double-click hasil untuk mengisi Place ID di tab Scraper.",
                  foreground="gray").grid(row=6, column=0, sticky="w")

        link = tk.Label(tab, text="Cara mendapatkan API Key →",
                        fg="blue", cursor="hand2", font=("Arial", 9, "underline"))
        link.grid(row=7, column=0, sticky="w", pady=(6, 0))
        link.bind("<Button-1>", lambda e: webbrowser.open_new(
            "https://developers.google.com/maps/documentation/places/web-service/get-api-key"))

    def _do_search(self):
        api_key = self._api_key
        query_base = self.entry_search_query.get().strip()
        city_name  = self.entry_city.get().strip()

        if not api_key:
            messagebox.showerror("Error", "API Key tidak ditemukan di .env")
            return
        if not query_base:
            messagebox.showerror("Error", "Masukkan nama tempat.")
            return

        # Gabungkan query: "Nama Tempat + Nama Kota"
        full_query = f"{query_base} {city_name}" if city_name else query_base

        for row in self.result_tree.get_children():
            self.result_tree.delete(row)

        def _thread():
            try:
                # Mengirim full_query ke fungsi search_place_id
                results = search_place_id(full_query, api_key)
                if not results:
                    self.root.after(0, lambda: messagebox.showinfo("Hasil", "Tidak ditemukan."))
                    return
                for r in results:
                    self.root.after(0, lambda row=r: self.result_tree.insert(
                        "", tk.END, values=(row["name"], row["rating"], row["reviews"], row["address"], row["id"])))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

        threading.Thread(target=_thread, daemon=True).start()

    def _on_result_select(self, event):
        selected = self.result_tree.selection()
        if not selected:
            return
        values = self.result_tree.item(selected[0], "values")
        if not values:
            return
        place_id, name = values[4], values[0]
        self.entry_place_id.delete(0, tk.END)
        self.entry_place_id.insert(0, place_id)
        messagebox.showinfo("Place ID Dipilih",
            f"Place ID untuk:\n'{name}'\n\n{place_id}\n\nSudah diisi di tab Scraper. ✓")

    # ── TAB 2: SCRAPER ────────────────────────────────────────────────────────

    def _build_scraper_tab(self):
        tab = self.tab_scraper
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(9, weight=1)   # log box stretches

        ttk.Label(tab, text="Place ID:").grid(row=0, column=0, sticky="w")
        self.entry_place_id = ttk.Entry(tab)
        self.entry_place_id.insert(0, "ChIJ43GkB4DtaS4RleDeD0fIP6g")
        self.entry_place_id.grid(row=1, column=0, sticky="ew", pady=(2, 12))

        ttk.Label(tab, text="Jumlah Maksimal Scroll:").grid(row=2, column=0, sticky="w")
        self.entry_scrolls = ttk.Entry(tab, width=15)
        self.entry_scrolls.insert(0, "20")
        self.entry_scrolls.grid(row=3, column=0, sticky="w", pady=(2, 12))

        # Format output (CSV / XLSX / BOTH)
        ttk.Label(tab, text="Format Output:").grid(row=4, column=0, sticky="w")
        fmt_frame = ttk.Frame(tab)
        fmt_frame.grid(row=5, column=0, sticky="w", pady=(2, 12))
        self.output_format = tk.StringVar(value="csv")
        ttk.Radiobutton(
            fmt_frame, text="CSV", value="csv", variable=self.output_format
        ).grid(row=0, column=0, padx=(0, 12))
        xlsx_text = "XLSX" if HAS_OPENPYXL else "XLSX (perlu openpyxl)"
        self.rb_xlsx = ttk.Radiobutton(
            fmt_frame, text=xlsx_text, value="xlsx", variable=self.output_format
        )
        self.rb_xlsx.grid(row=0, column=1, padx=(0, 12))
        self.rb_both = ttk.Radiobutton(
            fmt_frame, text="Keduanya (CSV + XLSX)", value="both",
            variable=self.output_format,
        )
        self.rb_both.grid(row=0, column=2)
        if not HAS_OPENPYXL:
            self.rb_xlsx.state(["disabled"])
            self.rb_both.state(["disabled"])

        btn_frame = ttk.Frame(tab)
        btn_frame.grid(row=6, column=0, sticky="ew", pady=(0, 8))
        btn_frame.columnconfigure((0, 1), weight=1)

        self.btn_run = ttk.Button(btn_frame, text="▶ Mulai Scraping", command=self.start_scraping)
        self.btn_run.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.btn_stop = ttk.Button(btn_frame, text="⏹ Stop Aktivitas",
                                   command=self.stop_scraping, state="disabled")
        self.btn_stop.grid(row=0, column=1, sticky="ew", padx=(4, 0))

        self.progress = ttk.Progressbar(tab, maximum=100, value=0)
        self.progress.grid(row=7, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(tab, text="Log Aktivitas:").grid(row=8, column=0, sticky="w")
        self.log_box = scrolledtext.ScrolledText(tab, bg="#f4f4f4", font=("Consolas", 9))
        self.log_box.grid(row=9, column=0, sticky="nsew", pady=(2, 0))
        tab.rowconfigure(9, weight=1)

    def _treeview_sort_column(self, tv, col, reverse):
        # Ambil semua data di kolom yang diklik
        l = [(tv.set(k, col), k) for k in tv.get_children('')]

        # Logika sorting khusus untuk angka (Rating & Reviews)
        try:
            # Coba convert ke float untuk rating (misal 4.2) atau int
            l.sort(key=lambda t: float(t[0]), reverse=reverse)
        except ValueError:
            # Jika gagal (berarti teks), gunakan sort standar
            l.sort(reverse=reverse)

        # Re-arrange baris sesuai urutan baru
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        # Update perintah klik agar klik berikutnya membalikkan urutan (Asc/Desc)
        tv.heading(col, command=lambda: self._treeview_sort_column(tv, col, not reverse))

    def _show_context_menu(self, event):
        # Mengidentifikasi baris mana yang diklik kanan
        item_id = self.result_tree.identify_row(event.y)
        if item_id:
            # Otomatis pilih baris tersebut saat diklik kanan
            self.result_tree.selection_set(item_id)
            # Tampilkan menu popup di posisi kursor
            self.context_menu.post(event.x_root, event.y_root)

    def _open_in_maps(self):
        selected = self.result_tree.selection()
        if not selected:
            return
        
        # Ambil nilai dari kolom Place ID (index ke-2)
        values = self.result_tree.item(selected[0], "values")
        if values:
            place_id = values[4]
            # URL resmi Google Maps untuk koordinat Place ID
            maps_url = f"https://www.google.com/maps/search/?api=1&query=Google&query_place_id={place_id}"
            webbrowser.open(maps_url)

    # ── SCRAPER ACTIONS ───────────────────────────────────────────────────────

    def start_scraping(self):
        place_id = self.entry_place_id.get().strip()

        if "place_id:" in place_id:
            place_id = place_id.split("place_id:")[-1].strip().split("&")[0]
        elif "google.com/maps" in place_id:
            messagebox.showerror("Error",
                "Masukkan Place ID saja, bukan URL penuh.\nContoh: ChIJ43GkB4DtaS4RleDeD0fIP6g")
            return

        if not place_id:
            messagebox.showerror("Error", "Place ID tidak boleh kosong!")
            return

        try:
            max_scrolls = int(self.entry_scrolls.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Jumlah scroll harus berupa angka!")
            return

        fmt = self.output_format.get()
        if fmt in ("xlsx", "both") and not HAS_OPENPYXL:
            messagebox.showerror(
                "Error",
                "Format XLSX dipilih tetapi paket 'openpyxl' belum terpasang.\n"
                "Jalankan: py -m pip install openpyxl"
            )
            return

        # Sesuaikan dialog Save berdasarkan format yang dipilih.
        if fmt == "xlsx":
            default_ext = ".xlsx"
            filetypes = [("Excel files", "*.xlsx"), ("All files", "*.*")]
            initialfile = "google_reviews.xlsx"
        elif fmt == "both":
            # User memilih basename; ekstensi akan dipakai sbg base saja.
            default_ext = ""
            filetypes = [("All files", "*.*")]
            initialfile = "google_reviews"
        else:  # csv
            default_ext = ".csv"
            filetypes = [("CSV files", "*.csv"), ("All files", "*.*")]
            initialfile = "google_reviews.csv"

        dialog_title = (
            "Simpan hasil sebagai... (akan dibuat .csv DAN .xlsx)"
            if fmt == "both" else "Simpan hasil sebagai..."
        )
        save_path = filedialog.asksaveasfilename(
            title=dialog_title,
            defaultextension=default_ext,
            filetypes=filetypes,
            initialfile=initialfile,
        )
        if not save_path:
            return

        self.stop_requested = False
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progress.config(value=0)
        self.log_box.delete("1.0", tk.END)

        threading.Thread(
            target=self.run_process,
            args=(place_id, max_scrolls, save_path, fmt),
            daemon=True
        ).start()

    def stop_scraping(self):
        print("\n[!] Perintah STOP diterima. Menunggu proses saat ini selesai...")
        self.stop_requested = True
        self.btn_stop.config(state="disabled")

    def run_process(self, place_id, max_scrolls, save_path, fmt="csv"):
        print("Mempersiapkan browser...")
        driver = None
        try:
            driver = build_driver()
            check_stop = lambda: self.stop_requested

            if not self.stop_requested:
                google_login(driver, check_stop)
            if not self.stop_requested:
                open_reviews_tab(driver, place_id)
            if not self.stop_requested:
                sort_by_newest(driver)
            if not self.stop_requested:
                reviews = scrape_reviews(driver, max_scrolls, check_stop, self)
                if reviews:
                    written = save_reviews(reviews, save_path, fmt)
                    files_str = "\n".join(written)
                    title = "Dihentikan" if self.stop_requested else "Selesai ✓"
                    head = (
                        "Scraping dihentikan secara manual."
                        if self.stop_requested else "Scraping selesai!"
                    )
                    msg = f"{head}\n{len(reviews)} ulasan disimpan ke:\n{files_str}"
                    messagebox.showinfo(title, msg)

        except RuntimeError as e:
            print(f"\n[ERROR] {e}")
            messagebox.showerror("Error", str(e))
        except Exception as e:
            print(f"\n[ERROR] Terjadi kesalahan: {e}")
            messagebox.showerror("Error", f"Terjadi kesalahan saat scraping:\n{e}")
        finally:
            if driver:
                print("[*] Menutup browser...")
                try:
                    driver.quit()
                except Exception:
                    pass
            self.root.after(0, lambda: self.btn_run.config(state="normal"))
            self.root.after(0, lambda: self.btn_stop.config(state="disabled"))
            print("\n[=] Script siap digunakan kembali.")

    # ── CLOSE ─────────────────────────────────────────────────────────────────

    def _on_close(self):
        sys.stdout = self.original_stdout
        self.root.destroy()

# ═════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app  = ScraperApp(root)
    root.mainloop()