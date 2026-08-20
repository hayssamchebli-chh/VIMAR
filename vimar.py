"""Vimar datasheet lookup and download.

vimar.com sits behind an Imperva/Incapsula WAF. The document repository
(/irj/go/km/docs/...) serves PDFs freely, but the product pages that link
them are challenged for automated traffic, so the downloader works through
four layers, cheapest first:

1. a PDF already saved on disk (manual folder or Downloads), always wins
2. a real Chrome window where the person passed Vimar's human check once
3. the site's own download-pdf endpoint, through a warmed-up session
4. the product page, scraped for links into the document repository

Layers 3-4 are refused for automated traffic in practice, so layer 2 is the
one that makes unattended downloading work.
"""

import base64
import io
import os
import re
import threading
import time
from html import unescape
from urllib.parse import quote

import requests
from pypdf import PdfReader

BASE_URL = "https://www.vimar.com"
HOME_URL = BASE_URL + "/en/int"
CATALOG_SECTIONS = ("product", "obsolete")  # current catalogue, then discontinued
PRODUCT_URL = BASE_URL + "/en/int/catalog/{section}/index/code/{code}"
DOWNLOAD_URL = BASE_URL + "/en/int/catalog/{section}/download-pdf/code/{code}?type=.pdf"
DOC_REPOSITORY = "/irj/go/km/docs/"

MANUAL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "manual_datasheets")

PAGE_TIMEOUT = (10, 30)
PDF_TIMEOUT = (10, 90)
ATTEMPTS = 2

BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

try:
    from playwright.sync_api import sync_playwright as _sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:  # pragma: no cover - optional dependency
    _sync_playwright = None
    PLAYWRIGHT_AVAILABLE = False


# ------------------------------------------------------------------
# Codes
# ------------------------------------------------------------------

def normalize_code(value) -> str:
    """Clean one Vimar code and drop the VIMA prefix.

    Vimar codes keep their dots and letters: 19755.2, 20755.3.B, 21860.
    """
    if value is None:
        return ""

    code = re.sub(r"\s+", "", str(value)).strip()
    code = re.sub(r"[^A-Za-z0-9.\-_/]", "", code)

    for prefix in ("VIMAR", "VIMA", "VIM"):
        if code.upper().startswith(prefix):
            code = code[len(prefix):]
            break

    return code.strip("-_. ").upper()


def extract_codes_from_text(text: str) -> list[str]:
    """Read codes from the paste box: one per line, commas also accepted."""
    if not text:
        return []

    codes = []
    for chunk in re.split(r"[\n,;\t]+", text):
        code = normalize_code(chunk)
        if code:
            codes.append(code)

    return codes


def dedupe(items: list[str]) -> list[str]:
    """Drop repeats while keeping the original order."""
    seen = set()
    result = []
    for item in items:
        if item.casefold() not in seen:
            seen.add(item.casefold())
            result.append(item)
    return result


# ------------------------------------------------------------------
# PDF helpers
# ------------------------------------------------------------------

def is_pdf(content: bytes) -> bool:
    return bool(content) and content[:5] == b"%PDF-"


def validate_pdf(content: bytes) -> None:
    """Raise unless the bytes are a PDF that actually parses."""
    if not content:
        raise ValueError("empty response")
    if not is_pdf(content):
        raise ValueError("response is not a PDF")
    PdfReader(io.BytesIO(content))


def looks_blocked(content: bytes) -> bool:
    """Detect the WAF interstitial returned instead of a document."""
    head = content[:1500].lower()
    return b"incapsula" in head or b"_incapsula_resource" in head


# ------------------------------------------------------------------
# Download layers
# ------------------------------------------------------------------

def _key(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def pdf_mentions_code(content: bytes, code: str) -> bool:
    """True when the code appears in the PDF's own text (first pages)."""
    try:
        reader = PdfReader(io.BytesIO(content))
        text = " ".join((page.extract_text() or "") for page in reader.pages[:2])
    except Exception:
        return False

    return _key(code) in _key(text)


def find_saved_datasheet(code: str, folders: list[str] | None = None) -> tuple[bytes, str]:
    """Find a PDF the user already downloaded, in any of the given folders.

    Two ways to match, so the file does not have to be renamed:
    - the filename contains the code (20582.pdf, VIMA-20755.3.B.pdf, 20755_3_B.pdf)
    - the PDF's own text contains the code, whatever the file is called

    Returns (content, path).
    """
    folders = [f for f in (folders or [MANUAL_DIR]) if f and os.path.isdir(f)]
    wanted = _key(code)

    by_name: list[str] = []
    others: list[str] = []

    for folder in folders:
        try:
            names = os.listdir(folder)
        except Exception:
            continue

        for filename in names:
            if not filename.lower().endswith(".pdf"):
                continue

            path = os.path.join(folder, filename)
            stem = _key(os.path.splitext(filename)[0])

            if stem in (wanted, "vima" + wanted, "vimar" + wanted) or wanted in stem:
                by_name.append(path)
            else:
                others.append(path)

    for path in by_name:
        try:
            with open(path, "rb") as f:
                content = f.read()
            if is_pdf(content):
                return content, path
        except Exception:
            continue

    # Nothing matched by name: check the contents of recent downloads instead.
    others.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    for path in others[:40]:
        try:
            with open(path, "rb") as f:
                content = f.read()
        except Exception:
            continue

        if is_pdf(content) and pdf_mentions_code(content, code):
            return content, path

    return b"", ""


def new_session() -> requests.Session:
    """A session warmed up on the homepage, so it carries the site cookies."""
    session = requests.Session()
    session.headers.update(BROWSER_HEADERS)
    try:
        session.get(HOME_URL, timeout=PAGE_TIMEOUT)
    except Exception:
        pass
    return session


def fetch_pdf(session: requests.Session, url: str, referer: str = "") -> tuple[bytes, str]:
    """Download a PDF, retrying once. Returns (content, error)."""
    headers = {"Accept": "application/pdf,*/*"}
    if referer:
        headers["Referer"] = referer

    last_error = "request failed"

    for attempt in range(1, ATTEMPTS + 1):
        try:
            response = session.get(url, timeout=PDF_TIMEOUT, headers=headers)

            if response.status_code != 200:
                last_error = f"HTTP {response.status_code}"
                if response.status_code == 404:
                    return b"", last_error
                continue

            content = response.content or b""

            if looks_blocked(content):
                return b"", "blocked by the Vimar firewall"
            if not is_pdf(content):
                return b"", "response was not a PDF"

            return content, ""

        except Exception as e:
            last_error = str(e)
            if attempt < ATTEMPTS:
                time.sleep(attempt)

    return b"", last_error


def document_links(html: str) -> list[str]:
    """Datasheet links on a product page, pointing into the doc repository."""
    links = []

    for href in re.findall(r'href="([^"]+)"', html):
        href = unescape(href)
        low = href.lower()
        if DOC_REPOSITORY in low or low.endswith(".pdf") or "download-pdf" in low:
            if href.startswith("/"):
                href = BASE_URL + href
            if href.startswith("http") and href not in links:
                links.append(href)

    # A data sheet is preferred over instruction sheets when both are offered.
    def rank(url: str) -> int:
        low = url.lower()
        if "download-pdf" in low:
            return 0
        if re.search(r"(scheda|data|tech|_st|_ft)", low):
            return 1
        return 2

    return sorted(links, key=rank)


def download_via_site(code: str, session: requests.Session) -> tuple[bytes, str, str]:
    """Layers 2-3: the download endpoint, then the product page's own links."""
    safe = quote(code, safe=".-_")
    product_url = PRODUCT_URL.format(section="product", code=safe)
    endpoint_error = ""

    for section in CATALOG_SECTIONS:
        url = DOWNLOAD_URL.format(section=section, code=safe)
        content, error = fetch_pdf(
            session, url, referer=PRODUCT_URL.format(section=section, code=safe)
        )
        if content:
            return content, url, ""
        endpoint_error = error

    try:
        page = session.get(product_url, timeout=PAGE_TIMEOUT)
        html = page.text if page.status_code == 200 else ""
    except Exception as e:
        return b"", "", f"{endpoint_error}; product page: {e}"

    if not html or looks_blocked(html.encode("utf-8", "ignore")):
        return b"", "", f"{endpoint_error}; product page blocked by the Vimar firewall"

    for link in document_links(html)[:4]:
        content, error = fetch_pdf(session, link, referer=product_url)
        if content:
            return content, link, ""

    return b"", "", f"{endpoint_error}; no downloadable datasheet on the product page"


def download_via_browser(code: str) -> tuple[bytes, str, str]:
    """Layer 4: repeat the flow inside a real browser, if Playwright is here."""
    if not PLAYWRIGHT_AVAILABLE or _sync_playwright is None:
        return b"", "", "Playwright is not installed"

    product_url = PRODUCT_URL.format(code=code)

    try:
        with _sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=[
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-blink-features=AutomationControlled",
                ],
            )
            context = browser.new_context(
                user_agent=BROWSER_HEADERS["User-Agent"],
                locale="en-US",
                viewport={"width": 1366, "height": 850},
                accept_downloads=True,
            )
            context.add_init_script(
                "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"
            )
            page = context.new_page()

            try:
                page.goto(HOME_URL, wait_until="domcontentloaded", timeout=45_000)
                page.wait_for_timeout(2_500)
                page.goto(product_url, wait_until="domcontentloaded", timeout=45_000)
                page.wait_for_timeout(2_500)

                html = page.content()
                if looks_blocked(html.encode("utf-8", "ignore")):
                    return b"", "", "product page blocked by the Vimar firewall"

                candidates = document_links(html)
                candidates.insert(0, DOWNLOAD_URL.format(code=code))

                for link in candidates[:5]:
                    response = context.request.get(
                        link,
                        headers={"Accept": "application/pdf,*/*", "Referer": product_url},
                        timeout=60_000,
                    )
                    if not response.ok:
                        continue

                    content = response.body()
                    if content and is_pdf(content) and not looks_blocked(content):
                        return content, link, ""

                return b"", "", "no downloadable datasheet found in the browser"

            finally:
                browser.close()

    except Exception as e:
        return b"", "", str(e)


def product_page_url(code: str, section: str = "product") -> str:
    """The page a person opens to save the datasheet by hand."""
    return PRODUCT_URL.format(section=section, code=normalize_code(code))


def download_datasheet(code: str, session: requests.Session | None = None,
                       use_browser: bool = True,
                       folders: list[str] | None = None) -> dict:
    """Fetch one Vimar datasheet, trying every layer in turn."""
    code = normalize_code(code)

    result = {
        "code": code,
        "success": False,
        "url": "",
        "source": "",
        "error": "",
        "content": None,
    }

    if not code:
        result["error"] = "empty code"
        return result

    saved, path = find_saved_datasheet(code, folders)
    if saved:
        try:
            validate_pdf(saved)
            result.update(success=True, content=saved, source="saved file", url=path)
            return result
        except Exception as e:
            result["error"] = f"saved file is not a valid PDF: {e}"

    source = "vimar.com"

    # The verified browser window, when the human check has been passed there.
    content, url, error = b"", "", ""
    if browser_port_open():
        content, url, error = download_via_open_browser(code)
        if content:
            source = "verified browser"

    if not content:
        site_content, site_url, site_error = download_via_site(code, session or new_session())
        if site_content:
            content, url = site_content, site_url
        else:
            error = f"{error + ' | ' if error else ''}{site_error}"

    if not content and use_browser:
        browser_content, browser_url, browser_error = download_via_browser(code)
        if browser_content:
            content, url, source = browser_content, browser_url, "browser"
        else:
            error = f"{error} | browser: {browser_error}"

    if not content:
        result["error"] = (
            f"vimar.com refused the download ({error}). "
            f"Open the Vimar browser window, pass the human check, then run again - "
            f"or save the data sheet yourself and the app will pick it up."
        )
        result["url"] = product_page_url(code)
        return result

    try:
        validate_pdf(content)
    except Exception as e:
        result["error"] = f"downloaded file is not a valid PDF: {e}"
        result["url"] = url
        return result

    result.update(success=True, content=content, url=url, source=source)
    return result


# ------------------------------------------------------------------
# Verified browser session
#
# Vimar shows a human check before letting a browser into the catalogue.
# A person passes it once in a real Chrome window; the clearance stays in
# that window's profile, and the app then downloads through that same
# session instead of fighting the check.
# ------------------------------------------------------------------

DEBUG_PORT = 9222

# One browser window, one page: downloads through it must not overlap or the
# workers navigate each other's page mid-fetch.
_BROWSER_LOCK = threading.Lock()
CHROME_PROFILE = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "vimar-datasheets-chrome"
)

CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
]


def find_browser() -> str:
    """Path to an installed Chrome or Edge."""
    for path in CHROME_CANDIDATES:
        if os.path.exists(path):
            return path
    return ""


def browser_port_open(port: int = DEBUG_PORT) -> bool:
    """True when a verification window is already running."""
    import socket

    try:
        with socket.create_connection(("127.0.0.1", port), timeout=1):
            return True
    except OSError:
        return False


def open_verification_browser(code: str = "21860", port: int = DEBUG_PORT) -> tuple[bool, str]:
    """Open the Vimar site in a real browser window for the human check.

    Reuses the window if one is already open. Returns (started, message).
    """
    if browser_port_open(port):
        return True, "The Vimar browser window is already open."

    browser = find_browser()
    if not browser:
        return False, "Could not find Chrome or Edge on this computer."

    os.makedirs(CHROME_PROFILE, exist_ok=True)

    try:
        import subprocess

        # Detached, so the window outlives whatever started it - otherwise the
        # verified session dies with the Streamlit process or a test script.
        creation_flags = 0
        if os.name == "nt":
            creation_flags = (
                getattr(subprocess, "DETACHED_PROCESS", 0)
                | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            )

        subprocess.Popen(
            [
                browser,
                f"--remote-debugging-port={port}",
                f"--user-data-dir={CHROME_PROFILE}",
                "--no-first-run",
                "--no-default-browser-check",
                product_page_url(code),
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creation_flags,
            close_fds=True,
        )
    except Exception as e:
        return False, f"Could not start the browser: {e}"

    for _ in range(40):
        if browser_port_open(port):
            return True, "Browser window opened - complete the human check in it."
        time.sleep(0.5)

    return False, "The browser started but its debugging port never opened."


def _attached_page(playwright, port: int):
    """Attach to the open window and return (browser, context, page)."""
    browser = playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{port}", timeout=15_000)
    context = browser.contexts[0] if browser.contexts else browser.new_context()
    page = context.pages[0] if context.pages else context.new_page()
    return browser, context, page


def verification_status(code: str = "21860", port: int = DEBUG_PORT) -> tuple[bool, str]:
    """Check whether the open window can now reach the Vimar catalogue."""
    if not PLAYWRIGHT_AVAILABLE or _sync_playwright is None:
        return False, "Playwright is not installed."

    if not browser_port_open(port):
        return False, "No Vimar browser window is open yet."

    try:
        with _BROWSER_LOCK, _sync_playwright() as p:
            browser, _, page = _attached_page(p, port)
            try:
                _open_verified(page, product_page_url(code, "product"))
                html = page.content()
                body_length = page.evaluate(
                    "document.body ? document.body.innerText.length : 0"
                )
            finally:
                pass  # keep the verification window open (see above)
    except Exception as e:
        return False, f"Could not talk to the browser window: {e}"

    if "Incapsula incident" in html:
        return False, ("The human check is still showing - complete it in the browser "
                       "window, then check again.")

    if body_length < 300 or "not available" in html:
        return False, ("The window opened but the catalogue did not load - "
                       "complete the human check in it, then check again.")

    return True, "Verified: the catalogue is reachable from the browser window."


def _open_verified(page, url: str, attempts: int = 4) -> tuple[bool, str]:
    """Load a catalogue page, giving the window time to clear a fresh check.

    Vimar re-checks a session now and then. In a real browser the check
    usually clears itself a moment later, so the page is reloaded a couple of
    times before giving up and asking the person to look at the window.
    """
    for attempt in range(1, attempts + 1):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=45_000)
        except Exception as e:
            if attempt == attempts:
                return False, str(e)
            time.sleep(2)
            continue

        page.wait_for_timeout(2_000)

        if "Incapsula incident" not in page.content():
            return True, ""

        if attempt < attempts:
            time.sleep(3)

    return False, "the human check is showing in the browser window"


def download_via_open_browser(code: str, port: int = DEBUG_PORT) -> tuple[bytes, str, str]:
    """Download one datasheet through the verified browser window.

    The PDF is fetched by the page itself rather than through Playwright's
    request API: only the page carries the full verified browser context, and
    Vimar answers anything else with 403. Discontinued items are not in the
    main catalogue, so the obsolete one is tried next.
    """
    if not PLAYWRIGHT_AVAILABLE or _sync_playwright is None:
        return b"", "", "Playwright is not installed"

    if not browser_port_open(port):
        return b"", "", "no verified browser window is open"

    fetch_in_page = """async (url) => {
        const response = await fetch(url, {credentials: 'include'});
        if (!response.ok) return {status: response.status, data: ''};
        const bytes = new Uint8Array(await response.arrayBuffer());
        let binary = '';
        const chunk = 0x8000;
        for (let i = 0; i < bytes.length; i += chunk) {
            binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
        }
        return {status: response.status, data: btoa(binary)};
    }"""

    last_error = "product code not found in either catalogue"

    try:
        with _BROWSER_LOCK, _sync_playwright() as p:
            browser, _, page = _attached_page(p, port)
            try:
                for section in CATALOG_SECTIONS:
                    page_url = product_page_url(code, section)

                    opened, why = _open_verified(page, page_url)
                    if not opened:
                        return b"", "", why

                    text = page.evaluate("document.body ? document.body.innerText : ''")
                    if "not available" in text or "not found" in page.title().lower():
                        continue

                    url = DOWNLOAD_URL.format(section=section, code=code)
                    result = page.evaluate(fetch_in_page, url)

                    content = b""
                    if result.get("data"):
                        content = base64.b64decode(result["data"])

                    if content and is_pdf(content) and not looks_blocked(content):
                        return content, url, ""

                    last_error = f"HTTP {result.get('status')} from the {section} catalogue"

                return b"", "", last_error
            finally:
                # Do NOT call browser.close(): on a CDP attachment that shuts
                # the person's verification window. Leaving the playwright
                # context disconnects our side and leaves Chrome running.
                pass

    except Exception as e:
        return b"", "", str(e)
