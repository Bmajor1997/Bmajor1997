# ==========================================================
# Project Name: test_document_verification.py
# ==========================================================

# ==========================================================
# DIRECTIONS
# ==========================================================
"""
Upload documents to ScribeIt, convert them into all required
output formats, download the results when available,
and verify that each conversion completes successfully,
failing the test if any required conversion does not
succeed.
"""
# ==========================================================
# IMPORTS
# ==========================================================
from playwright.sync_api import Page
from pathlib import Path
import re
import pytest
# ==========================================================
# CONSTANTS
# ==========================================================
WEBSITE = "https://test.scribeit.io/"
BASE_DIR = Path(__file__).resolve().parent
STATE = BASE_DIR / "storage_state.json"
TEST_DIR = BASE_DIR / "test_documents" / "various_test_documents_for_projects"
OUT_DIR = BASE_DIR / "downloads"
SCREENSHOT_DIR = BASE_DIR / "ci_debug_screenshots"
SUPPORTED_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".epub", ".mobi",
    ".rtf", ".ppt", ".pptx", ".txt", ".gif", ".png",
    ".odt", ".ods", ".odp", ".tif", ".jpeg", ".bmp"
}

OUTPUT_FORMATS = [
    "PDF",
    "Downloadable web page",
    "EPUB",
    "DAISY",
    "MOBI",
    "Audio",
    "Braille",
    "Read full version in browser",
    "Microsoft Word",
]

PAGE_TIMEOUT = 60_000
SHORT_TIMEOUT = 5_000
MEDIUM_TIMEOUT = 30_000
POST_START_TIMEOUT_MS = 750_000
FORMAT_SCREEN_TIMEOUT_MS = 60_000
EXTENSION_TO_FORMAT = {
    ".pdf": "PDF",
    ".doc": "Microsoft Word",
    ".docx": "Microsoft Word",
    ".epub": "EPUB",
    ".mobi": "MOBI",
    ".rtf": "Microsoft Word",
    ".ppt": "Microsoft Word",
    ".pptx": "Microsoft Word",
    ".txt": "Microsoft Word",
    ".odt": "Microsoft Word",
    ".ods": "Microsoft Word",
    ".odp": "Microsoft Word",
    ".bmp": "PDF",
    ".png": "PDF",
    ".gif": "PDF",
    ".jpeg": "PDF",
    ".tif": "PDF",
}
# ==========================================================
# SOURCE CODE
# ==========================================================
def get_test_document_path():
    file_path = TEST_DIR / "Appendix 3_Water Quality Tables_0007.pdf"

    if not file_path.exists():
        raise Exception(f"Test document not found: {file_path}")

    return file_path

def sanitize_name(text):
    sanitized = re.sub(r'[<>:"/\\|?*]+', "_", text)
    sanitized = re.sub(r"\s+", "_", sanitized).strip("_")
    return sanitized[:150] if sanitized else "debug_file"

def click(page, role=None, name=None, selector=None, text=None, timeout=SHORT_TIMEOUT):
    if selector:
        page.locator(selector).first.click(timeout=timeout)
        return

    if role:
        page.get_by_role(role, name=name).first.click(timeout=timeout)
        return

    if text:
        page.get_by_text(text, exact=False).first.click(timeout=timeout)
        return

    raise ValueError("No click strategy was provided.")

def upload_document(page, file_path):
    file_input = page.locator("input[type='file']").first
    file_input.wait_for(state="attached", timeout=MEDIUM_TIMEOUT)
    file_input.set_input_files(str(file_path))

def save_debug_screenshot(page, file_name, suffix):
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
    screenshot_path = SCREENSHOT_DIR / f"{sanitize_name(file_name)}_{suffix}.png"
    page.screenshot(path=str(screenshot_path), full_page=True)
    return screenshot_path

def click_start_and_select_format(page, file_name):
    start_btn = page.get_by_role("button", name="Start")
    start_btn.wait_for(state="visible", timeout=MEDIUM_TIMEOUT)
    start_btn.click()

    elapsed_ms = 0
    poll_ms = 2_000

    while elapsed_ms < POST_START_TIMEOUT_MS:
        try:
            remediate_btn = page.get_by_role("button", name="Remediate Full Document").first
            if remediate_btn.count() > 0 and remediate_btn.is_visible():
                raise Exception("Scribe requires remediation; cannot proceed to Select Format.")
        except Exception as error:
            if "Scribe requires remediation" in str(error):
                raise

        try:
            select_format = page.get_by_role("button", name="Select Format").first
            if select_format.count() > 0 and select_format.is_visible():
                if not select_format.is_enabled():
                    raise Exception("'Select Format' is visible but disabled.")

                select_format.click(timeout=POST_START_TIMEOUT_MS)
                return
        except Exception as error:
            if "Select Format" in str(error):
                raise

        page.wait_for_timeout(poll_ms)
        elapsed_ms += poll_ms

    raise Exception("Timed out waiting for Select Format after Start.")

def wait_for_format_screen(page, file_name):
    format_markers = [
        "Audio", "Braille", "Downloadable web page", "EPUB",
        "Microsoft Word", "DAISY", "MOBI", "Skip Settings and Convert",
        "Read full version in browser", "PDF",
    ]

    elapsed_ms = 0
    poll_ms = 1_000

    page.wait_for_load_state("domcontentloaded", timeout=PAGE_TIMEOUT)

    while elapsed_ms < FORMAT_SCREEN_TIMEOUT_MS:
        for marker in format_markers:
            try:
                if page.get_by_text(marker, exact=False).first.is_visible(timeout=500):
                    return
            except Exception:
                continue

        page.wait_for_timeout(poll_ms)
        elapsed_ms += poll_ms

    screenshot_path = save_debug_screenshot(page, file_name, "format_screen_missing")
    raise Exception(f"Format selection screen did not appear. Screenshot: {screenshot_path}")

def selecting_format(page, target_format, current_format):
    if target_format.lower().strip() == (current_format or "").lower().strip():
        raise Exception(f"Target format '{target_format}' is already the current format.")

    options = page.locator("[role='option'], [role='menuitem'], li, button")

    for i in range(options.count()):
        try:
            text = options.nth(i).inner_text(timeout=2000).strip()
            if text.lower() == target_format.lower():
                options.nth(i).click(timeout=MEDIUM_TIMEOUT)
                return text
        except:
            continue

    raise Exception(f"Target format '{target_format}' not found.")

def click_convert(page):
    for pattern in [r"skip settings and convert", r"^convert$", r"^go$", r"^start$"]:
        try:
            click(page, role="button", name=re.compile(pattern, re.I))
            return
        except:
            continue
    raise Exception("Could not find Convert button.")


def click_download_if_present(page, file_name, target_format):
    for role in ("button", "link"):
        try:
            with page.expect_download(timeout=MEDIUM_TIMEOUT) as download_info:
                click(page, role=role, name=re.compile(r"\bdownload\b", re.I))
            download = download_info.value
            download.save_as(str(OUT_DIR / download.suggested_filename))
            return True
        except:
            continue

    return False

def validate_required_files():
    if not STATE.exists():
        raise FileNotFoundError("Missing storage_state.json")

def reset_to_home(page):
    page.goto(WEBSITE)
# ==========================================================
# TEST FUNCTION
# ==========================================================
@pytest.mark.browser_context_args(storage_state=str(STATE))
def test_document_verification(page: Page):
    validate_required_files()

    file_path = get_test_document_path()
    file_name = file_path.name
    current_format = EXTENSION_TO_FORMAT.get(file_path.suffix.lower())

    print(f"Using file: {file_name}")

    passed = []
    failed = []

    for target_format in OUTPUT_FORMATS:
        try:
            if target_format == current_format:
                print(f"Skipping {file_name} because it is already in {target_format}")
                continue

            reset_to_home(page)

            upload_document(page, file_path)
            click_start_and_select_format(page, file_name)
            wait_for_format_screen(page, file_name)

            selecting_format(page, target_format, current_format)
            click_convert(page)
            click_download_if_present(page, file_name, target_format)

            print(f"✅ {file_name} -> {target_format}")
            passed.append(f"{file_name} -> {target_format}")

        except Exception as error:
            print(f"❌ {file_name} -> {target_format}: {error}")
            failed.append(f"{file_name} -> {target_format}: {error}")

    assert not failed, f"Some conversions failed: {failed}"

    print(f"Passed conversions: {passed}")
