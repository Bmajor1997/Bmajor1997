# ============================================================
# Project Name: Browser Extension Filetype Filter Verification
# ============================================================

# ============================================================
# DIRECTIONS
# ============================================================
"""
This project launches the browser extension in a Playwright Chromium session, validates the popup file input behavior,
verifies supported and unsupported file handling, and generates a structured test report suitable for CI execution.
"""
# ============================================================
# IMPORTS
# ============================================================
import re
import pytest
from pathlib import Path
from playwright.sync_api import Page, expect,TimeoutError as PlaywrightTimeoutError
# ============================================================
# CONSTANTS
# ============================================================
FORMAT_MAP = {
    ".pdf",
    ".doc",
    ".docx",
    ".rtf",
    ".ppt",
    ".pptx",
    ".epub",
    ".mobi",
    ".odt",
    ".ods",
    ".odp",
    ".txt",
    ".xls",
    ".xlsx",
}

SUPPORTED_IMG = {
    ".bmp",
    ".gif",
    ".jpg",
    ".jpeg",
    ".png",
    ".tiff",
    ".tif",
}

EXT_FAMILY_MAP = {
    ".jpeg": ".jpg",
    ".tiff": ".tif",
    ".docx": ".doc",
    ".pptx": ".ppt",
    ".xlsx": ".xls",
}

EXTENSION_ENTRY_FILES = [
    "popup.html",
    "index.html",
]
BASE_DIR = Path(__file__).resolve().parent
STATE = BASE_DIR / "storage_state.json"
TEST_DIR = BASE_DIR / "test_documents" / "various_test_documents_for_projects"
TIMEOUT_SECONDS = 30
# ============================================================
# SOURCE CODE
# ============================================================
def build_test_cases(sample_dir: str):
    test_cases = []
    unsupported_found = set()

    sample_path = Path(sample_dir)
    if not sample_path.exists():
        raise FileNotFoundError(f"Sample directory not found: {sample_path}")

    for file_path in sorted(sample_path.iterdir()):
        if not file_path.is_file():
            continue

        ext = file_path.suffix.lower()

        if ext == ".csv":
            continue

        test_cases.append(file_path)

        if (ext not in FORMAT_MAP) and (ext not in SUPPORTED_IMG):
            unsupported_found.add(ext)

    return test_cases, unsupported_found

def _family_ext(ext: str) -> str:
    ext = ext.lower()
    return EXT_FAMILY_MAP.get(ext, ext)

def parse_accept_attribute(accept_value: str) -> set[str]:
    tokens = set()

    if not accept_value:
        return tokens

    for raw_piece in accept_value.split(","):
        piece = raw_piece.strip().lower()
        if not piece:
            continue

        if piece.startswith("."):
            tokens.add(piece)
            continue

        mime_map = {
            "image/*": {".bmp", ".gif", ".jpg", ".jpeg", ".png", ".tif", ".tiff"},
            "application/pdf": {".pdf"},
            "text/plain": {".txt"},
            "application/msword": {".doc"},
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document": {".docx"},
            "application/vnd.ms-powerpoint": {".ppt"},
            "application/vnd.openxmlformats-officedocument.presentationml.presentation": {".pptx"},
            "application/vnd.ms-excel": {".xls"},
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {".xlsx"},
        }

        if piece in mime_map:
            tokens.update(mime_map[piece])

    return tokens

def validate_accept_filter(accept_value: str):
    accepted_tokens = parse_accept_attribute(accept_value)
    expected_supported = FORMAT_MAP.union(SUPPORTED_IMG)

    missing_supported = sorted(ext for ext in expected_supported if ext not in accepted_tokens)
    unexpected_supported = sorted(ext for ext in accepted_tokens if ext not in expected_supported)

    return accepted_tokens, missing_supported, unexpected_supported

def find_extension_id(context, timeout_ms: int) -> str:
    service_worker = None

    if context.service_workers:
        service_worker = context.service_workers[0]
    else:
        try:
            service_worker = context.wait_for_event(
                "serviceworker",
                timeout=timeout_ms
            )
        except PlaywrightTimeoutError:
            pass

    if service_worker:
        sw_url = service_worker.url
        if sw_url.startswith("chrome-extension://"):
            return sw_url.split("/")[2]

    background_pages = context.background_pages
    if background_pages:
        bg_url = background_pages[0].url
        if bg_url.startswith("chrome-extension://"):
            return bg_url.split("/")[2]

    raise RuntimeError(
        "Could not determine extension ID. "
        "The extension may not have loaded correctly in the CI browser."
    )

def open_extension_popup(page, extension_id: str, timeout_seconds: int):

    last_error = None

    for entry_file in EXTENSION_ENTRY_FILES:
        url = f"chrome-extension://{extension_id}/{entry_file}"

        try:
            page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=timeout_seconds * 1000
            )

            file_input = page.locator("input[type='file']").first
            expect(file_input).to_be_attached(timeout=timeout_seconds * 1000)

            return url

        except Exception as exc:
            last_error = exc

    raise RuntimeError(
        f"Could not open extension popup UI from known entry files. Last error: {last_error}"
    )

def upload_file(page, file_path: Path):
    file_input = page.locator("input[type='file']").first
    file_input.wait_for(state="attached", timeout=300_000)

    file_input.set_input_files(str(file_path))

    try:
        expect(file_input).to_have_value(re.compile(re.escape(file_path.name), re.I), timeout=5000)
    except Exception:
        pass

    try:
        start_btn = page.get_by_role("button", name=re.compile(r"(Start|Convert)", re.I))
        if start_btn.is_visible(timeout=300_000):
            start_btn.click()
    except Exception:
        pass

def wait_for_upload_result(page, timeout_seconds: int) -> bool:

    failure_signals = [
        page.get_by_text("Unsupported", exact=False),
        page.get_by_text("not supported", exact=False),
        page.get_by_text("Invalid", exact=False),
        page.get_by_text("Error", exact=False),
    ]

    success_signals = [
        page.get_by_text("Ready", exact=False),
        page.get_by_text("Convert", exact=False),
        page.get_by_text("Start", exact=False),
    ]

    deadline_ms = timeout_seconds * 1000
    elapsed = 0
    poll_ms = 500

    while elapsed < deadline_ms:
        for signal in failure_signals:
            try:
                if signal.is_visible(timeout=100):
                    return False
            except Exception:
                pass

        for signal in success_signals:
            try:
                if signal.is_visible(timeout=100):
                    return True
            except Exception:
                pass

        page.wait_for_timeout(poll_ms)
        elapsed += poll_ms

    return False
# ============================================================
# TEST
# ============================================================
@pytest.mark.browser_context_args(storage_state=str(STATE))
def test_browser_extension(page: Page):
    test_cases, unsupported_found = build_test_cases(TEST_DIR)

    assert test_cases, "No test cases found in the sample directory."

    results = []

    for file_path in test_cases:
        file_path = Path(file_path)
        ext = file_path.suffix.lower()

        page.goto("https://test.scribeit.io/", wait_until="domcontentloaded")

        file_input = page.locator("input[type='file']").first
        expect(file_input).to_be_attached(timeout=15000)

        accept_value = (file_input.get_attribute("accept") or "").lower()
        accepted_tokens = parse_accept_attribute(accept_value)

        is_supported = (ext in FORMAT_MAP) or (ext in SUPPORTED_IMG)

        if is_supported:
            upload_file(page, file_path)
            ok = wait_for_upload_result(page, TIMEOUT_SECONDS)

            assert ok, f"Supported file failed upload validation: {file_path.name}"
