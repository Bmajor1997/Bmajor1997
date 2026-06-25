# ==================================================
# Project Name: Focus Change Events
# ==================================================

# ==================================================
# DIRECTIONS
# ==================================================
"""
This project will automate validation of logical focus placement in the UI
whenever programmatic focus changes are expected to happen during the document
workflows.
"""
# ============================================================
# IMPORTS
# ============================================================
import re
import time
from pathlib import Path
import pytest
from playwright.sync_api import Page, expect
# ============================================================
# CONSTANTS
# ============================================================
FORMAT_MAP = {
    ".pdf", ".doc", ".docx", ".rtf", ".ppt", ".pptx",
    ".epub", ".mobi", ".odt", ".ods", ".odp", ".txt",
    ".xls", ".xlsx",
}

SUPPORTED_IMG = {
    ".bmp", ".gif", ".jpg", ".jpeg", ".png", ".tiff", ".tif",
}

EXT_FAMILY_MAP = {
    ".jpeg": ".jpg",
    ".tiff": ".tif",
    ".docx": ".doc",
    ".pptx": ".ppt",
    ".xlsx": ".xls",
}
BASE_DIR = Path(__file__).resolve().parent
STATE = BASE_DIR / "storage_state.json"
HOME_URL = "https://test.scribeit.io/"
SAMPLE_DIR = BASE_DIR / "test_documents" / "various_test_documents_for_projects"
TIMEOUT_SECONDS = 60
# ============================================================
# TEST CASE DISCOVERY
# ============================================================
def validate_sample_directory(sample_path: Path):
    if not sample_path.exists() or not sample_path.is_dir():
        raise FileNotFoundError("Sample directory does not exist.")

def is_supported_file(file_path: Path):
    file_extension = file_path.suffix.lower()
    return file_extension in FORMAT_MAP or file_extension in SUPPORTED_IMG

def collect_supported_test_files(sample_path: Path):
    test_cases = []
    unsupported_extensions = set()

    for file_path in sorted(sample_path.iterdir()):
        if not file_path.is_file():
            continue

        file_extension = file_path.suffix.lower()

        if is_supported_file(file_path):
            test_cases.append(file_path)
        else:
            unsupported_extensions.add(file_extension)

    return test_cases, unsupported_extensions

def build_test_cases(sample_dir: str):
    sample_path = Path(sample_dir)

    validate_sample_directory(sample_path)

    test_cases, unsupported_extensions = collect_supported_test_files(sample_path)

    if not test_cases:
        raise RuntimeError("No supported test files were found.")

    return test_cases, unsupported_extensions

def get_default_sample_dir() -> str:
    paths = [
        Path.home() / "test_documents" / "various_test_documents_for_projects",
        Path(__file__).resolve().parent / "test_documents",
    ]

    for path in paths:
        if path.is_dir():
            return path

    raise FileNotFoundError("Sample directory not found.")

def _family_ext(ext: str) -> str:
    ext = ext.lower()
    return EXT_FAMILY_MAP.get(ext, ext)
# ============================================================
# UPLOAD / RESET FLOW
# ============================================================
def reset_upload_state(page: Page, timeout_seconds: int):
    page.goto(HOME_URL, wait_until="networkidle")
    page.locator("input[type='file']").first.wait_for(
        state="attached",
        timeout=timeout_seconds * 1000,
    )

def upload_file(page: Page, file_path: Path):
    file_path = Path(file_path)

    try:
        page.get_by_text("Convert another document", exact=False).click(timeout=1500)
    except Exception:
        pass

    file_input = page.locator("input[type='file']").first
    file_input.wait_for(state="attached", timeout=15000)

    file_input.set_input_files(str(file_path))

    start_button = page.get_by_role(
        "button",
        name=re.compile(r"(Start|Convert)", re.I),
    )

    start_button.wait_for(state="visible", timeout=15000)
    start_button.click()

def wait_for_upload_result(page: Page, timeout_seconds: int) -> bool:
    success_signals = [
        page.get_by_text("Convert another document", exact=False),
        page.get_by_role("link", name="Document", exact=False),
        page.get_by_role("link", name="Audio", exact=False),
    ]

    failure_signals = [
        page.get_by_text("Error", exact=False),
        page.get_by_text("Something went wrong", exact=False),
        page.get_by_text("Try again", exact=False),
        page.get_by_text("Failed", exact=False),
    ]

    start_time = time.monotonic()

    while time.monotonic() - start_time < timeout_seconds:
        for signal in success_signals:
            try:
                if signal.is_visible(timeout=500):
                    return True
            except Exception:
                pass

        for signal in failure_signals:
            try:
                if signal.is_visible(timeout=500):
                    return False
            except Exception:
                pass

        page.wait_for_timeout(250)

    raise TimeoutError("Timed out waiting for upload result.")

def upload_test_file(page, test_file_path):
    test_file_path = Path(test_file_path)

    assert test_file_path.exists(), f"UPLOAD: test file does not exist: {test_file_path}"

    convert_another = page.get_by_text("Convert another document", exact=False)

    try:
        if convert_another.is_visible(timeout=1500):
            convert_another.click()
    except Exception:
        pass

    file_input = page.locator("input[type='file']").first

    try:
        file_input.wait_for(state="attached", timeout=15000)
    except Exception as e:
        pytest.fail(f"UPLOAD: file input not found/attached. {e}")

    file_input.set_input_files(str(test_file_path))

    try:
        expect(file_input).to_have_value(
            re.compile(re.escape(test_file_path.name), re.I),
            timeout=5000,
        )
    except Exception:
        pass

    start_btn = page.get_by_role("button", name=re.compile(r"^Start$", re.I))

    try:
        start_btn.wait_for(state="visible", timeout=15000)
    except Exception as e:
        pytest.fail(f"UPLOAD: Start button did not become visible. {e}")

    expect(start_btn).to_be_enabled(timeout=5000)

    start_btn.click()
# ============================================================
# PREVIEW / FORMAT SELECTION FLOW
# ============================================================
def open_preview_convert_screen(page):
    rem_doc = page.get_by_role("button", name="Remediate Full Document")
    select_format = page.get_by_role("button", name="Select Format")

    page.wait_for_load_state("domcontentloaded")

    try:
        if rem_doc.is_visible(timeout=1500):
            pytest.fail(
                "Visible Button: Remediate Full Document. You are not logged into an account."
            )
    except Exception as e:
        if "Visible Button: Remediate Full Document" in str(e):
            raise
        pass

    try:
        select_format.wait_for(state="visible", timeout=300000)
        return {"screen_type": "format_selection"}
    except Exception:
        pass

    pytest.fail(
        "Expected Select Format screen was not detected after Start. "
        "The workflow did not stay in the required pipeline."
    )

def choose_conversion_format(page):
    option_locator = page.locator("[role='option'], [role='menuitem'], li, button")
    option_locator.first.wait_for(state="visible", timeout=4000)

    option_count = option_locator.count()

    for index in range(option_count):
        element = option_locator.nth(index)
        element_txt = element.inner_text().strip()

        if not element_txt:
            continue

        if element_txt == "Read full version in browser":
            element.click()
            return

    pytest.fail("Read full version in browser option not found in dropdown.")

def wait_for_preview_ready(page):
    processing = page.get_by_text("Processing...")
    document_title = page.locator("h1")

    page.wait_for_load_state("domcontentloaded")

    try:
        processing.wait_for(state="visible", timeout=3000)
        processing.wait_for(state="hidden", timeout=1000)
    except Exception:
        pass

    document_title.wait_for(state="visible", timeout=10000)

    title = document_title.inner_text().strip()

    assert title, "Preview loaded, but document title was empty."

    return title
# ============================================================
# FOCUS CHECK HELPERS
# ============================================================
def get_focused_element_details(page):
    try:
        focus_data = page.evaluate("""
        () => {
            const el = document.activeElement;

            return {
                name: el ? el.getAttribute("name") : null,
                text: el ? (el.innerText || el.value || "") : null,
                role: el ? el.getAttribute("role") : null
            };
        }
        """)

        return focus_data

    except Exception as e:
        return {
            "name": None,
            "text": None,
            "role": None,
            "error": str(e),
        }

def check_focus_after_format_dropdown_opens(page):

    focus_data = get_focused_element_details(page)

    assert focus_data and not focus_data.get("error"), (
        f"Failed to retrieve focus data after opening format dropdown: {focus_data}"
    )

    focus_text = (focus_data.get("text") or "").lower().strip()
    focus_role = (focus_data.get("role") or "").lower().strip()

    valid_focus = (
        "read full version in browser" in focus_text
        or focus_role in {"option", "menuitem"}
    )

    assert valid_focus, (
        "Focus did not move logically into the format dropdown after it opened.\n"
        f"Focus data: {focus_data}"
    )

def check_focus_event_preview(page, expected_title):
    pass_threshold = 80

    expected_focus_parts = [
        "This document has x page/pages",
        "Document complete.",
    ]

    if expected_title and expected_title.strip():
        expected_focus_parts.append(expected_title)

    focus_data = get_focused_element_details(page)

    assert focus_data and not focus_data.get("error"), (
        f"Failed to retrieve focus data: {focus_data}"
    )

    focus_name = focus_data.get("name")
    focus_text = focus_data.get("text")
    focus_role = focus_data.get("role")

    assert focus_name or focus_text or focus_role, (
        "No focused element found. "
        f"Focus data: {focus_data}"
    )

    combined_text = ((focus_name or "") + " " + (focus_text or "")).lower().strip()

    matched_parts = []
    missing_parts = []

    for piece in expected_focus_parts:
        normalized_piece = piece.lower().strip()

        if piece == "This document has x page/pages":
            if "this document has" in combined_text and (
                "page" in combined_text or "pages" in combined_text
            ):
                matched_parts.append(piece)
            else:
                missing_parts.append(piece)

        elif normalized_piece in combined_text:
            matched_parts.append(piece)

        else:
            missing_parts.append(piece)

    total_parts = len(expected_focus_parts)
    matched_count = len(matched_parts)

    if total_parts == 0:
        matched_percent = 0
    else:
        matched_percent = (matched_count / total_parts) * 100
        matched_percent = round(matched_percent, 2)

    assert matched_percent >= pass_threshold, (
        f"Focus check failed.\n"
        f"Matched percent: {matched_percent:.0f}%\n"
        f"Matched parts: {matched_parts}\n"
        f"Missing parts: {missing_parts}\n"
        f"Focus details: {focus_data}"
    )

    return {
        "status": "PASS",
        "matched_percent": matched_percent,
        "matched_parts": matched_parts,
        "missing_parts": missing_parts,
        "focus_data": focus_data,
    }

def check_focus_event_convert(page):

    convert_btn = page.get_by_role("button", name="Convert", exact=True)
    convert_section = page.locator("body")

    expect(convert_btn).to_be_visible(timeout=5000)
    expect(convert_btn).to_be_enabled(timeout=5000)

    focus_data = get_focused_element_details(page)

    assert focus_data and not focus_data.get("error"), (
        f"Failed to retrieve focus data on convert screen: {focus_data}"
    )

    expected_parts = [
        "Read full version in browser",
        "Skip Settings and Convert",
        "Language",
        "Cancel",
    ]

    convert_section_text = convert_section.inner_text()
    assert convert_section_text, "The convert section text could not be read."

    clean_text = convert_section_text.lower().strip()

    missing_parts = [
        part for part in expected_parts
        if part.lower().strip() not in clean_text
    ]

    assert not missing_parts, (
        f"Expected content in the convert section is missing.\n"
        f"Missing parts: {missing_parts}\n"
        f"Focus data: {focus_data}"
    )

    focus_text = (focus_data.get("text") or "").lower().strip()
    focus_name = (focus_data.get("name") or "").lower().strip()
    focus_role = (focus_data.get("role") or "").lower().strip()

    combined_focus = f"{focus_name} {focus_text}".strip()

    valid_focus = (
            "skip settings and convert" in combined_focus
            or "convert" in combined_focus
            or "language" in combined_focus
            or "cancel" in combined_focus
            or focus_role in {"button", "combobox", "listbox"}
    )

    assert valid_focus, (
        "Focus is not on a logical element on the convert screen.\n"
        f"Focus data: {focus_data}"
    )

def check_focus_final(page, expected_title=None, expected_content_snippets=None):

    expected_final_parts = [
        "Document ready.",
    ]

    if expected_content_snippets is None:
        expected_content_snippets = []

    if expected_title and expected_title.strip():
        expected_content_snippets.append(expected_title)

    page.get_by_text("Document ready.").wait_for(state="visible", timeout=300000)

    completed_screen_text = page.locator("body").inner_text().lower().strip()
    expected_parts = expected_final_parts + expected_content_snippets

    matched_parts = []
    missing_parts = []

    for part in expected_parts:
        normalized_focus_part = part.lower().strip()

        if normalized_focus_part in completed_screen_text:
            matched_parts.append(part)
        else:
            missing_parts.append(part)

    total_parts = len(expected_parts)
    matched_count = len(matched_parts)

    matched_percent = (matched_count / total_parts) * 100 if total_parts else 0

    pass_threshold = 80

    focus_text = page.evaluate("document.activeElement.innerText")
    focus_role = page.evaluate('document.activeElement.getAttribute("role")')

    assert focus_text and focus_text.strip(), "No focus target text was found."

    normalized_focus_element = focus_text.lower().strip()
    valid_focus_element = False
    acceptable_focus_indicators = expected_final_parts + expected_content_snippets

    for indicator in acceptable_focus_indicators:
        normalized_indicator = indicator.lower().strip()

        if normalized_indicator in normalized_focus_element:
            valid_focus_element = True
            break

    assert matched_percent >= pass_threshold and valid_focus_element, (
        f"Final screen validation failed.\n"
        f"Matched percent: {matched_percent:.0f}%\n"
        f"Matched parts: {matched_parts}\n"
        f"Missing parts: {missing_parts}\n"
        f"Focus text: {focus_text}\n"
        f"Focus role: {focus_role}"
    )
# ============================================================
# CONVERSION FLOW
# ============================================================
def click_convert_button(page):
    candidate_names = [
        re.compile(r"^Skip Settings and Convert$", re.I),
        re.compile(r"^Convert$", re.I),
    ]

    last_error = None

    for pattern in candidate_names:
        try:
            btn = page.get_by_role("button", name=pattern).first
            btn.wait_for(state="visible", timeout=10000)
            expect(btn).to_be_enabled(timeout=5000)
            btn.click()
            return
        except Exception as e:
            last_error = e

    raise Exception(f"Could not click the Convert button. Last error: {last_error}")

def run_conversion_and_check_ready_state(page, expected_title):
    expected_parts = [
        "Document ready.",
    ]

    if expected_title and expected_title.strip():
        expected_parts.append(expected_title)

    click_convert_button(page)

    processing = page.get_by_text("Processing...")
    processing.wait_for(state="visible", timeout=300000)
    processing.wait_for(state="hidden", timeout=300000)

    completed_screen_text = page.locator("body").inner_text().lower().strip()

    matched_parts = []
    missing_parts = []

    for part in expected_parts:
        normalize_focus_part = part.lower().strip()

        if part == "This document has x page/pages":
            if "this document has" in completed_screen_text and (
                "page" in completed_screen_text or "pages" in completed_screen_text
            ):
                matched_parts.append(part)
            else:
                missing_parts.append(part)

        elif normalize_focus_part in completed_screen_text:
            matched_parts.append(part)

        else:
            missing_parts.append(part)

    total_parts = len(expected_parts)
    matched_count = len(matched_parts)

    matched_percent = (matched_count / total_parts) * 100
    pass_threshold = 80

    assert matched_percent >= pass_threshold, (
        f"Completed screen validation failed.\n"
        f"Matched percent: {matched_percent:.0f}%\n"
        f"Matched parts: {matched_parts}\n"
        f"Missing parts: {missing_parts}"
    )
# ============================================================
# ORCHESTRATOR
# ============================================================
def run_focus_change_events_test(page, test_file_path):
    upload_test_file(page, test_file_path)

    screen_result = open_preview_convert_screen(page)
    screen_type = screen_result.get("screen_type")

    if screen_type == "format_selection":
        preview_focus_result = check_focus_event_preview(page, "")

        assert preview_focus_result.get("status") == "PASS", preview_focus_result

        select_format_button = page.get_by_role("button", name="Select Format")
        expect(select_format_button).to_be_enabled(timeout=5000)
        select_format_button.click()

        check_focus_after_format_dropdown_opens(page)

        choose_conversion_format(page)

    elif screen_type != "preview_ready":
        pytest.fail(f"Unknown screen type returned from open_preview_convert_screen: {screen_type}")

    expected_title = wait_for_preview_ready(page)

    check_focus_event_convert(page)

    run_conversion_and_check_ready_state(page, expected_title)

    check_focus_final(page, expected_title)
# ============================================================
# TEST FUNCTION
# ============================================================
@pytest.mark.browser_context_args(storage_state=str(STATE))
def test_browser_extension(page: Page):
    test_cases, unsupported_extensions = build_test_cases(str(SAMPLE_DIR))
    test_file_path = test_cases[0]

    reset_upload_state(page, TIMEOUT_SECONDS)

    run_focus_change_events_test(page, test_file_path)