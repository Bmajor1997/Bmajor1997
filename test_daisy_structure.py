# ==========================================================
# Project Name: test_daisy_structure
# ==========================================================

# ==========================================================
# DIRECTIONS
# ==========================================================
"""
This project uploads a source document to ScribeIt, converts it to DAISY ZIP format, validates the generated output,
performs accuracy comparisons between the source and converted content.
"""
# ==========================================================
# IMPORTS
# ==========================================================
from playwright.sync_api import Page, expect
from pathlib import Path
import zipfile
import pytest
import time
import re
from pypdf import PdfReader
# ==========================================================
# CONSTANTS
# ==========================================================
WEBSITE = "https://test.scribeit.io/"
BASE_DIR = Path(__file__).resolve().parent
STATE = BASE_DIR / "storage_state.json"
TEST_DIR = BASE_DIR / "test_documents" / "PDF"
OUTPUT_DIR = BASE_DIR / "daisy_reports"
# ======================================================
# HELPERS
# ======================================================
def normalize_text(text):
    return re.sub(r"\s+", " ", text).strip().lower()

def extract_pdf_text(pdf_path):
    reader = PdfReader(str(pdf_path))
    text = ""

    for page in reader.pages:
        text += page.extract_text() or ""

    return normalize_text(text)
#======================================================
#SOURCE CODE
#======================================================
def get_test_document_path():
    files = list(TEST_DIR.glob("*"))
    if not files:
        raise Exception("No test documents found.")
    return files[0]

def convert_file_to_daisy_zip_via_scribe(page, file_path):

    page.goto(WEBSITE, wait_until="domcontentloaded")

    # Upload
    file_input = page.locator("input[type='file']").first
    expect(file_input).to_be_attached()
    file_input.set_input_files(str(file_path))

    # Click Start
    start_button = page.get_by_role("button", name="Start")
    expect(start_button).to_be_visible()
    expect(start_button).to_be_enabled()
    start_button.click()

    # Wait for Select Format
    select_format_button = page.get_by_role("button", name="Select Format")
    expect(select_format_button).to_be_visible(timeout=300000)
    expect(select_format_button).to_be_enabled(timeout=300000)
    select_format_button.click()

    # Select DAISY
    daisy_option = page.get_by_text("DAISY", exact=False)
    expect(daisy_option).to_be_visible()
    daisy_option.click()

    # Convert
    convert_button = page.get_by_role("button", name="Skip Settings and Convert")
    expect(convert_button).to_be_visible()
    expect(convert_button).to_be_enabled()
    convert_button.click()

    # Download
    page.wait_for_url("**/daisy", timeout=300000)

    download_btn = page.get_by_role("link", name="Download", exact=False)
    expect(download_btn).to_be_visible(timeout=300000)
    expect(download_btn).to_be_enabled(timeout=300000)

    with page.expect_download(timeout=300000) as download_info:
        download_btn.click()

    download = download_info.value

    OUTPUT_DIR.mkdir(exist_ok=True)

    timestamp = int(time.time())
    output_path = OUTPUT_DIR / f"{timestamp}_{file_path.stem}.zip"

    download.save_as(str(output_path))

    return output_path

def validate_daisy_zip(zip_path, source_file):

    if not zipfile.is_zipfile(zip_path):
        return {"status": "FAIL", "reason": "Not a valid ZIP file"}

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        file_list = zip_ref.namelist()

    if not file_list:
        return {"status": "FAIL", "reason": "ZIP is empty"}

    has_opf = any(file.endswith(".opf") for file in file_list)
    has_ncx = any(file.endswith(".ncx") for file in file_list)

    if not has_opf:
        return {"status": "FAIL", "reason": "ZIP file missing:.opf"}

    if not has_ncx:
        return {"status": "FAIL", "reason": "ZIP file missing:.ncx"}

    has_html = any(file.endswith(".html") for file in file_list)
    has_xhtml = any(file.endswith(".xhtml") for file in file_list)
    has_xml = any(file.endswith(".xml") for file in file_list)
    has_smil = any(file.endswith(".smil") for file in file_list)

    if not (has_html or has_xhtml or has_xml or has_smil):
            return {"status": "FAIL", "reason": "ZIP file missing: .html, .xhtml, xml, or .smil"}

    expected_title = Path(source_file).stem.lower()

    searchable_files = [
        file for file in file_list
        if Path(file).suffix.lower() in {".opf", ".ncx", ".html", ".xhtml", ".xml", ".smil"}
    ]

    package_text = ""

    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        for file in searchable_files:
            package_text += zip_ref.read(file).decode("utf-8", errors="ignore").lower()

    if expected_title not in package_text:
        return {
            "status": "FAIL",
            "reason": "Expected source title/name was not found in the DAISY package"
        }

    non_empty_file_count = 0

    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        for file in file_list:
            file_size = zip_ref.getinfo(file).file_size

            if file_size > 0:
                non_empty_file_count += 1

    if non_empty_file_count == 0:
        return {
            "status": "FAIL",
            "reason": "Every file in the DAISY package equals 0 bytes."
        }

    source_text = extract_pdf_text(source_file)
    daisy_text = normalize_text(package_text)

    if source_text[:200] not in daisy_text:
        return {
            "status": "FAIL",
            "reason": "Expected source document text was not found in the DAISY output"
        }

    return {"status": "PASS", "files": file_list}
# ==========================================================
# TEST
# ==========================================================
@pytest.mark.browser_context_args(storage_state=str(STATE))
def test_daisy_structure(page: Page):

    source_file = get_test_document_path()

    daisy_zip_path = convert_file_to_daisy_zip_via_scribe(
        page,
        source_file
    )

    validation = validate_daisy_zip(daisy_zip_path, source_file)

    assert validation["status"] == "PASS", validation.get("reason")
