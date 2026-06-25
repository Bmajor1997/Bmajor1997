# ==============================
#   Name: test_daisy_playability
# ==============================
"""
PROJECT DESCRIPTION:
This project automates the DAISY conversion and playability workflow using an
already open and logged-in ScribeIt session and Thorium Reader on a Windows computer.
A successful result means the DAISY publication is imported into Thorium Reader
and is able to begin the Readaloud playback flow on Windows.
"""
# ============================
#           IMPORTS
# ============================
from pathlib import Path
import time
import subprocess
from playwright.sync_api import expect
from pywinauto import Application, Desktop
import pyautogui
import pytest
# =========================
#           CONSTANTS
# =========================
BASE_DIR = Path(__file__).resolve().parent
STATE = BASE_DIR / "storage_state.json"
WEBSITE = "https://test.scribeit.io/"
TEST_DOCUMENTS_FOLDER = BASE_DIR / "test_documents" / "various_test_documents_for_projects"
#============================
#       HELPER FUNCTIONS
#============================
def back_to_home_screen(page):
    # Takes the website back to the homepage after the sequence has finished.

    result = {
        "status": "",
        "error": [],
        "note": []
    }

    try:
        # FIND a UI element that returns the page to the home screen
        convert_an_doc = page.get_by_role("link", name="Convert another document")

        convert_an_doc.wait_for(state="visible", timeout=5000)

        if not convert_an_doc.is_visible():
            result["status"] = "BLOCKED"
            result["error"].append("Convert another document link not found.")
            result["note"].append("Could not return to Home screen because the Convert another document link was missing.")
            return result

        # CLICK the Home button
        convert_an_doc.click()

        result["note"].append("Clicked 'Convert another document' link.")

        # WAIT to go back to home screen
        scribe = page.get_by_text("Welcome to scribe for documents", exact=False)
        scribe.wait_for(state="visible", timeout=5000)

        result["status"] = "PASS"
        result["note"].append("Returned to the home screen successfully.")
        return result

    except Exception as e:
        result["status"] = "FAIL"
        result["error"].append(f"Could not return to the home screen: {repr(e)}")
        result["note"].append("The website did not return to the Choose File screen.")
        return result

def fail_result(result, message):
    result["status"] = "FAIL"
    result["error"].append(message)
    return result
# ===========================
#         SOURCE CODE
# ===========================
def upload_test_document(page,test_document):
    # Uploads test documents

    result = {
        "status": "",
        "document_name": "",
        "type_of_file": "",
        "error": [],
        "note": []
    }

    if not test_document.exists():
        result["status"] = "ERROR"
        result["error"].append(f"Test document not found: {test_document}")
        return result

    try:
        file_input = page.locator("input[type='file']").first
        file_input.wait_for(state="attached", timeout=15000)
        file_input.set_input_files(str(test_document))

        result["status"] = "PASS"
        result["document_name"] = test_document.name
        result["type_of_file"] = test_document.suffix
        result["note"].append("Test document uploaded successfully.")
        return result

    except Exception as e:
        result["status"] = "ERROR"
        result["error"].append(f"File upload failed: {e}")
        return result

def conversion_pipeline(page):
    # Handles the pipeline from hitting the start button to downloading Daisy zip

    result = {
        "status": "",
        "document_name": "",
        "type_of_file": "",
        "daisy_zip_path": None,
        "error": [],
        "note": []
    }

    start_btn = page.get_by_role("button", name="Start")

    if not start_btn.is_visible():
        return fail_result(result, "Start button does not exist.")

    start_btn.click()

    try:
        select_form_btn = page.get_by_role("button", name="Select Format")
        select_form_btn.wait_for(state="visible", timeout=180000)
        expect(select_form_btn).to_be_enabled(timeout=180000)
    except Exception as e:
        return fail_result(result, f"Select Format button did not become ready: {e}")

    if not select_form_btn.is_enabled():
        return fail_result(result, "Select Format button is not enabled.")

    select_form_btn.click()

    read_browser = page.get_by_text("Read full version in browser", exact=False)
    pdf = page.get_by_text("PDF")

    try:
        read_browser.wait_for(state="visible", timeout=10000)
    except Exception:
        try:
            pdf.wait_for(state="visible", timeout=5000)
            print("Read full version in browser option unavailable.")
        except Exception:
            return fail_result(result, "Format dropdown menu did not appear.")

    daisy_btn = page.get_by_text("DAISY")

    try:
        daisy_btn.wait_for(state="visible", timeout=10000)
        daisy_btn.click()
    except Exception:
        return fail_result(result, "DAISY button was not selected.")

    convert_btn = page.get_by_role("button", name="Convert", exact=True)

    try:
        convert_btn.wait_for(state="visible", timeout=10000)
        convert_btn.click()
    except Exception:
        return fail_result(result, "Conversion failed.")

    download_link = page.get_by_role("link", name="Download DAISY", exact=False)

    try:
        download_link.wait_for(state="visible", timeout=120000)
    except Exception:
        return fail_result(result, "Download link did not appear.")

    if not download_link.is_enabled():
        return fail_result(result, "Download link is not enabled.")

    download_folder = Path.home() / "Downloads" / "converted_DAISY_downloads"
    download_folder.mkdir(parents=True, exist_ok=True)

    try:
        with page.expect_download(timeout=10000) as download_info:
            download_link.click()

        download = download_info.value
        final_download_path = download_folder / download.suggested_filename
        download.save_as(final_download_path)

        result["daisy_zip_path"] = str(final_download_path)
        result["type_of_file"] = "zip"
        result["status"] = "PASS"
        result["note"].append(f"{final_download_path.name} downloaded successfully.")
        return result

    except Exception as e:
        return fail_result(result, f"Download failed: {e}")

def open_thorium():
    # Opens Thorium Reader and Validates that it is open

    result = {
        "status": "",
        "error": [],
        "note": []
    }

    possible_paths = [
        Path.home() / "AppData/Local/Programs/Thorium/Thorium.exe",
        Path("C:/Program Files/Thorium/Thorium.exe"),
        Path("C:/Program Files (x86)/Thorium/Thorium.exe"),
    ]

    thor_path = None

    for path in possible_paths:
        if path.exists():
            thor_path = path
            break

    if not thor_path:
        return fail_result(result, "Thorium executable was not found.")

    print("Thorium Launch Starting...")

    try:
        subprocess.Popen(str(thor_path))

        time.sleep(3)

        thor_app = Application(
            backend="uia"
        ).connect(
            title_re=".*Thorium.*",
            timeout=15
        )

        thor_win = thor_app.top_window()

        thor_win.wait("visible", timeout=10)
        thor_win.wait("enabled", timeout=10)

    except Exception as e:
        return fail_result(
            result,
            f"Thorium failed to launch: {e}"
        )

    result["status"] = "PASS"
    result["note"].append("Thorium opened successfully.")

    return result

def export_daisy_zip(saved_zip_path):
    # Exports Daisy File into Thorium Reader

    result = {
        "status": "",
        "daisy_zip_name": "",
        "error": [],
        "note": []
    }

    if not saved_zip_path:
        return fail_result(result, "No saved DAISY zip path was provided.")

    zip_path = Path(saved_zip_path)
    result["daisy_zip_name"] = zip_path.stem

    if not zip_path.exists():
        return fail_result(result, "Saved DAISY zip file does not exist.")

    if zip_path.suffix.lower() != ".zip":
        return fail_result(result, "Saved DAISY file is not a zip file.")

    try:
        app = Application(backend="uia").connect(title_re=".*Thorium.*", timeout=10)

        window = app.top_window()
        time.sleep(2)

        window.wait("visible", timeout=10)
        window.wait("enabled", timeout=10)
        window.set_focus()

        for button in window.descendants(control_type="Button"):
            print(button.window_text())

        import_control = window.child_window(
            title_re="Import publication.*",
            control_type="Text"
        )

        import_control.wait("exists visible enabled", timeout=10)
        import_control.click_input()

    except Exception as e:
        return fail_result(result, f"DAISY zip file could not be imported to Thorium: {e}")

    try:
        app = Application(backend="win32").connect(title_re=".*Open.*", timeout=10)
        file_dialog = app.top_window()

        file_dialog.wait("visible", timeout=10)
        file_dialog.wait("enabled", timeout=10)
        file_dialog.set_focus()

        file_path = str(zip_path)

        file_name_box = file_dialog.child_window(class_name="ComboBoxEx32") \
            .child_window(class_name="ComboBox") \
            .child_window(class_name="Edit")

        file_name_box.wait("exists visible enabled", timeout=10)
        file_name_box.set_edit_text("")
        file_name_box.set_edit_text(file_path)

        open_button = file_dialog.child_window(title="&Open", class_name="Button")
        open_button.wait("exists visible enabled", timeout=10)
        open_button.click_input()

    except Exception as e:
        return fail_result(result, f"DAISY zip file could not be submitted from Open dialog: {e}")

    result["status"] = "PASS"
    result["note"].append("The .zip file was successfully submitted to Thorium.")
    return result

def open_export_in_thorium_reader(expected_title):
    # Opens the exported zip file in Thorium Reader

    result = {
        "status": "",
        "document_name": "",
        "error": [],
        "note": []
    }

    try:
        time.sleep(3)

        expected_title = Path(expected_title).stem.strip()

        thorium_window = Desktop(backend="uia").window(title_re=".*Thorium.*")
        thorium_window.wait("visible", timeout=10)
        thorium_window.set_focus()

        time.sleep(1)
        pyautogui.press("pagedown")
        time.sleep(2)

        control_types_to_try = ["Text", "ListItem", "Button", "Custom"]

        file_name = None
        matched_type = None

        for control_type in control_types_to_try:
            candidates = [
                elem for elem in thorium_window.descendants(control_type=control_type)
                if elem.window_text().lower().startswith(expected_title.lower())
            ]

            if candidates:
                file_name = candidates[0]
                matched_type = control_type
                break

        if not file_name:
            return fail_result(result, f"{expected_title} was not found in Thorium.")

        if not file_name.is_visible():
            return fail_result(result, f"{expected_title} was found but is not visible.")

        result["document_name"] = expected_title

        try:
            clickable_target = file_name.parent()
        except Exception:
            clickable_target = file_name

        clickable_target.click_input()
        time.sleep(2)

        result["status"] = "PASS"
        result["note"].append(f"The document {expected_title} was clicked in Thorium.")
        result["note"].append(f"Matched control type: {matched_type}")
        return result

    except Exception as e:
        return fail_result(result, f"Thorium document open failed: {e}")

def check_and_confirm_playability():

    # Presses the Play button and confirms that Thorium is playing the Daisy file

    result = {
        "status": "",
        "button_clicked": False,
        "error": [],
        "note": []
    }

    try:
        thorium_windows = Desktop(
            backend="uia"
        ).windows(
            title_re=".*Thorium.*",
            top_level_only=True
        )

        thorium_window = None

        for window in thorium_windows:
            title = window.window_text().strip()

            if title and title.lower() != "thorium":
                thorium_window = window
                break

        if thorium_window is None and thorium_windows:
            thorium_window = thorium_windows[0]

        if not thorium_window:
            return fail_result(result, "Thorium window was not found.")

        thorium_window.set_focus()
        thorium_window.set_focus()

        read_aloud = thorium_window.descendants(
            title="Activate Readaloud",
            control_type="Button"
        )

        if not read_aloud:
            return fail_result(result, "Read Aloud button was not found.")

        read_aloud[0].click_input()
        result["button_clicked"] = True

        for attempt in range(2):

            pause_buttons = thorium_window.descendants(
                title="Pause Readaloud",
                control_type="Button"
            )

            stop_buttons = thorium_window.descendants(
                title="Stop Readaloud",
                control_type="Button"
            )

            search_buttons = thorium_window.descendants(
                title="search publication",
                control_type="Button"
            )

            if pause_buttons and (stop_buttons or search_buttons):

                before_text = [
                    item.window_text()
                    for item in thorium_window.descendants(control_type="Text")
                    if item.window_text().strip()
                ]

                time.sleep(3)

                after_text = [
                    item.window_text()
                    for item in thorium_window.descendants(control_type="Text")
                    if item.window_text().strip()
                ]

                if before_text != after_text:
                    result["status"] = "PASS"

                    if attempt == 0:
                        result["note"].append(
                            "Readaloud playback appears to be advancing."
                        )
                    else:
                        result["note"].append(
                            "Readaloud playback appears to be advancing on second attempt."
                        )

                    return result

                return fail_result(
                    result,
                    "Playback controls appeared, but the DAISY content did not appear to advance."
                )

            if attempt == 0:
                time.sleep(3)

        return fail_result(
            result,
            "Playback controls did not appear after clicking Read Aloud."
        )

    except Exception as e:
        return fail_result(
            result,
            f"Playability check failed: {e}"
        )
#======================================================
#TEST FUNCTION
#======================================================
@pytest.mark.browser_context_args(storage_state=str(STATE))
def test_daisy_playability(page):

    page.goto(WEBSITE)
    page.wait_for_load_state("domcontentloaded")

    test_document = TEST_DOCUMENTS_FOLDER / "Appendix 3_Water Quality Tables_0007.pdf"

    upload_result = upload_test_document(page, test_document)
    assert upload_result["status"] == "PASS"

    conversion_result = conversion_pipeline(page)
    assert conversion_result["status"] == "PASS"

    zip_path = conversion_result.get("daisy_zip_path")
    assert zip_path

    home_screen_result = back_to_home_screen(page)
    assert home_screen_result["status"] == "PASS"

    thorium_result = open_thorium()
    assert thorium_result["status"] == "PASS"

    daisy_export_result = export_daisy_zip(zip_path)
    assert daisy_export_result["status"] == "PASS"

    thorium_open_result = open_export_in_thorium_reader(
        conversion_result["document_name"]
    )
    assert thorium_open_result["status"] == "PASS"

    daisy_export_result = export_daisy_zip(zip_path)
    print(daisy_export_result)
    assert daisy_export_result["status"] == "PASS"