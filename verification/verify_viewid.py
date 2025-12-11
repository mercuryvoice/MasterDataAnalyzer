
from playwright.sync_api import sync_playwright, expect
import os

def run(playwright):
    browser = playwright.chromium.launch(headless=True)
    page = browser.new_page()

    # Enable console logging
    page.on("console", lambda msg: print(f"PAGE LOG: {msg.text}"))

    # Get absolute path
    file_path = os.path.abspath("verification/mock_report_settings_viewid.html")
    print(f"Loading: file://{file_path}")
    page.goto(f"file://{file_path}")

    # Wait for initialization
    page.wait_for_selector("#select-source-file-btn:not([disabled])")

    # Click select file button
    print("Clicking select file button...")
    page.click("#select-source-file-btn")

    # Verify that the Picker was built with SPREADSHEETS view
    # We exposed 'window.lastPickerViewId' in our mock
    view_id = page.evaluate("window.lastPickerViewId")
    print(f"Captured View ID: {view_id}")

    if view_id == 'SPREADSHEETS':
        print("SUCCESS: Picker is using SPREADSHEETS view.")
    else:
        print(f"FAILURE: Picker is using {view_id} view.")
        raise Exception(f"Expected SPREADSHEETS view, got {view_id}")

    page.screenshot(path="verification/picker_view_id.png")
    browser.close()

with sync_playwright() as playwright:
    run(playwright)
