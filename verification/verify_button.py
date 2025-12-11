
from playwright.sync_api import sync_playwright, expect
import os

def run(playwright):
    browser = playwright.chromium.launch(headless=True)
    page = browser.new_page()

    # Enable console logging
    page.on("console", lambda msg: print(f"PAGE LOG: {msg.text}"))
    page.on("pageerror", lambda err: print(f"PAGE ERROR: {err}"))

    # Get absolute path
    file_path = os.path.abspath("verification/mock_report_settings.html")
    print(f"Loading: file://{file_path}")
    page.goto(f"file://{file_path}")

    # --- Scenario 1: Initial Load (No Settings) ---
    print("Waiting for select-source-file-btn to be enabled...")
    # Wait for initialization with a shorter timeout for debugging
    try:
        page.wait_for_selector("#select-source-file-btn:not([disabled])", timeout=5000)
    except Exception as e:
        print("Timeout waiting for button enable. Checking status div...")
        status = page.text_content("#status")
        print(f"Status div content: {status}")
        # Take a screenshot to see what's wrong
        page.screenshot(path="verification/debug_timeout.png")
        raise e

    # Check that Select Sheet button is DISABLED initially
    select_sheet_btn = page.locator("#select-sheet-btn")
    expect(select_sheet_btn).to_be_disabled()

    # Take screenshot 1
    page.screenshot(path="verification/step1_disabled.png")
    print("Step 1: Button is disabled as expected.")

    # --- Scenario 2: Simulate File Pick ---
    print("Simulating file pick...")
    # Click select file button (triggers mock picker creation)
    page.click("#select-source-file-btn")

    # Simulate picking a file
    page.evaluate("window.simulateFilePick('file123', 'My Data.xlsx')")

    # Check that Select Sheet button is ENABLED now
    expect(select_sheet_btn).to_be_enabled()

    # Take screenshot 2
    page.screenshot(path="verification/step2_enabled_after_pick.png")
    print("Step 2: Button is enabled after picking file.")

    # --- Scenario 3: Reload with Existing Settings ---
    print("Reloading with mock settings...")
    # Reload page, but inject settings this time
    page.add_init_script("window.mockSettings = {sourceFileId: 'saved123', sheetName: 'Sheet1'};")
    page.reload()

    # Wait for initialization
    page.wait_for_selector("#select-source-file-btn:not([disabled])")

    # Check that Select Sheet button is ENABLED immediately
    expect(select_sheet_btn).to_be_enabled()

    # Take screenshot 3
    page.screenshot(path="verification/step3_enabled_from_settings.png")
    print("Step 3: Button is enabled from saved settings.")

    # --- Create composite screenshot for verification ---
    from PIL import Image

    img1 = Image.open("verification/step1_disabled.png")
    img2 = Image.open("verification/step2_enabled_after_pick.png")
    img3 = Image.open("verification/step3_enabled_from_settings.png")

    # Combine images vertically
    dst = Image.new('RGB', (img1.width, img1.height + img2.height + img3.height))
    dst.paste(img1, (0, 0))
    dst.paste(img2, (0, img1.height))
    dst.paste(img3, (0, img1.height + img2.height))

    dst.save("verification/combined_verification.png")

    browser.close()

with sync_playwright() as playwright:
    run(playwright)
