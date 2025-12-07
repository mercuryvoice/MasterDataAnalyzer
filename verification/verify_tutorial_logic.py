from playwright.sync_api import sync_playwright
import os

# Mock translations based on MasterDataTranslation.js
MOCK_TRANSLATIONS_EN = {
    "tutorialPrev": "Previous",
    "tutorialNext": "Next",
    "tutorialClose": "Close",
    "tutorialYes": "Yes, continue",
    "tutorialNo": "No, I'm done for now",
    "tutorialProgressText": "Step {CURRENT} / {TOTAL}",
    "endTutorialTitle": "Tutorial Paused",
    "endTutorialInstruction": "Great! You have completed the 'Data Import' section.\nYou can reopen this tutorial at any time or continue exploring other features.",
    "endTutorialReturnButton": "Return to 'Task 1 Complete' step",
    "businessReturnToStep13": "Return to Step 13 Congratulations, Analysis Complete!",

    # Business Steps
    "businessWelcomeTitle": "Welcome! (Business & Sales)",
    "businessWelcomeInstruction": "Welcome to the interactive tutorial for the <code>Business & Sales Example</code>.\n\nOur goal is to transform an incomplete sales record into a complete analysis report using the features of MasterDataAnalyzer.\n\nClick <code>Next</code> to start our first task!",
    "businessTask1Step1Title": "Task 1: Data Import (1/6)",
    "businessTask1Step1Instruction": "First, please ensure you have activated the <code>{SHEET_NAME}</code> sheet.\n\nOur goal is to filter and import sales records from <code>{SOURCE_SHEET_NAME}</code> into the current dashboard.",
    "businessTask1Step2Title": "Task 1: Data Import (2/6)",
    "businessTask1Step2Instruction": "Please click on <code>MasterDataAnalyzer > Data Import Tool > ⚙️ Data Import Settings</code> in the top menu.",
    "businessTask1Step3Title": "Task 1: Data Import (3/6)",
    "businessTask1Step3Instruction": "In the settings window, please configure the following:\n1. <code>Source Data Select</code>: Click the <code>Select</code> button and choose the current spreadsheet file from the file picker.\n2. <code>Source Data Sheet Name</code>: Select <code>{SOURCE_SHEET_NAME}</code>\n3. <code>Target Data Sheet Name</code>: Should be auto-filled with <code>{TARGET_SHEET_NAME}</code>",
    "businessTask1Step4Title": "Task 1: Data Import (4/6)",
    "businessTask1Step4Instruction": "Next, set the data ranges:\n1. <code>Header Start Row for Data Import</code>: <code>1</code>\n2. <code>Data Import Start Row</code>: <code>2</code>\n3. <code>Source Data Import Range</code>: <code>A2:E9</code>",
    "businessTask1Step5Title": "Task 1: Data Import (5/6)",
    "businessTask1Step5Instruction": "Please add a filter condition:<br>1. <code>Header Name</code>: <br>Customer ID<br>2. <code>Keywords</code>: <br>Click <code>Select</code> then check <code>Select All</code><br>Reminder: It is recommended to import all data here to facilitate the application of report generation later.<br><br>In <code>Source Header Import Range Settings</code>, you can enter a header range to limit the scope of header detection. If left blank, all headers will be used by default.",
    "businessTask1Step6Title": "Task 1: Data Import (6/6)",
    "businessTask1Step6Instruction": "Great! All settings are complete.\n\nPlease click <code>Save Settings</code>, close the window, and then run <code>MasterDataAnalyzer > Data Import Tool > ▶️ Run Import</code> from the menu.",
    "businessCheckpointTitle": "Task 1 Complete!",
    "businessCheckpointInstruction": "Congratulations! You have successfully imported the raw data into the dashboard.\n\nWould you like to continue learning about the next core feature, <code>Data Validation</code>?",
    "businessTask2Step1Title": "Task 2: Enrich Customer Data (1/3)",
    "businessTask2Step1Instruction": "Excellent! The dashboard now has raw data but lacks detailed customer information.\n\nNext, we'll use the <code>Data Comparison</code> feature to look up and fill in customer data from the <code>Source | Customer List</code>.\n\nOpen <code>Data Import Tool > ⚙️ Data Comparison Settings</code> and begin configuring Task 2:",
    "businessTask2Step2Title": "Task 2: Enrich Customer Data (2/3)",
    "businessTask2Step2Instruction": "Please apply the following settings:\n1. <code>Source Sheets</code>: Choose <code>Source Data Sheet Name</code> as <code>{SOURCE_SHEET_NAME}</code> and <code>Target Data Sheet Name</code> as <code>{TARGET_SHEET_NAME}</code>.\n2. <code>Ranges</code>: Set <code>Target Data Start Row</code> to <code>2</code> and <code>Source Data Compare Range</code> to <code>A2:D6</code>.\n3. <code>Field Mapping</code>: Set up the fields in the following order:\n  - <code>Target Lookup Column</code>: <code>B</code> (Customer ID)\n  - <code>Source Compare Column</code>: <code>A</code> (Customer ID)\n  - <code>Source Return Column</code>: <code>B</code> (Customer Name)\n  - <code>Target Write Column</code>: <code>C</code> (Customer Name)",
    "businessTask2Step3Title": "Task 2: Enrich Customer Data (3/3)",
    "businessTask2Step3Instruction": "After saving the settings, run <code>MasterDataAnalyzer > Data Import Tool > ▶️ Run Data Comparison</code>.\n\nYou will see the customer names have been successfully imported. Repeat this process to also fill in the <code>Region</code> and <code>Salesperson</code> to complete the dashboard.",
    "businessTask3Step1Title": "Task 3: Compare Sales Targets (1/2)",
    "businessTask3Step1Instruction": "The dashboard data is becoming more complete! For the final step, let's compare the sales targets for each product.\n\nPlease open the <code>Data Comparison Settings</code> again.",
    "businessTask3Step2Title": "Task 3: Compare Sales Targets (2/2)",
    "businessTask3Step2Instruction": "This time, use <code>Product Name</code> as the lookup key and select <code>\"Source | Product Targets\"</code> as the source sheet.<br>Set the <code>Source Data Compare Range</code> to <code>A2:B5</code>, and configure the field mapping as follows:<br>  - <code>Target Lookup Column</code>: <code>F</code> (Product Name)<br>  - <code>Source Compare Column</code>: <code>A</code> (Product Name)<br>  - <code>Source Return Column</code>: <code>B</code> (Monthly Target)<br>  - <code>Target Write Column</code>: <code>J</code> (Monthly Target)",
    "businessFinalStepTitle": "Congratulations, Analysis Complete!",
    "businessFinalStepInstruction": "Then, run the <code>Data Comparison</code> again.\n\nAll data is now in place! You now have a clean, complete dataset ready for analysis.\n\nNext, you can manually enter or click the button below to insert formulas in the corresponding columns to complete the final calculations:\n- **Total Sales**: <code>=H2*I2</code> (Quantity * Unit Price)\n- **Achievement Rate**: <code>=G2/J2</code> (Total Sales / Monthly Target)",
    "businessFinalStepButton": "Insert Formulas Automatically",
    "businessProcessingButton": "Processing...",
    "businessSuccessButton": "Formulas Inserted!",
    "businessErrorButton": "Error, please retry",
    "businessCheckpoint2Title": "Task 4: Data Report Generation Settings!",
    "businessCheckpoint2Instruction": "Congratulations! You have successfully transformed the raw data into an analyzable dashboard with the help of MasterDataAnalyzer.\n\nWould you like to continue learning about the next core feature, 'Data Report Generation Settings'?",
    "businessTask4Step1Title": "Task 4: Data Report Generation Settings (1/4)",
    "businessTask4Step1Instruction": "Please click on <code>MasterDataAnalyzer > Data Management Tool > ⚙️ Report Generation Settings</code> in the top menu.",
    "businessTask4Step2Title": "Task 4: Report Data Source Settings (2/4)",
    "businessTask4Step2Instruction": "In the settings window, please configure the following:\n1. <code>Source Data Select</code>: Click the <code>Select</code> button and choose the current spreadsheet file from the file picker.\n2. <code>Source Data Sheet Name</code>: Select <code>{SHEET_NAME}</code>\n3. <code>Source Data Import Range</code>: Please enter the data range to be analyzed from <code>{SHEET_NAME}</code>, for example: <code>A1:K9</code>",
    "businessTask4Step3Title": "Task 4: Report Field Mapping Settings (3/4)",
    "businessTask4Step3Instruction": "1. To generate a report, at least one combination of a dimension (Y) and a metric (X) is required. You can click <code>+Add Analysis Field</code> to map multiple sets for numerical analysis.<br>2. You can follow the example settings:<br><strong>First Set</strong>: Dimension <code>Customer Name</code>, Metric <code>Total Sales</code><br><strong>Second Set</strong>: Dimension <code>Product Name</code>, Metric <code>Monthly Target</code><br><strong>Third Set</strong>: Dimension <code>Salesperson</code>, Metric <code>Achievement Rate</code><br>3. Click <code>Save Settings</code>, then click <code>Generate Report</code>.",
    "businessTask4Step4Title": "Task 4: Generate and Export Report (4/4)",
    "businessTask4Step4Instruction": "Great, you should now see the generated chart report with multiple dimensions and metrics. Please click <code>Confirm Export</code> to generate the report in a new Google Sheet tab.",
    "businessFinal2StepTitle": "Congratulations, Tutorial Complete!",
    "businessFinal2StepInstruction": "Congratulations, you have completed the interactive tutorial for report generation and export!<br>MasterDataAnalyzer has successfully exported the report for you. You can go to the new Google Sheet tab to view or start editing the native report.",

    # Variables
    "exampleDashboardSheet_Sales": "Dashboard | Sales Analysis",
    "exampleSalesLogSheet_Sales": "Source | Sales Log",
    "exampleCustomerMasterSheet_Sales": "Source | Customer List",
    "exampleProductTargetsSheet_Sales": "Source | Product Targets",
    "productLaptop": "High-Performance Laptop",
    "headerProductName": "Product Name"
}

import json

def test_tutorial_updates(page):
    page.on("console", lambda msg: print(f"CONSOLE: {msg.text}"))

    # Read the HTML template
    with open('GuidePageTraining.html', 'r') as f:
        html_template = f.read()

    # Prepare translations
    translations_json = json.dumps(MOCK_TRANSLATIONS_EN).replace("'", "\\'")

    # Inject variables
    html_content = html_template.replace('<?= tutorialType ?>', 'business')
    html_content = html_content.replace('<?= translations ?>', translations_json)

    # Write temporary file
    with open('verification/temp_guide_en.html', 'w') as f:
        f.write(html_content)

    # Open the page
    page.goto(f"file://{os.path.abspath('verification/temp_guide_en.html')}")

    # --- 1. Verify "Return to Step 13" Button Logic ---
    print("Verifying 'Return to Step 13' button...")

    # Wait for hydration
    page.wait_for_timeout(500)

    # Go to Checkpoint 1
    for _ in range(20):
        if page.get_by_text("Yes, continue").is_visible():
            print("Reached Checkpoint 1 (Task 1 Complete). Clicking Yes.")
            page.get_by_text("Yes, continue").click()
            page.wait_for_timeout(500)
            break

        if page.get_by_role("button", name="Next").is_disabled():
            break

        page.get_by_role("button", name="Next").click()
        page.wait_for_timeout(200)

    # Go to Checkpoint 2
    for _ in range(20):
        if page.get_by_role("heading", name="Task 4: Data Report Generation Settings!").is_visible():
            print("Reached Checkpoint 2 (Task 4 Start).")
            # Click 'No, I'm done for now' to trigger endTutorial()
            page.get_by_text("No, I'm done for now").click()
            page.wait_for_timeout(500)
            break

        page.get_by_role("button", name="Next").click()
        page.wait_for_timeout(200)

    expected_button_text = MOCK_TRANSLATIONS_EN["businessReturnToStep13"]

    if page.get_by_role("button", name=expected_button_text).is_visible():
        print(f"SUCCESS: Found button '{expected_button_text}'")
    else:
        print(f"FAILURE: Did not find button '{expected_button_text}'")
        # take screenshot of failure
        page.screenshot(path="verification/failure_button.png")

    page.screenshot(path="verification/verification_en_button.png")

    page.get_by_role("button", name=expected_button_text).click()
    page.wait_for_timeout(500)

    if page.get_by_role("heading", name="Congratulations, Analysis Complete!").is_visible():
        print("SUCCESS: Returned to 'Congratulations, Analysis Complete!' step.")
    else:
        print("FAILURE: Did not return to correct step.")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    test_tutorial_updates(page)
    browser.close()
