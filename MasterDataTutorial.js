// @ts-check

// =================================================================================================
// =================================== VERSION & FEATURE SUMMARY ===================================
// =================================================================================================
//
// V4.11 (UI & Translation Refinement):
// - UI ENHANCEMENT: Added <code> tags to key phrases in the Business Tutorial (Task 3) for
//   better visual emphasis, per user feedback.
// - TRANSLATION FIX: Added new translation keys for the "Insert Formulas" button states
//   (processing, success, error) to ensure they display correctly in all languages.
//
// V4.10 (Code Integrity Restoration):
// - FIX: Restored the complete and correct code structure for the entire script.
//
// =================================================================================================

/**
 * A complete dictionary of all strings for the tutorial sidebar with CORRECT sheet names.
 * @returns {Object.<string, string>} An object containing all UI strings in the correct language.
 */
function getTutorialTranslations() {
    const locale = Session.getActiveUserLocale();

    const translations = {
        en: {
          // ... (other keys)
          businessTask3Step1Instruction: "The dashboard data is becoming more complete! For the final step, let's compare the sales targets for each product.\n\nPlease open the <code>'Data Comparison Settings'</code> again.",
          businessTask3Step2Instruction: "This time, use \"Product Name\" as the lookup key and select <code>\"Source | Product Targets\"</code> as the source sheet.<br>Set the \"Source Data Compare Range\" to <code>A2:B5</code>, and configure the field mapping as follows:<br>  - **Target Lookup Column**: <code>F</code> (Product Name)<br>  - **Source Compare Column**: <code>A</code> (Product Name)<br>  - **Source Return Column**: <code>B</code> (Monthly Target)<br>  - **Target Write Column**: <code>J</code> (Monthly Target)",
          businessFinalStepInstruction: "Then, run the <code>'Data Comparison'</code> again.\n\nAll data is now in place! You now have a clean, complete dataset ready for analysis.\n\nNext, you can manually enter or click the button below to insert formulas in the corresponding columns to complete the final calculations:\n- **Total Sales**: <code>=H2*I2</code> (Quantity * Unit Price)\n- **Achievement Rate**: <code>=G2/J2</code> (Total Sales / Monthly Target)",
          businessFinalStepButton: "Insert Formulas Automatically",
          businessProcessingButton: "Processing...", // NEW
          businessSuccessButton: "Formulas Inserted!", // NEW
          businessErrorButton: "Error, please retry", // NEW
          // ... (other keys)
        },
        'zh_TW': {
          // ... (other keys)
          businessTask3Step1Instruction: "儀表板的資料越來越完整了！最後一步，讓我們來比對每項產品的銷售目標。\n\n請再次打開<code>「資料比對設定」</code>。",
          businessTask3Step2Instruction: "這次，請用 \"產品名稱\" 作為目標對比基準，並選擇 \"來源資料分頁名稱\" 為 <code>{SOURCE_SHEET_NAME}</code>。<br>\"來源資料比對範圍\" 請設定為 <code>A2:B5</code>，\"資料比對欄位對應\" 請依照下方順序逐一設定：<br>  - **目標查找欄位**: <code>F</code> (產品名稱)<br>  - **來源比對欄位**: <code>A</code> (產品名稱)<br>  - **來源返回欄位**: <code>B</code> (目標月銷售額)<br>  - **目標寫入欄位**: <code>J</code> (目標月銷售額)",
          businessFinalStepInstruction: "接著，請再次執行<code>「資料比對」</code>。\n\n所有資料都已到位！您現在擁有了一份乾淨、完整的數據。\n\n接下來，您可以手動或點擊下方按鈕，在對應的欄位中填入公式來完成最後的計算：\n- **總銷售額**: <code>=H2*I2</code> (數量 * 單價)\n- **業績達成率**: <code>=G2/J2</code> (總銷售額 / 目標月銷售額)",
          businessFinalStepButton: "自動帶入公式",
          businessProcessingButton: "處理中...", // NEW
          businessSuccessButton: "公式已成功帶入！", // NEW
          businessErrorButton: "發生錯誤，請重試", // NEW
          // ... (other keys)
        }
    };
    // In a real implementation, you would merge these updates into the full translation object.
    // For this response, I will provide the full, corrected object below.
    return getFullTranslationsWithAllUpdates();
}

// Helper to provide the full, updated translation object
function getFullTranslationsWithAllUpdates() {
    const locale = Session.getActiveUserLocale();
    const allTranslations = {
        en: {
          businessGuideTutorialTitle: 'Business & Sales Example - Tutorial',
          manufacturingGuideTutorialTitle: 'Manufacturing Example - Tutorial',
          tutorialNext: "Next",
          tutorialPrev: "Previous",
          tutorialClose: "Close",
          tutorialYes: "Yes, continue",
          tutorialNo: "No, I'm done for now",
          tutorialProgressText: "Step {CURRENT} / {TOTAL}",
          endTutorialTitle: "Tutorial Paused",
          endTutorialInstruction: "Great! You have completed the 'Data Import' section.\nYou can reopen this tutorial at any time or continue exploring other features.",
          endTutorialReturnButton: "Return to 'Task 1 Complete' step",
          exampleDashboardSheet_Sales: 'Dashboard | Sales Analysis',
          exampleSalesLogSheet_Sales: 'Source | Sales Log',
          exampleCustomerMasterSheet_Sales: 'Source | Customer List',
          exampleProductTargetsSheet_Sales: 'Source | Product Targets',
          exampleTargetSheet: 'Target Sheet (Manufacturing)',
          exampleImportSourceSheet: 'Source_Import (Manufacturing)',
          exampleVerifySourceSheet: 'Source_Verify (Manufacturing)',
          exampleCompareSourceSheet: 'Source_Compare (Manufacturing)',
          headerQuantity: 'Quantity',
          headerUnitPrice: 'Unit Price',
          headerTotalSales: 'Total Sales',
          headerMonthlyTarget: 'Monthly Target',
          headerAchievementRate: 'Achievement Rate',
          productLaptop: 'High-Performance Laptop',
          headerProductName: 'Product Name',
          businessWelcomeTitle: "Welcome! (Business & Sales)",
          businessWelcomeInstruction: "Welcome to the interactive tutorial for the 'Business & Sales Example'.\n\nOur goal is to transform an incomplete sales record into a complete analysis report using the features of MasterDataAnalyzer.\n\nClick 'Next' to start our first task!",
          businessTask1Step1Title: "Task 1: Data Import (1/6)",
          businessTask1Step1Instruction: "First, please ensure you have activated the <code>{SHEET_NAME}</code> sheet.\n\nOur goal is to filter and import sales records from <code>{SOURCE_SHEET_NAME}</code> into the current dashboard.",
          businessTask1Step2Title: "Task 1: Data Import (2/6)",
          businessTask1Step2Instruction: "Please click on <code>MasterDataAnalyzer > Data Import Tool > ⚙️ Data Import Settings</code> in the top menu.",
          businessTask1Step3Title: "Task 1: Data Import (3/6)",
          businessTask1Step3Instruction: "In the settings window, please configure the following:\n1. **Source Spreadsheet URL**: (Paste the URL of the current file)\n2. **Source Data Sheet Name**: Select <code>{SOURCE_SHEET_NAME}</code>\n3. **Target Data Sheet Name**: Should be auto-filled with <code>{TARGET_SHEET_NAME}</code>",
          businessTask1Step4Title: "Task 1: Data Import (4/6)",
          businessTask1Step4Instruction: "Next, set the data ranges:\n1. **Header Start Row for Data Import**: <code>1</code>\n2. **Data Import Start Row**: <code>2</code>\n3. **Source Data Import Range**: <code>A2:E9</code>",
          businessTask1Step5Title: "Task 1: Data Import (5/6)",
          businessTask1Step5Instruction: "Suppose we only want to analyze sales for the \"{PRODUCT_NAME}\". Please add a filter condition:\n1. **Header Name**: <code>{HEADER_NAME}</code>\n2. **Keywords**: <code>{PRODUCT_NAME}</code>\n\nIn [Source Header Import Range Settings], you can enter a header range to limit the scope of header detection. If left blank, all headers will be used by default.",
          businessTask1Step6Title: "Task 1: Data Import (6/6)",
          businessTask1Step6Instruction: "Great! All settings are complete.\n\nPlease click 'Save Settings', close the window, and then run <code>MasterDataAnalyzer > Data Import Tool > ▶️ Run Import</code> from the menu.",
          businessCheckpointTitle: "Task 1 Complete!",
          businessCheckpointInstruction: "Congratulations! You have successfully imported the raw data into the dashboard.\n\nWould you like to continue learning about the next core feature, 'Data Comparison'?",
          businessTask2Step1Title: "Task 2: Enrich Customer Data (1/3)",
          businessTask2Step1Instruction: "Excellent! The dashboard now has raw data but lacks detailed customer information.\n\nNext, we'll use the 'Data Comparison' feature to look up and fill in customer data from the \"[Source] Customer Master\".\n\nOpen <code>Data Import Tool > ⚙️ Data Comparison Settings</code> and begin configuring Task 2:",
          businessTask2Step2Title: "Task 2: Enrich Customer Data (2/3)",
          businessTask2Step2Instruction: "Please apply the following settings:\n1. **Source Sheets**: Choose \"Source Data Sheet Name\" as <code>{SOURCE_SHEET_NAME}</code> and \"Target Data Sheet Name\" as <code>{TARGET_SHEET_NAME}</code>.\n2. **Ranges**: Set \"Target Data Start Row\" to <code>2</code> and \"Source Data Compare Range\" to <code>A2:D6</code>.\n3. **Field Mapping**: Set up the fields in the following order:\n  - **Target Lookup Column**: <code>B</code> (Customer ID)\n  - **Source Compare Column**: <code>A</code> (Customer ID)\n  - **Source Return Column**: <code>B</code> (Customer Name)\n  - **Target Write Column**: <code>C</code> (Customer Name)",
          businessTask2Step3Title: "Task 2: Enrich Customer Data (3/3)",
          businessTask2Step3Instruction: "After saving the settings, run <code>MasterDataAnalyzer > Data Import Tool > ▶️ Run Data Comparison</code>.\n\nYou will see the customer names have been successfully imported. Repeat this process to also fill in the \"Region\" and \"Salesperson\" to complete the dashboard.",
          businessTask3Step1Title: "Task 3: Compare Sales Targets (1/2)",
          businessTask3Step1Instruction: "The dashboard data is becoming more complete! For the final step, let's compare the sales targets for each product.\n\nPlease open the <code>'Data Comparison Settings'</code> again.",
          businessTask3Step2Title: "Task 3: Compare Sales Targets (2/2)",
          businessTask3Step2Instruction: "This time, use \"Product Name\" as the lookup key and select <code>\"Source | Product Targets\"</code> as the source sheet.<br>Set the \"Source Data Compare Range\" to <code>A2:B5</code>, and configure the field mapping as follows:<br>  - **Target Lookup Column**: <code>F</code> (Product Name)<br>  - **Source Compare Column**: <code>A</code> (Product Name)<br>  - **Source Return Column**: <code>B</code> (Monthly Target)<br>  - **Target Write Column**: <code>J</code> (Monthly Target)",
          businessFinalStepTitle: "Congratulations, Analysis Complete!",
          businessFinalStepInstruction: "Then, run the <code>'Data Comparison'</code> again.\n\nAll data is now in place! You now have a clean, complete dataset ready for analysis.\n\nNext, you can manually enter or click the button below to insert formulas in the corresponding columns to complete the final calculations:\n- **Total Sales**: <code>=H2*I2</code> (Quantity * Unit Price)\n- **Achievement Rate**: <code>=G2/J2</code> (Total Sales / Monthly Target)",
          businessFinalStepButton: "Insert Formulas Automatically",
          businessProcessingButton: "Processing...",
          businessSuccessButton: "Formulas Inserted!",
          businessErrorButton: "Error, please retry",
          mfgWelcomeTitle: "Welcome! (Manufacturing)",
          mfgWelcomeInstruction: "Welcome to the interactive tutorial for the 'Manufacturing Example'.\n\nThis guide will walk you through using MasterDataAnalyzer's 'Data Import (Array Mode)' and 'Data Validation' features to convert a complex bill of materials into a standardized format and compare it against master data.",
          mfgTask1Step1Title: "Task 1: Data Import (1/6)",
          mfgTask1Step1Instruction: "First, please ensure you have activated the <code>{SHEET_NAME}</code> sheet.\n\nOur goal is to convert and import unstructured data from <code>{SOURCE_SHEET_NAME}</code> into the current target sheet.",
          mfgTask1Step2Title: "Task 1: Import Settings (2/6)",
          mfgTask1Step2Instruction: "Please click on <code>MasterDataAnalyzer > Data Import Tool > ⚙️ Data Import Settings</code> in the top menu.",
          mfgTask1Step3Title: "Task 1: Import Settings (3/6)",
          mfgTask1Step3Instruction: "In the settings window, please configure the following:\n1. **Source Spreadsheet URL**: (Paste the URL of the current file)\n2. **Source Data Sheet Name**: Select <code>{SOURCE_SHEET_NAME}</code>\n3. **Target Data Sheet Name**: Should be auto-filled with <code>{TARGET_SHEET_NAME}</code>",
          mfgTask1Step4Title: "Task 1: Import Settings (4/6)",
          mfgTask1Step4Instruction: "Next is the key step, enabling 'Data Array Comparison' mode:\n1. **Header Start Row for Data Import**: <code>3</code>\n2. **Data Import Start Row**: <code>4</code>\n3. **Source Data Import Range**: <code>A9:E14</code>\n4. **Header Start Row for Other Blocks**: <code>F2:I2</code>\n5. **Data Range within Header of Other Blocks**: <code>F9:I14</code>",
          mfgTask1Step5Title: "Task 1: Filter Settings (5/6)",
          mfgTask1Step5Instruction: "In the Filter & Validation section, you can select which headers and keywords to import.\n1. In \"Source Header Import Range Settings\", enter <code>A8:E8</code>\n2. For \"Keyword Filter Conditions (AND)\", select <code>Owner</code>\n3. Click the \"Select\" box to the right of the keyword and choose <code>Mark</code>, <code>Linda</code>, and <code>Amanda</code>.",
          mfgTask1Step6Title: "Task 1: Execute Import (6/6)",
          mfgTask1Step6Instruction: "Once the import settings are configured, don't forget to click \"Save Settings\". After saving, proceed to run <code>▶️ Run Import</code>.\n\nYou should now see imported data in columns A - G, indicating that the data has been successfully filtered and imported.",
          mfgCheckpointTitle: "Task 1 Complete!",
          mfgCheckpointInstruction: "Congratulations! You have successfully transformed and imported the array data.\n\nNext, would you like to learn how to use the 'Data Validation' feature to compare the imported data with master data?",
          mfgTask2Step1Title: "Task 2: Validation Settings (1/4)",
          mfgTask2Step1Instruction: "Great! Now let's verify the accuracy of the imported data.\n\nPlease open <code>MasterDataAnalyzer > Data Validation Tool > ⚙️ Data Validation Settings</code>.",
          mfgTask2Step2Title: "Task 2: Data Range Settings (2/4)",
          mfgTask2Step2Instruction: "In the [Data Validation] settings window, please configure the following:<br>1. **Source Spreadsheet URL**: (Paste the URL of the current file)<br>2. **Source Data Sheet Name**: Select <code>{SOURCE_SHEET_NAME}</code><br>3. Please fill in the data range settings in the following order:<br>   - Data Import Start Row after Validation: <code>4</code><br>   - Target Data Header Start Row: <code>3</code><br>   - Source Data Header Start Row: <code>1</code><br><br>Proceed to the next step to begin setting up the 'Field Validation Conditions' and 'Validation Result Outputs'.",
          mfgTask2Step3Title: "Task 2: Field Validation Settings (3/4)",
          mfgTask2Step3Instruction: "In the Field and Validation Conditions section, the first step is to map the 'Target Column' and 'Source Column'.<br>We recommend using the <b>Auto-map Validation Fields</b> feature first, which will <b>automatically match and recommend</b> suitable headers for you.<br>For this task, the validation conditions are as follows:<br>Target Column - Source Column<br><code>B</code> - <code>B</code><br><code>C</code> - <code>C</code><br><code>D</code> - <code>D</code><br><code>E</code> - <code>E</code><br><code>G</code> - <code>A</code>",
          mfgTask2Step4Title: "Task 2: Validation Output Settings (4/4)",
          mfgTask2Step4Instruction: "Next, configure the data columns to be returned from the source upon successful validation.<br>Similarly, you can use the <b>Auto-map Output Fields</b> feature to speed up the setup.<br>The output field conditions are as follows:<br>Target Column - Source Column<br><code>H</code> - <code>F</code><br><code>I</code> - <code>G</code><br><code>J</code> - <code>H</code><br><br>Finally, we will set the \"Mismatch Info Output Column\" to <code>K</code>, so the script can write error messages there.",
          mfgTask2FinalStepTitle: "Task 2: Execute Validation (5/5)",
          mfgTask2FinalStepInstruction: "Please remember to click [Save Settings], then start by running <code>Data Validation Tool > ▶️ Run Validation (MS Mode)</code>!",
          mfgCheckpoint2Title: "Task 2 Complete!",
          mfgCheckpoint2Instruction: "Congratulations! You have successfully validated the data.\n\nWould you like to proceed to the final task: enriching the data with shipping statuses?",
          mfgTask3Step1Title: "Task 3: Compare Shipping Status (1/3)",
          mfgTask3Step1Instruction: "Excellent! For the final task, we will use 'Data Comparison' to fetch the latest shipping status for each item from the <code>{SOURCE_SHEET_NAME}</code> sheet.",
          mfgTask3Step2Title: "Task 3: Shipping Status Settings (2/3)",
          mfgTask3Step2Instruction: "Please open <code>MasterDataAnalyzer > Data Import Tool > ⚙️ Data Comparison Settings</code> and configure the following:",
          mfgTask3Step3Title: "Task 3: Shipping Status Field Mapping (3/3)",
          mfgTask3Step3Instruction: "Apply these settings to link the tracking number and retrieve the status:<br>- **Source Data Sheet Name**: <code>{SOURCE_SHEET_NAME}</code><br>- **Target Data Start Row**: <code>4</code><br>- **Source Data Compare Range**: <code>A2:B22</code> (covers all data)<br>- **Target Lookup Column**: <code>H</code> (Shipment Tracking Number)<br>- **Source Compare Column**: <code>A</code> (Shipment Tracking Number)<br>- **Source Return Column**: <code>B</code> (Shipment Status)<br>- **Target Write Column**: <code>L</code> (Shipment Status)",
          mfgFinalStepTitle: "All Tasks Complete!",
          mfgFinalStepInstruction: "After saving the settings, please run <code>▶️ Run Data Comparison</code> from the menu.<br><br>Check column L in the <code>{SHEET_NAME}</code> sheet. You have now successfully integrated production, validation, and shipping data!",
        },
        'zh_TW': {
          businessGuideTutorialTitle: '業務統計範例 - 互動教學',
          manufacturingGuideTutorialTitle: '生產製造範例 - 互動教學',
          tutorialNext: "下一步",
          tutorialPrev: "上一步",
          tutorialClose: "關閉",
          tutorialYes: "是，繼續教學",
          tutorialNo: "否，稍後決定",
          tutorialProgressText: "步驟 {CURRENT} / {TOTAL}",
          endTutorialTitle: "教學暫告一段落",
          endTutorialInstruction: "很棒！您已完成「資料匯入」的學習。\n\n您可以隨時重新打開此教學，或繼續探索其他功能。",
          endTutorialReturnButton: "返回「任務一完成」步驟",
          exampleDashboardSheet_Sales: '月報分析儀表板 (業務統計)',
          exampleSalesLogSheet_Sales: '[來源] 銷售紀錄',
          exampleCustomerMasterSheet_Sales: '[來源] 客戶清單',
          exampleProductTargetsSheet_Sales: '[來源] 產品目標',
          exampleTargetSheet: '目標資料表 (生產製造)',
          exampleImportSourceSheet: '來源_匯入 (生產製造)',
          exampleVerifySourceSheet: '來源_驗證 (生產製造)',
          exampleCompareSourceSheet: '來源_比對 (生產製造)',
          headerQuantity: '數量',
          headerUnitPrice: '單價',
          headerTotalSales: '總銷售額',
          headerMonthlyTarget: '目標月銷售額',
          headerAchievementRate: '業績達成率',
          productLaptop: '高效能筆電',
          headerProductName: '產品名稱',
          businessWelcomeTitle: "歡迎！ (業務統計)",
          businessWelcomeInstruction: "歡迎來到「業務統計範例」的互動教學。\n\n我們的目標是將一份不完整的銷售紀錄，透過 MasterDataAnalyzer 的功能，變成一份完整的分析報表。\n\n點擊「下一步」開始我們的第一個任務！",
          businessTask1Step1Title: "任務一：資料匯入 (1/6)",
          businessTask1Step1Instruction: "首先，請確認您已啟用 <code>{SHEET_NAME}</code> 分頁。\n\n我們的目標是將 <code>{SOURCE_SHEET_NAME}</code> 的銷售紀錄，篩選後匯入到目前的儀表板中。",
          businessTask1Step2Title: "任務一：資料匯入 (2/6)",
          businessTask1Step2Instruction: "請點擊頂端選單的 <code>MasterDataAnalyzer > 資料匯入工具 > ⚙️ 資料匯入設定</code>。",
          businessTask1Step3Title: "任務一：資料匯入 (3/6)",
          businessTask1Step3Instruction: "在設定視窗中，請進行以下設定：\n1. **來源資料表 URL**: (貼上當前檔案的網址)\n2. **來源資料分頁名稱**: 選擇 <code>{SOURCE_SHEET_NAME}</code>\n3. **目標資料表分頁名稱**: 應會自動帶入 <code>{TARGET_SHEET_NAME}</code>",
          businessTask1Step4Title: "任務一：資料匯入 (4/6)",
          businessTask1Step4Instruction: "接著，設定資料範圍：\n1. **資料匯入的標頭之起始列數**: <code>1</code>\n2. **資料匯入的起始列數**: <code>2</code>\n3. **來源資料匯入範圍**: <code>A2:E9</code>",
          businessTask1Step5Title: "任務一：資料匯入 (5/6)",
          businessTask1Step5Instruction: "假設我們只想分析「{PRODUCT_NAME}」的業績，請新增一筆篩選條件：\n1. **標頭名稱**: <code>{HEADER_NAME}</code>\n2. **關鍵字**: <code>{PRODUCT_NAME}</code>\n\n此功能 [來源標頭匯入範圍設定] 您可以輸入標頭範圍來限定標頭的檢測範圍，若無輸入則會以預設的所有標頭做為檢測範圍。",
          businessTask1Step6Title: "任務一：資料匯入 (6/6)",
          businessTask1Step6Instruction: "太棒了！所有設定都已完成。\n\n請點擊「儲存設定」並關閉視窗，然後從選單執行 <code>MasterDataAnalyzer > 資料匯入工具 > ▶️ 資料匯入</code>。",
          businessCheckpointTitle: "任務一完成！",
          businessCheckpointInstruction: "恭喜您！您已成功將原始數據匯入儀表板。\n\n您想繼續學習下一個核心功能「資料比對」嗎？",
          businessTask2Step1Title: "任務二：豐富客戶資料 (1/3)",
          businessTask2Step1Instruction: "太棒了！現在儀表板有了原始數據，但還缺少客戶的詳細資訊。\n\n接下來，我們將使用「資料比對」功能，從 \"[來源] 客戶主檔\" 中查找並填入客戶資料。\n\n打開 <code>資料匯入工具 > ⚙️ 資料比對設定</code>，並開始進行任務二的設定：",
          businessTask2Step2Title: "任務二：豐富客戶資料 (2/3)",
          businessTask2Step2Instruction: "請進行以下設定：\n1. **來源分頁**: 請選擇 \"來源資料分頁名稱\" 為 <code>{SOURCE_SHEET_NAME}</code>，而 \"目標資料表分頁名稱\" 為 <code>{TARGET_SHEET_NAME}</code>。\n2. \"目標資料起始列數\" 設定為 <code>2</code>，\"來源資料比對範圍\" 設定為 <code>A2:D6</code>。\n3. \"資料比對欄位對應\" 請依照下方順序逐一設定：\n  - **目標查找欄位**: <code>B</code> (客戶ID)\n  - **來源比對欄位**: <code>A</code> (客戶ID)\n  - **來源返回欄位**: <code>B</code> (客戶全名)\n  - **目標寫入欄位**: <code>C</code> (客戶全名)",
          businessTask2Step3Title: "任務二：豐富客戶資料 (3/3)",
          businessTask2Step3Instruction: "設定儲存後，請執行 <code>MasterDataAnalyzer > 資料匯入工具 > ▶️ 執行資料比對</code>。\n\n這時您將會看到客戶全名已被成功匯入。請重複此操作，將 \"所屬地區\" 和 \"負責業務員\" 也一併填入，您將可以得到完整的業務統計儀表板囉。",
          businessTask3Step1Title: "任務三：比對銷售目標 (1/2)",
          businessTask3Step1Instruction: "儀表板的資料越來越完整了！最後一步，讓我們來比對每項產品的銷售目標。\n\n請再次打開<code>「資料比對設定」</code>。",
          businessTask3Step2Title: "任務三：比對銷售目標 (2/2)",
          businessTask3Step2Instruction: "這次，請用 \"產品名稱\" 作為目標對比基準，並選擇 \"來源資料分頁名稱\" 為 <code>{SOURCE_SHEET_NAME}</code>。<br>\"來源資料比對範圍\" 請設定為 <code>A2:B5</code>，\"資料比對欄位對應\" 請依照下方順序逐一設定：<br>  - **目標查找欄位**: <code>F</code> (產品名稱)<br>  - **來源比對欄位**: <code>A</code> (產品名稱)<br>  - **來源返回欄位**: <code>B</code> (目標月銷售額)<br>  - **目標寫入欄位**: <code>J</code> (目標月銷售額)",
          businessFinalStepTitle: "恭喜您，分析完成！",
          businessFinalStepInstruction: "接著，請再次執行<code>「資料比對」</code>。\n\n所有資料都已到位！您現在擁有了一份乾淨、完整的數據。\n\n接下來，您可以手動或點擊下方按鈕，在對應的欄位中填入公式來完成最後的計算：\n- **總銷售額**: <code>=H2*I2</code> (數量 * 單價)\n- **業績達成率**: <code>=G2/J2</code> (總銷售額 / 目標月銷售額)",
          businessFinalStepButton: "自動帶入公式",
          businessProcessingButton: "處理中...",
          businessSuccessButton: "公式已成功帶入！",
          businessErrorButton: "發生錯誤，請重試",
          mfgWelcomeTitle: "歡迎！ (生產製造)",
          mfgWelcomeInstruction: "歡迎來到「生產製造範例」的互動教學。\n\n此教學將引導您使用 MasterDataAnalyzer 的「資料匯入 (陣列模式)」與「資料驗證」功能，將複雜的物料清單轉換為標準化格式，並與主資料進行比對。",
          mfgTask1Step1Title: "任務一：資料匯入 (1/6)",
          mfgTask1Step1Instruction: "首先，請確認您已啟用 <code>{SHEET_NAME}</code> 分頁。\n\n我們的目標是將 <code>{SOURCE_SHEET_NAME}</code> 中非結構化的資料，轉換並匯入到目前的目標工作表中。",
          mfgTask1Step2Title: "任務一：資料匯入設定 (2/6)",
          mfgTask1Step2Instruction: "請點擊頂端選單的 <code>MasterDataAnalyzer > 資料匯入工具 > ⚙️ 資料匯入設定</code>。",
          mfgTask1Step3Title: "任務一：資料匯入設定 (3/6)",
          mfgTask1Step3Instruction: "在設定視窗中，請進行以下設定：\n1. **來源資料表 URL**: (貼上當前檔案的網址)\n2. **來源資料分頁名稱**: 選擇 <code>{SOURCE_SHEET_NAME}</code>\n3. **目標資料表分頁名稱**: 應會自動帶入 <code>{TARGET_SHEET_NAME}</code>",
          mfgTask1Step4Title: "任務一：資料匯入設定 (4/6)",
          mfgTask1Step4Instruction: "接下來是關鍵步驟，啟用「資料陣列比對」模式：\n1. **資料匯入的標頭之起始列數**: <code>3</code>\n2. **資料匯入的起始列數**: <code>4</code>\n3. **來源資料匯入範圍**: <code>A9:E14</code>\n4. **其他區塊的標頭的起始列數**: <code>F2:I2</code>\n5. **其他區塊的標頭內的數據範圍**: <code>F9:I14</code>",
          mfgTask1Step5Title: "任務一：資料篩選設定 (5/6)",
          mfgTask1Step5Instruction: "在篩選與驗證功能區塊裡，您可以自行挑選想匯入的標頭與關鍵字篩選。\n1. 在 \"來源標頭匯入範圍設定\" 輸入 <code>A8:E8</code>\n2. 在 \"關鍵字篩選條件 (AND)\" 選擇 <code>Owner</code>\n3. 點擊關鍵字右側 \"選取\" 方塊選擇 <code>Mark</code>、<code>Linda</code>、<code>Amanda</code>",
          mfgTask1Step6Title: "任務一：執行資料匯入 (6/6)",
          mfgTask1Step6Instruction: "當資料匯入設定完成後請別忘記點擊 \"儲存設定\"。設定與儲存完成後，請繼續執行 <code>▶️ 資料匯入</code>。\n\n這時您應會看到 A - G 欄的匯入資料，代表著資料已成功篩選並匯入囉。",
          mfgCheckpointTitle: "任務一完成！",
          mfgCheckpointInstruction: "恭喜！您已成功將陣列資料轉換並匯入。\n\n接下來，您想學習如何使用「資料驗證」功能，來比對匯入的資料與主資料的差異嗎？",
          mfgTask2Step1Title: "任務二：資料驗證設定 (1/4)",
          mfgTask2Step1Instruction: "很好！現在我們來驗證匯入資料的正確性。\n\n請打開 <code>MasterDataAnalyzer > 資料驗證工具 > ⚙️ 資料驗證設定</code>。",
          mfgTask2Step2Title: "任務二：資料範圍設定 (2/4)",
          mfgTask2Step2Instruction: "在[資料驗證]設定視窗中，請進行以下設定：<br>1. **來源資料表 URL**: (貼上當前檔案的網址)<br>2. **來源資料分頁名稱**: 選擇 <code>{SOURCE_SHEET_NAME}</code><br>3. 請依照下列設定順序填妥資料範圍設定：<br>   - 資料驗證後匯入的起始列數請填入 <code>4</code><br>   - 目標資料標頭起始列請填入 <code>3</code><br>   - 來源資料標頭起始列請填入 <code>1</code><br><br>接下來繼續下一步，將開始設定 \"欄位驗證條件\" 與 \"驗證結果輸出\"。",
          mfgTask2Step3Title: "任務二：欄位驗證設定 (3/4)",
          mfgTask2Step3Instruction: "在欄位與驗證條件功能區塊裡，首先要設定的是 \"目標工作表欄位\" 與 \"來源工作表欄位\" 的對應關係。<br>建議您可以優先使用 <b>自動對應驗證欄位</b> 功能，此功能將會 <b>自動匹配與推薦</b> 合適的標頭給您設定。<br>在此項任務裡，欄位驗證要設定的條件如下：<br>目標工作表欄位 - 來源工作表欄位<br><code>B</code> - <code>B</code><br><code>C</code> - <code>C</code><br><code>D</code> - <code>D</code><br><code>E</code> - <code>E</code><br><code>G</code> - <code>A</code>",
          mfgTask2Step4Title: "任務二：驗證結果輸出設定 (4/4)",
          mfgTask2Step4Instruction: "接著設定當驗證成功時，要從來源回填的資料欄位。<br>同樣地，您也可以使用 <b>自動對應輸出欄位</b> 功能來加速設定。<br>驗證結果輸出要設定的欄位條件如下：<br>目標工作表欄位 - 來源工作表欄位<br><code>H</code> - <code>F</code><br><code>I</code> - <code>G</code><br><code>J</code> - <code>H</code><br><br>最後，我們將 \"不吻合資訊輸出\" 的欄位設為 <code>K</code>，以便腳本將錯誤訊息寫入此處。",
          mfgTask2FinalStepTitle: "任務二：執行驗證 (5/5)",
          mfgTask2FinalStepInstruction: "請記得點擊 [儲存設定]，然後開始執行 <code>資料驗證工具 > ▶️ 執行驗證 (MS 累加項模式)</code>！",
          mfgCheckpoint2Title: "任務二完成！",
          mfgCheckpoint2Instruction: "恭喜！您已成功驗證資料。\n\n您想繼續最後一個任務：用貨運狀態來豐富資料嗎？",
          mfgTask3Step1Title: "任務三：比對貨運狀態 (1/3)",
          mfgTask3Step1Instruction: "太棒了！最後一個任務，我們將使用「資料比對」功能，從 <code>{SOURCE_SHEET_NAME}</code> 工作表中，取得每個項目的最新貨運狀態。",
          mfgTask3Step2Title: "任務三：比對貨運狀態設定 (2/3)",
          mfgTask3Step2Instruction: "請打開 <code>MasterDataAnalyzer > 資料匯入工具 > ⚙️ 資料比對設定</code>，並進行以下設定：",
          mfgTask3Step3Title: "任務三：貨運狀態欄位對應 (3/3)",
          mfgTask3Step3Instruction: "請套用以下設定，來關聯追蹤號碼並取得狀態：<br>- **來源資料分頁名稱**: <code>{SOURCE_SHEET_NAME}</code><br>- **目標資料起始列數**: <code>4</code><br>- **來源資料比對範圍**: <code>A2:B22</code> (涵蓋所有資料)<br>- **目標查找欄位**: <code>H</code> (Shipment Tracking Number)<br>- **來源比對欄位**: <code>A</code> (Shipment Tracking Number)<br>- **來源返回欄位**: <code>B</code> (Shipment Status)<br>- **目標寫入欄位**: <code>L</code> (Shipment Status)",
          mfgFinalStepTitle: "所有任務完成！",
          mfgFinalStepInstruction: "儲存設定後，請從選單執行 <code>▶️ 執行資料比對</code>！<br><br>請檢查 <code>{SHEET_NAME}</code> 工作表的 L 欄。您現在已成功整合了生產、驗證與貨運的資料！",
        }
    };
    return allTranslations[locale] || allTranslations.en;
}


function showTutorialSidebar(tutorialType) {
  const T = getTutorialTranslations();
  let title = '';
  let htmlFile = 'GuidePageTraining.html';

  if (tutorialType === 'business') {
    generateBusinessExample(); 
    title = T.businessGuideTutorialTitle;
  } else if (tutorialType === 'manufacturing') {
    generateManufacturingExample();
    title = T.manufacturingGuideTutorialTitle;
  } else {
    throw new Error('Unknown tutorial type specified.');
  }

  const htmlTemplate = HtmlService.createTemplateFromFile(htmlFile);
  htmlTemplate.tutorialType = tutorialType;
  htmlTemplate.translations = JSON.stringify(T); 

  const htmlOutput = htmlTemplate.evaluate().setTitle(title);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function startBusinessTutorial() {
  showTutorialSidebar('business');
}

function startManufacturingTutorial() {
  showTutorialSidebar('manufacturing');
}

function tutorial_highlightSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.activate();
    }
  } catch (e) {
    Logger.log(`Could not highlight sheet "${sheetName}": ${e.message}`);
  }
}

function tutorial_highlightRange(sheetName, rangeA1) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.activate();
      const range = sheet.getRange(rangeA1);
      range.activate();
      const originalColors = range.getBackgrounds();
      range.setBackground("#fffde7"); 
      Utilities.sleep(3000); 
      range.setBackgrounds(originalColors);
    }
  } catch (e) {
    Logger.log(`Could not highlight range "${rangeA1}" on sheet "${sheetName}": ${e.message}`);
  }
}

function findColumnIndexByHeader(headers, headerName) {
  if (!headers || !Array.isArray(headers)) return -1;
  const index = headers.findIndex(h => h.toString().trim() === (headerName || '').toString().trim());
  return index !== -1 ? index + 1 : -1;
}

function insertSalesMetricsFormulas() {
  const T = getTutorialTranslations();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(T.exampleDashboardSheet_Sales);

  if (!targetSheet) {
    throw new Error(`Could not find the example sheet: "${T.exampleDashboardSheet_Sales}"`);
  }

  const lastRow = targetSheet.getLastRow();
  const headerRow = 1; 
  const startRow = 2;

  if (lastRow < startRow) {
    return; 
  }
  const numRows = lastRow - startRow + 1;
  const headers = targetSheet.getRange(headerRow, 1, 1, targetSheet.getMaxColumns()).getValues()[0];

  const qtyCol = findColumnIndexByHeader(headers, T.headerQuantity);
  const unitPriceCol = findColumnIndexByHeader(headers, T.headerUnitPrice);
  const totalSalesCol = findColumnIndexByHeader(headers, T.headerTotalSales);
  const monthlyTargetCol = findColumnIndexByHeader(headers, T.headerMonthlyTarget);
  const achievementRateCol = findColumnIndexByHeader(headers, T.headerAchievementRate);

  if (totalSalesCol > 0 && qtyCol > 0 && unitPriceCol > 0) {
    const totalSalesRange = targetSheet.getRange(startRow, totalSalesCol, numRows, 1);
    const formulaR1C1 = `=RC[${qtyCol - totalSalesCol}]*RC[${unitPriceCol - totalSalesCol}]`;
    totalSalesRange.setFormulaR1C1(formulaR1C1);
  } else {
    Logger.log("Could not find all required columns for Total Sales calculation.");
  }

  if (achievementRateCol > 0 && totalSalesCol > 0 && monthlyTargetCol > 0) {
    const achievementRateRange = targetSheet.getRange(startRow, achievementRateCol, numRows, 1);
    const formulaR1C1 = `=IF(RC[${monthlyTargetCol - achievementRateCol}]=0, 0, RC[${totalSalesCol - achievementRateCol}]/RC[${monthlyTargetCol - achievementRateCol}])`;
    achievementRateRange.setFormulaR1C1(formulaR1C1);
    achievementRateRange.setNumberFormat("0.00%");
  } else {
     Logger.log("Could not find all required columns for Achievement Rate calculation.");
  }
}
