# 11508 活動介入初次評估自動化工具

這是一個基於 Google Apps Script (GAS) 開發的自動化腳本。當使用者提交 Google 表單時，系統會自動抓取填答內容、複製指定的 Google 評估範本（Google Doc），並自動填入個案基本資料、對應診斷、重視問題、生理功能、輔具、長短期目標以及介入項目，最後自動重新命名並存入指定的 Google 雲端硬碟資料夾中。

---

## 🚀 主要功能特色

1. **表單自動觸發 (`onFormSubmitAutoFillNew`)**：與 Google 表單連動，表單送出後即在背景自動產出評估文件。
2. **智慧資料夾防呆**：若指定的資料夾 ID 無效，會自動改存至範本同層目錄或根目錄，避免程式因路徑錯誤而中斷。
3. **自動文字清洗 (`cleanItemText`)**：自動清除介入項目中的「團：」或「個：」前綴，並與 11508 標準項目進行精準的白名單比對。
4. **動態核取方塊替換 (`replaceCheckmark`)**：自動將表單選項對應至文件中的核取方塊（將 `□` 轉換為 `■`）。

---

## ⚙️ 系統設定與對應表

程式碼開頭包含以下重要常數與對應設定：

* **基礎設定**
  * `TEMPLATE_DOC_ID`：Google Doc 評估範本 ID
  * `SPECIFIED_FOLDER_ID`：產出檔案儲存的 Google 雲端硬碟資料夾 ID

* **對應清單與對照表**
  * `DIAGNOSIS_MAP`：一、診斷對應表（`d1` ~ `d13`，如痛風、高血壓、失智症等）
  * `PROBLEM_MAP`：二、重視問題對應表（`q1` ~ `q20`）
  * `SINGLE_CHOICE_MAP`：三、生理功能單選對應表（`s1` ~ `s28`）
  * `AIDS_EXACT_MAP`：四、11508 標準輔具對應表
  * `MAIN_PROBLEMS`：五、主要問題文字清單
  * `SHORT_GOALS` / `LONG_GOALS`：六、短、長期目標文字清單
  * `INTERVENTION_ITEMS`：七、11508 介入項目標準清單

---

## 💻 安裝與部署步驟

1. **準備 Google 檔案**：
   * 建立一份 Google Doc 作為評估單範本，並確保內文有對應的文字欄位與核取方塊（如 `案號：`、`姓名：`、`□ 男` 等）。
   * 建立一份對應的 Google 表單。
2. **部署指令碼**：
   * 進入 Google 表單的擴充功能，打開 **Apps Script** 編輯器。
   * 將你的 `.gs` 程式碼貼入。
   * 填入你的 `TEMPLATE_DOC_ID` 與 `SPECIFIED_FOLDER_ID`。
3. **設定觸發條件**：
   * 在 Apps Script 左側點選「觸發條件」（時鐘圖示）。
   * 新增觸發條件：選擇 `onFormSubmitAutoFillNew` 函式，事件來源選擇「**來自表單**」，事件類型選擇「**提交表單時**」。
