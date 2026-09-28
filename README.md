# 急診護理排班決策支援系統

作者：胡修睿（GitHub：HuHsiuJui）。

這個專案以 Python 與混合整數規劃處理急診護理排班，將人員資格、預假、跨月出勤與帶教安排放在同一個模型中計算。求解後會檢查班表，輸出 Excel 與 JSON 報告，並用獨立帳本處理補休、特休與本人意願。目前以虛構資料驗證，作為研究所申請的實作作品，尚未用於正式排班。

## 在 VS Code 執行

測試環境為 Python 3.13。用 VS Code 開啟專案資料夾後，在終端機執行：

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py --demo --leave-input examples/leave_balances.json
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

不必啟用 PowerShell 腳本執行權限。若用其他 Python 版本，先確認套件相容性。入口必須擇一指定 `--demo` 或 `--input`；在 VS Code 偵錯時也須傳入對應參數。

成功時會在 `results/執行時間_識別碼/` 產生 `研究展示班表.xlsx` 與 `report.json`。每次執行使用獨立資料夾，失敗時留下診斷資訊。示範中的 38 名人員均為虛構，不需要 Google 帳號。

## 專案組成

| 檔案 | 功能 |
| --- | --- |
| main.py | 示範／本機 JSON 輸入、前置檢查、求解、驗證與輸出入口 |
| scheduler_core.py | 排班模型、資料解析、帶教與區域約束、結果驗證及既有匯出函式 |
| leave_accounting.py | 四週週期定位、既有餘額、加班選擇、補休與特休詢問／扣抵 |
| workbook_export.py | 純 Python 研究班表輸出、D/E/N 分頁、津貼、詢問分頁 |
| tests/test_portfolio.py | 休假同意、餘額、跨月輸入及安全檢查 |
| RESEARCH_STATEMENT.md | 研究動機、方法、已實作成果、AI 協作與研究限制 |
| TEST_RESULTS.md | 實際執行環境、測試結果與驗證範圍 |

## 使用自己的整理資料

```powershell
python main.py --write-demo-input private/my_input.json
python main.py --input private/my_input.json
```

第一行會建立可修改的虛構範例。換成自己的資料後，將 `demo` 改為 `false`；有對應的休假帳本時，再加上 `--leave-input private/leave_balances.json`。未提供帳本的人員會保留空白餘額，空白不代表零。`private/` 已由 `.gitignore` 排除，真實名單與餘額應留在此目錄，不隨專案提交。

- `nurses`：唯一員編、姓名、到職日、原班 D/E/N、資格 `areas`、`max_hours`（正常工時門檻）、部分工時 `part_time`、已驗證帶教累積 `training_completed_seed`。
- `requests`：員編、日期、`OFF_LOCK`／`ANNUAL_LEAVE`／`MUST_SHIFT`／`PREFERRED_SHIFT`／`AVOID_SHIFT`，指定班別另填 `shift`。原始特休保持硬條件，不自動刪除。
- `boundary`：每人上月最後 14 天，明列班別。行政班使用 `8-5` 並可附 `start_hour`、`end_hour`，仍算出勤。不可把未知資料當 OFF。
- `official_holidays`：明確提供院方確認之日期，展示用空陣列不是正式年度日曆。
- `max_cross_shift_days`、`max_overtime_shifts`：預設每人 5 天、1 班；需嚴格少於 5 天時請設 4。
- `training_plan`：每筆 `nurse_id/day/shift/role`，流動 `Observation1`、留觀 `Observation2`。每班最多 2 人。`preferred_preceptors` 以學生員編對應優先老師員編。老師由 `areas` 含 Teaching 人員篩選，核心仍需符合實際區域資格。
- 累積補休記錄於休假帳本；本機入口不接受以 `prior_unused_hours` 抵減當月工時目標。

目前入口只讀取本機 JSON，不連線 Google，也不讀取舊 pickle 快取。核心保留的 Google 載入函式需另行整合，依賴 `gspread` 與 `google-auth`；安裝套件後仍須配置機構來源與憑證。專案不附私人名單、憑證或既有月份快取，空白的機構設定表示連線預設停用，不影響示範執行。

## 休假帳本

`examples/leave_balances.json` 是獨立的虛構數值範例，不是匿名化真實員工餘額。可用時數精度為 0.01 小時，排班核心仍以完整 8 小時班求解。

- `annual_opening/comp_opening`：先前累積餘額；未知填 null，不填 0 代替。
- 月加班由求解結果自動計算。`overtime_choice` 為未選擇、領加班費或轉補休；只有轉補休增加試算時數。未實際取得的加班不能拿來支付本次補休。
- `approved_annual_hours` 是原始特休之外已核准額外使用的小時；原始 `ANNUAL_LEAVE` 每天 8 小時由入口自動加計，請勿重複填入。`approved_comp_hours` 是既有已同意且核定的使用量。
- `cycle_verified`、`regular_rest_remaining`、`additional_leave_hours` 由排班管理者在完整四週（例如 9/10～10/7）核對後提供，不由當月 OFF 數猜測。一般休假仍有餘額時不扣補休或特休。
- 補休需 `comp_consent=同意`；不足時要求 `annual_asked=true` 且 `annual_consent=同意` 才扣新增特休。拒絕或餘額不足列警示，不擅自變更排班。
- `opening_verified/entries_verified/overtime_earned` 控制正式餘額。Excel 結算是本次執行的快照，修改本人意願後更新 JSON 再執行，不會從 Excel 的手動修改自動回寫。

帳本的「正式」欄位依輸入核定旗標與檢查結果產生，程式不會向院方查證。範例保留「未選擇／未確認」，用來確認未取得同意時不會扣假。Excel 的夜班津貼採固定費率試算，不包含正式加班費與薪資核定。

## 適用範圍

1. 示範採固定三上三休，且每位人員具備足夠資格，主要檢查程式流程。示範的零跨班、零加班不能推論至真實月份。
2. 模型以年資 ≥2.5 月或 18 班帶教累積判定獨立資格；實際任用資格由主管確認。
3. 四例四休、國定假日替代與跨月已核准休假需要完整週期明細與院方規範。`cycle_for` 可定位 28 天週期（預設錨點 2026-09-10，可由函式參數修改），帳本採外部核定後計算及詢問流程，尚無完整週期假別自動分類或合法性判定；當月排班的前 14 天邊界不等於完整四週台帳。
4. 核心保留條件式 Plan C、外援與可移動預假等進階設定，但本機入口未提供這些情境的操作流程，亦未附過去的月份資料。程式不會自行改列部分工時或降低工時目標。
5. DEMO 與 LOCAL_REVIEW 報告分別列出排班檢查與 Google 來源檢查結果，兩種模式皆為 `clinical_publishable=false`。求解逾時表示未在時限內完成，不能據此判定無解。
6. 現有測試涵蓋指定案例，未涵蓋所有核心分支。真實跨月資料、複合帶教、可移動預假與不同院制仍需另外驗證。
7. 核心已有規則式缺口檢查與建議；本機入口失敗時輸出原因及一般檢查建議，尚無完整的最小衝突集合分析或自動情境比較。

## 開發說明

本專案延續既有排班程式，開發與修訂過程使用 AI 協助整理程式、新增功能、建立測試及編修文件。修改可由 Git 紀錄追溯，實際執行結果見 [TEST_RESULTS.md](TEST_RESULTS.md)。
