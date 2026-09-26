# 急診護理排班決策支援系統

研究所申請用 Python 專案。保留既有混合整數規劃排班核心，新增可攜入口、匿名示範、本人意願優先的休假帳本、測試與研究說明。這是研究原型，不是院方核准的排班或薪資系統。

## 在 VS Code 執行

建議使用 Python 3.13（本次實測版本）。用 VS Code「開啟資料夾」開啟本專案，不要只開單一 Python 檔。終端機執行：

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py --demo --leave-input examples/leave_balances.json
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

不必啟用 PowerShell 腳本執行權限。若用其他 Python 版本，先確認套件相容性。選取 `.venv` 解譯器後亦可按 F5 執行匿名範例。

成功時在 `results/執行時間_識別碼/` 產生 `研究展示班表.xlsx` 與 `report.json`。失敗只產生診斷；每次使用新資料夾，避免把前次舊班表誤當成功結果。展示人員完全虛構，無 Google 帳號也能執行。

## 專案組成

| 檔案 | 功能 |
| --- | --- |
| main.py | 示範／本機 JSON 輸入、前置檢查、求解、驗證與輸出入口 |
| scheduler_core.py | 既有完整排班模型、資料解析、帶教／區域約束、驗證及原匯出函式；機構識別資料已移除 |
| leave_accounting.py | 四週週期定位、既有餘額、加班選擇、補休與特休詢問／扣抵 |
| workbook_export.py | 純 Python 研究班表輸出、D/E/N 分頁、津貼、詢問分頁 |
| tests/test_portfolio.py | 休假同意、餘額、跨月輸入及安全檢查 |
| RESEARCH_STATEMENT.md | 研究動機、方法、驗證與限制，供申請者依實際經歷調整 |

## 使用自己的整理資料

```powershell
python main.py --write-demo-input examples/my_input.json
python main.py --input examples/my_input.json --leave-input examples/leave_balances.json
```

範例 JSON 會列出完整欄位。修改前先另存私人版本，切勿把真實人員資料提交至公開專案。

- `nurses`：唯一員編、姓名、到職日、原班 D/E/N、資格 `areas`、`max_hours`（正常工時門檻）、部分工時 `part_time`、已驗證帶教累積 `training_completed_seed`。
- `requests`：員編、日期、`OFF_LOCK`／`ANNUAL_LEAVE`／`MUST_SHIFT`／`PREFERRED_SHIFT`／`AVOID_SHIFT`，指定班別另填 `shift`。原始特休保持硬條件，不自動刪除。
- `boundary`：每人上月最後 14 天，明列班別。行政班使用 `8-5` 並可附 `start_hour`、`end_hour`，仍算出勤。不可把未知資料當 OFF。
- `official_holidays`：明確提供院方確認之日期，展示用空陣列不是正式年度日曆。
- `max_cross_shift_days`、`max_overtime_shifts`：預設每人 5 天、1 班；需嚴格少於 5 天時請設 4。
- `training_plan`：每筆 `nurse_id/day/shift/role`，流動 `Observation1`、留觀 `Observation2`。每班最多 2 人。`preferred_preceptors` 以學生員編對應優先老師員編。老師由 `areas` 含 Teaching 人員篩選，核心仍需符合實際區域資格。
- 所有累積補休留在休假帳本，不能直接塞入 `prior_unused_hours` 降低當月目標。

本機入口不自動連線 Google，也不讀取舊 pickle 快取。既有唯讀 Google 解析／載入函式保留於核心，整合機構來源時可另裝 `requirements-google.txt`，由管理者提供來源 ID、核定名單及憑證。公開版不含來源 ID、金鑰、真人資料或既有月份快取；LIVE 整合未於此公開版驗收，不能假稱已連上院方最新資料。

## 休假帳本

`examples/leave_balances.json` 是獨立的虛構數值範例，不是匿名化真實員工餘額。可用時數精度為 0.01 小時，排班核心仍以完整 8 小時班求解。

- `annual_opening/comp_opening`：先前累積餘額；未知填 null，不填 0 代替。
- 月加班由求解結果自動計算。`overtime_choice` 為未選擇、領加班費或轉補休；只有轉補休增加試算時數。未實際取得的加班不能拿來支付本次補休。
- `approved_annual_hours` 是原始特休之外已核准額外使用的小時；原始 `ANNUAL_LEAVE` 每天 8 小時由入口自動加計，請勿重複填入。`approved_comp_hours` 是既有已同意且核定的使用量。
- `cycle_verified`、`regular_rest_remaining`、`additional_leave_hours` 由排班管理者在完整四週（例如 9/10～10/7）核對後提供，不由當月 OFF 數猜測。一般休假仍有餘額時不扣補休或特休。
- 補休需 `comp_consent=同意`；不足時要求 `annual_asked=true` 且 `annual_consent=同意` 才扣新增特休。拒絕或餘額不足列警示，不擅自變更排班。
- `opening_verified/entries_verified/overtime_earned` 控制正式餘額。Excel 結算是本次執行的快照，修改本人意願後更新 JSON 再執行，不會從 Excel 的手動修改自動回寫。

## 目前邊界

1. 示範使用固定三上三休、充分資格人力，測的是可執行性與規則管線，不是臨床人力成效；不能由示範推論真實月份一定外援零、加班零。
2. 原核心保留年資 ≥2.5 月或既有帶教累積之獨立判定。實際臨床獨立資格仍須主管確認，不能靠年資自動認證專業能力。
3. 四例四休、國定假日替代與跨月已核准休假需要完整週期明細與院方規範；目前帳本為核定後計算及詢問流程，並未實作全自動合法性判定。
4. 條件式 Plan C 最少使用天數、外援歸零等舊情境需要經審核的具體人力及契約設定；本機入口不將某人改列部分工時或任意調低目標。舊快取情境不隨包附送，也不假稱已重現。
5. LOCAL_REVIEW 會保留 Google 來源驗證錯誤，僅區分排班約束通過與否，永遠 `clinical_publishable=false`。求解逾時並不等於數學無解。
6. 測試不涵蓋所有既有數千行核心分支，尤其真實跨月、複合帶教、可移動預假與不同院制；使用前需再做回歸驗證。

## 作者與工具揭露

請在申請前填上本人姓名、實際職務、需求定義與驗證貢獻。此專案在既有排班程式基礎上使用 AI 協助整理與新增程式；不得聲稱未曾使用 AI、已通過臨床試驗、已發表研究或已取得尚不存在的成效。參照申請學校的 AI／作品規範揭露實際協作方式。

