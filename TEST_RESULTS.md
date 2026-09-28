# 驗證紀錄

日期：2026-09-28（Asia/Taipei）

環境：Windows 11、Python 3.13.7、NumPy 2.3.3、SciPy 1.18.0、openpyxl 3.1.5。

本次從 `7f11f7ffbf1e7e7baedb2a6aa3b366bc56370127` 修訂後重跑，更新原 2026-09-14 的紀錄。測試使用虛構資料；Google 相關測試以 mock 模擬，不連線至真實機構。

## 已執行

- `python -m unittest discover -s tests -v`：30 項測試全部通過（原 25 項，加上 5 項來源／診斷回歸測試）。
- `python main.py --demo --leave-input examples/leave_balances.json`：成功輸出研究展示 Excel 與 JSON 報告。
- `python -m compileall -q main.py scheduler_core.py leave_accounting.py workbook_export.py tests`：通過。
- Excel 匯出後回讀，逐人逐日比對班別，與求解結果一致。

## 虛構資料示範結果

38 名虛構人員，30 天，共 570 筆出勤配置。示範模型 16,312 個變數、44,316 個約束，求解回報 OPTIMAL。排班約束錯誤 0，排班驗證警告 0；示範跨班與加班均為 0。

Google 來源檢查有 2 項錯誤，原因是示範使用本機資料，沒有 LIVE 來源；報告因此保留 `clinical_publishable=false`。這組固定預假的範例用於檢查執行流程，未用來評估臨床效益或複雜問題的求解效能。

休假帳本另有 7 則提醒（D001：4 則、E001：3 則），與排班驗證警告分開計算。範例未提供完整核定與本人同意，因此未扣新增假，兩人的正式餘額均為 null，符合預期。Excel 夜班津貼為固定費率公式試算，本次未驗證院方費率、薪資核定或試算表軟體重新計算後的金額。

## 代表性反例

- 8/30、8/31 出勤接 9/1～9/4：偵測連續六天。
- 在固定 OFF 排入工作：偵測預假違反。
- 移除核心區域人員：偵測核心缺口。
- 正式帶教缺老師：偵測帶教缺漏。
- 部分工時可上 15 天：目標不被固定成 10 天。
- 需休 8H、補休 5.07H：特休差額為 2.93H；未詢問／未同意不扣新增特休。
- 未取得之預估加班補休：可列試算，不能拿來支付本次休假。
- 一般休假尚有餘額／週期未核定：不自動扣補休。

## 本次來源與診斷回歸

- 模擬 Google 成功讀取與權限不足：可取得試算表物件，拒絕存取時提供 service account email；不再呼叫缺失的函式。
- 機構來源 ID 為空：在驗證憑證或連線前明確拒絕，維持預設停用。
- 憑證缺少 email：回報可辨識的輸入錯誤。
- 求解時限用盡：建議文字不將未取得可行解判定為數學無解。
- 本機提供帶教累積：來源記錄不再聲稱來自未讀取的 `Nurses.xlsx`。

## 全庫內容複查

範圍為 `main` 基準版本的全部 11 個追蹤檔案及本次修正後的同一組檔案：`.gitignore`、`README.md`、`RESEARCH_STATEMENT.md`、`TEST_RESULTS.md`、`examples/leave_balances.json`、`leave_accounting.py`、`main.py`、`requirements.txt`、`scheduler_core.py`、`tests/test_portfolio.py`、`workbook_export.py`。不包含其他分支、Git 歷史、私人外部資料或被忽略的執行產物。

姓名已統一為胡修睿，原有編輯指示已改為專案說明。另修正 AI 使用、研究範圍、離線／LIVE 流程與津貼試算的描述，移除不存在的依賴檔案指引。空白機構設定、未確認同意、未知餘額與待補人力各有程式用途，保留原本狀態。此次內容掃描不等同所有核心分支的功能驗證。

## 尚未驗證

真實 Google 來源整合、所有既有情境分支、完整四週例休／國假自動分類、不同院制與長期人力成效仍未驗證。本次也未進行壓力測試、倫理審查或薪資核定驗證，結果限於上述測試案例。
