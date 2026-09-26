# 本次驗證紀錄

日期：2026-09-14（Asia/Taipei）

環境：Windows、Python 3.13、NumPy 2.3.3、SciPy 1.18.0、openpyxl 3.1.5。

## 已執行

- `python -m unittest discover -s tests -v`：25 項測試全部通過。
- `python main.py --demo --leave-input examples/leave_balances.json`：成功輸出研究展示 Excel 與 JSON 報告。
- `python -m compileall -q main.py scheduler_core.py leave_accounting.py workbook_export.py tests`：通過。
- Excel 匯出後回讀，逐人逐日比對班別，與求解結果一致。

## 匿名範例結果

38 名虛構人員，30 天，共 570 筆出勤配置。示範模型 16,312 個變數、44,316 個約束，求解回報 OPTIMAL。排班約束錯誤 0，警告 0；示範跨班與加班均為 0。

Google 來源檢查仍保留 2 項錯誤（示範不具 LIVE 來源），`clinical_publishable=false`。這是有意保留的來源保護，不將本機示範偽裝成院方正式資料。此固定預假的簡化範例不是臨床效益測試，也不是複雜問題效能基準。

## 代表性反例

- 8/30、8/31 出勤接 9/1～9/4：偵測連續六天。
- 在固定 OFF 排入工作：偵測預假違反。
- 移除核心區域人員：偵測核心缺口。
- 正式帶教缺老師：偵測帶教缺漏。
- 部分工時可上 15 天：目標不被固定成 10 天。
- 需休 8H、補休 5.07H：特休差額為 2.93H；未詢問／未同意不扣新增特休。
- 未取得之預估加班補休：可列試算，不能拿來支付本次休假。
- 一般休假尚有餘額／週期未核定：不自動扣補休。

## 尚未驗證

真實 Google 來源整合、所有既有情境分支、完整四週例休／國假自動分類、臨床正式使用、不同院制及長期人力成效。未做壓力測試、倫理審查或薪資核定驗證。申請資料不可將以上未驗證項目寫成已完成成果。

