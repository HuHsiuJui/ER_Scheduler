# 核心截斷修復與驗證

驗證日期：2026-10-12（Asia/Taipei）。

## 原因與修復範圍

GitHub `HuHsiuJui/ER_Scheduler` 的 `5cb22d1fe3daf237ca9faa31c38dd6e208322c63` 中，scheduler_core.py 原第1724行實際含有 `…15445 tokens truncated…ids`，Python 無法解析。比對歷史可見，`a450b5ba18243b943b205435ba09f40eae02a860` 把前版1038行替換成該截斷文字；無法僅從版本紀錄判定操作過程。

從 Git 歷史 `7f11f7ffbf1e7e7baedb2a6aa3b366bc56370127` 恢復該區段，保留目前其餘內容。恢復範圍包含帶教狀態、個人目標、前置檢查、建模及部分求解限制。

另依現有回歸測試補回兩項文件已描述的行為：
- 帶教輸入來源記錄為 NurseProfile.training_completed_seed，避免聲稱本機未讀取的 Nurses.xlsx。
- 求解未取得解的診斷不直接聲稱數學無解，明列逾時不等於已證明無解。

未新增政府日曆下載功能；本機入口仍由 JSON 提供 official_holidays。未改變原有排班限制與目標設計。修復時未修改既有測試。

## 本次實測

Python 3.13.7、NumPy 2.3.3、SciPy 1.18.0、openpyxl 3.1.5。

- `py -3.13 -X utf8 -m unittest discover -s tests -v`：30項全部通過。
- `py -3.13 -m compileall -q main.py scheduler_core.py leave_accounting.py workbook_export.py tests`：通過。
- `py -3.13 main.py --demo --leave-input examples/leave_balances.json`：成功，包含匯出後逐人逐日班別回讀比對。
- 虛構示範：38人、570筆出勤、16312變數、44316限制，OPTIMAL；排班硬條件錯誤0、排班警告0。
- `clinical_publishable=false`：本機示範仍非院方核定班表；真實來源及全部核心情境未全面驗證。
- 差異檢查採 `git -c core.whitespace=cr-at-eol diff --check`，接受原有CRLF換行格式，通過。

舊版 TEST_RESULTS.md 保留為歷史紀錄；目前缺損 GitHub 快照無法執行，不能用歷史通過紀錄替代本次修復驗證。

修復工作分支：`codex/restore-truncated-core`。遠端更新狀態以 GitHub 實際分支與提交為準。
