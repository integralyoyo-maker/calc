# 違規入案代碼表更新方式
1. 取得新版 Excel（含「條款代碼表」「不得郵繳條款」「慢車不得郵繳條款」「吊扣代碼」四個工作表）。
2. 執行：`python3 tools/build_codes.py 新版代碼表.xlsx`（需 openpyxl；檔名含 YYYYMMDD 西元日期會顯示為檔案日期）。
3. 會重新產生 `data/violation_codes.json`（現行）與 `data/violation_codes_old.json`（已不適用），commit 後推送即可。
代碼表版本＝Excel 檔名上的日期（民國，例如 20261125 → 1151125）；現行／已不適用依備註欄「新增／不適用／重新適用」判斷；應到案依「不得郵繳條款」工作表比對。
