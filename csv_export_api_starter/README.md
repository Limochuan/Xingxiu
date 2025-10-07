# Yaoguang Excel Download API

FastAPI 服务，连接阿里云 MySQL（`xingxiu_db`.`xingxiu_daily_report`），按参数筛选并导出 Excel。
已适配：
- `DATE_STR=YYYY-MM-DD`（单日）
- 时间段：`DATE_FROM`+`DATE_TO`，或 `DATE_RANGE=2025-10-01到2025-10-15` / `...to...` / `...-...`
- `PROJECT_NAME=...`
- `COMPANY=...`
- 任意组合；若完全不传，默认 `PROJECT_NAME=Kalteng GIJ 中加一园`

Excel 版式严格按需求：标题三行、行高、列宽 A..AE、Y 列印尼货币两位小数、AC 列数据顶端左对齐、全区域细边框+外框加粗等。

## 部署（Railway）

1. **环境变量**
   - `DB_HOST` (默认: rm-k1a5w7qk9cnm74r25wo.mysql.ap-southeast-5.rds.aliyuncs.com)
   - `DB_PORT` (默认: 3306)
   - `DB_USER` (默认: script_xingxiu)
   - `DB_PASS` (默认: Julong678678678)
   - `DB_NAME` (默认: xingxiu_db)
   - `DB_TABLE` (默认: xingxiu_daily_report)

2. **Start Command**
