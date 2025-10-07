#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, io, re, datetime as dt
from typing import List, Dict, Any, Optional, Tuple
from urllib.parse import quote

import pymysql
from fastapi import FastAPI, Query
from fastapi.responses import StreamingResponse, JSONResponse
from uvicorn import run as uvicorn_run
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================
# 配置
# =========================
DB_HOST = os.getenv("DB_HOST", "rm-k1a5w7qk9cnm74r25wo.mysql.ap-southeast-5.rds.aliyuncs.com")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_USER = os.getenv("DB_USER", "script_xingxiu")
DB_PASS = os.getenv("DB_PASS", "Julong678678678")
DB_NAME = os.getenv("DB_NAME", "xingxiu_db")
DB_TABLE = os.getenv("DB_TABLE", "xingxiu_daily_report")

DEFAULT_PROJECT = "Kalteng GIJ 中加一园"
FONT_TITLE_NAME = "仿宋"
FONT_BODY_NAME  = "仿宋"

app = FastAPI(title="Yaoguang Excel Download API", version="1.2.3")

# =========================
# 固定列定义（A..AE 共 31 列，顺序不可变）
# (显示字段名, 中文, 取值别名列表)
COLUMN_SPECS = [
    ("ID",                 "序号",         ["ID"]),
    ("DEVICE_ID",          "工时通编号",   ["DEVICE_ID"]),
    ("PROJECT_NAME",       "系统项目名称", ["PROJECT_NAME"]),
    ("MECHANICAL_NO",      "系统编号",     ["MECHANICAL_NO"]),
    ("CAR_TYPE",           "系统类型",     ["CAR_TYPE"]),
    ("DATE_STR",           "数据日期",     ["DATE_STR"]),
    ("RENT_TYPE",          "系统计算类型", ["RENT_TYPE"]),
    ("VALID_DURATION",     "有效工时",     ["VALID_DURATION"]),
    ("IDLING_DURATION",    "怠速工时",     ["IDLING_DURATION"]),
    ("VALID_PERCENT",      "工时有效比%",  ["VALID_PERCENT"]),
    ("DAY_OIL",            "油耗",         ["DAY_OIL"]),
    ("DAY_REFUEL",         "加油",         ["DAY_REFUEL"]),
    ("DAY_MILEAGE",        "里程",         ["DAY_MILEAGE"]),
    ("WORKHOUR_AVG_OIL",   "工时平均油耗", ["WORKHOUR_AVG_OIL"]),
    ("TRANSPORT_AVG_OIL",  "运输平均油耗", ["TRANSPORT_AVG_OIL"]),
    ("COMPANY",            "公司",         ["COMPANY"]),
    ("ESTATE",             "区域",         ["ESTATE"]),
    ("MACHINE_TYPE",       "机械类型",     ["MACHINE_TYPE"]),
    ("MACHINE_CATEGORY",   "机械具体类型", ["MACHINE_CATEGORY"]),
    ("MACHINE_NO",         "园区编号",     ["MACHINE_NO"]),
    ("BRAND_SPEC",         "具体型号",     ["BRAND_SPEC"]),
    ("ORDER_NO SAP",       "SAP订单号",    ["ORDER_NO SAP","ORDER_NO_SAP","ORDER_NO"]),
    ("PURCHASE_DATE",      "购买日期",     ["PURCHASE_DATE"]),
    ("DRIVER_COUNT",       "司机数量",     ["DRIVER_COUNT"]),
    ("PURCH_PRICE",        "购买价格",     ["PURCH_PRICE","PURCHASE_PRICE"]),   # ← Y 列
    ("FUEL_DIFF",          "标准油耗差",   ["FUEL_DIFF"]),
    ("INSERT_TIME",        "数据插入时间", ["INSERT_TIME","CREATE_TIME"]),      # ← AA 列
    ("SCORE",              "得分",         ["SCORE"]),
    ("SUMMARY",            "AI分析",       ["SUMMARY"]),                        # ← AC 列
    ("ANALYSIS_TIME",      "分析时间",     ["ANALYSIS_TIME"]),
    ("MODEL_NAME",         "使用模型",     ["MODEL_NAME"]),
]

FIELD_TO_COLIDX = {field: idx+1 for idx, (field, _, _) in enumerate(COLUMN_SPECS)}
COLIDX_Y  = FIELD_TO_COLIDX["PURCH_PRICE"]   # 25
COLIDX_AC = FIELD_TO_COLIDX["SUMMARY"]       # 29

# =========================
# 工具函数
# =========================
def connect_db():
    return pymysql.connect(
        host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASS,
        database=DB_NAME, charset="utf8mb4", cursorclass=pymysql.cursors.DictCursor
    )

def parse_date_like(s: Optional[str]) -> Optional[dt.date]:
    if not s: return None
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try: return dt.datetime.strptime(s.strip(), fmt).date()
        except Exception: pass
    m = re.match(r"^\s*(\d{4})[-./](\d{1,2})[-./](\d{1,2})\s*$", s or "")
    if m:
        y, mo, d = map(int, m.groups())
        return dt.date(y, mo, d)
    return None

def parse_date_range(raw: Optional[str]) -> Tuple[Optional[dt.date], Optional[dt.date]]:
    if not raw: return None, None
    s = raw.strip().replace("—","-").replace("–","-").replace("到","-").replace("to","-").replace("TO","-")
    s = re.sub(r"-{2,}", "-", s)
    parts = [p for p in s.split("-") if p.strip()]
    if len(parts) >= 2:
        return parse_date_like(parts[0]), parse_date_like(parts[-1])
    return None, None

def detect_date_span_from_rows(rows: List[Dict[str, Any]]) -> Tuple[Optional[dt.date], Optional[dt.date]]:
    dates = []
    for r in rows:
        v = r.get("DATE_STR")
        if isinstance(v, dt.datetime): dates.append(v.date())
        elif isinstance(v, dt.date):   dates.append(v)
        elif isinstance(v, str):
            d = parse_date_like(v)
            if d: dates.append(d)
    return (min(dates), max(dates)) if dates else (None, None)

def build_where_and_params(date_str, date_from, date_to, date_range, project_name, company):
    where, params = [], []
    span = (None, None); single_day = False

    d = parse_date_like(date_str) if date_str else None
    if d:
        where.append("DATE_STR = %s"); params.append(d.strftime("%Y-%m-%d"))
        span = (d, d); single_day = True

    left = parse_date_like(date_from) if date_from else None
    right = parse_date_like(date_to) if date_to else None
    if not (left and right) and date_range:
        l2, r2 = parse_date_range(date_range)
        left, right = left or l2, right or r2
    if left and right:
        where.append("DATE_STR BETWEEN %s AND %s")
        params.extend([left.strftime("%Y-%m-%d"), right.strftime("%Y-%m-%d")])
        span = (left, right); single_day = (left == right)

    # 只有当传入了非空 project_name 时才加该条件（默认 GIJ 在外层统一处理）
    if project_name:
        where.append("PROJECT_NAME = %s"); params.append(project_name)

    if company:
        where.append("COMPANY = %s"); params.append(company)

    return (" WHERE " + " AND ".join(where)) if where else "", params, span, single_day

def col_widths_spec() -> List[float]:
    return [9,15,28,20,23,14,14,18,18,15,12,13,13,21,21,13,12,15,21,20,40,18,19,16,20,13,21,9,200,20,15]

def to_period_str(span: Tuple[Optional[dt.date], Optional[dt.date]], single_day: bool) -> str:
    if span[0] and span[1]:
        return f"（{span[0].strftime('%Y.%m.%d')}）" if single_day \
               else f"（{span[0].strftime('%Y.%m.%d')}-{span[1].strftime('%Y.%m.%d')}）"
    return "（）"

# =========================
# Excel 渲染（严格按版式）
# =========================
def _first_present(row: Dict[str, Any], alts: list) -> Any:
    for k in alts:
        if k in row and row[k] is not None:
            return row[k]
    return ""

def make_excel(rows: List[Dict[str, Any]], project_name_for_title: str,
               span: Tuple[Optional[dt.date], Optional[dt.date]], single_day: bool) -> bytes:
    wb = Workbook(); ws = wb.active; ws.title = "AI Report"

    # 列宽
    for idx, w in enumerate(col_widths_spec(), start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    # 第1行 标题
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=31)
    if not (span[0] and span[1]) and rows:
        span = detect_date_span_from_rows(rows); single_day = (span[0] and span[1] and span[0]==span[1])
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = f"{project_name_for_title}CATATAN ANALISIS OTOMATIS AI DARI DATA KENDARAAN\n车联网数据AI自动分析记录\n{to_period_str(span, single_day)}"
    title_cell.font = Font(name=FONT_TITLE_NAME, size=26, bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 100

    # 第2行 表头（两行：字段名+中文）
    ws.row_dimensions[2].height = 40
    header_font  = Font(name=FONT_BODY_NAME, size=10, bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx, (field, zh, _) in enumerate(COLUMN_SPECS, start=1):
        c = ws.cell(row=2, column=col_idx, value=f"{field}\n{zh}")
        c.font = header_font; c.alignment = header_align

    # 数据行
    body_font    = Font(name=FONT_BODY_NAME, size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    top_left     = Alignment(horizontal="left",   vertical="top",   wrap_text=True)
    thin, thick  = Side(style="thin", color="000000"), Side(style="thick", color="000000")
    border_all   = Border(left=thin, right=thin, top=thin, bottom=thin)

    for r_idx, row in enumerate(rows, start=3):
        ws.row_dimensions[r_idx].height = 40
        for c_idx, (_, _, alts) in enumerate(COLUMN_SPECS, start=1):
            val = _first_present(row, alts)
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = body_font
            cell.alignment = top_left if c_idx == COLIDX_AC else center_align
            cell.border = border_all

    # 外层加粗外框
    max_row = max(2, 2 + len(rows)); max_col = 31
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            left   = thick if c == 1       else cell.border.left
            right  = thick if c == max_col else cell.border.right
            top    = thick if r == 1       else cell.border.top
            bottom = thick if r == max_row else cell.border.bottom
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)

    # Y 列（PURCH_PRICE）印尼卢比两位小数
    for r in range(3, max_row + 1):
        ws.cell(row=r, column=COLIDX_Y).number_format = 'Rp#,##0.00'

    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio.read()

# =========================
# API
# =========================
@app.get("/healthz")
def healthz():
    try:
        conn = connect_db(); conn.close()
        return {"ok": True, "table": DB_TABLE}
    except Exception as e:
        return {"ok": False, "error": str(e), "table": DB_TABLE}

@app.get("/download_excel")
def download_excel(
    DATE_STR: Optional[str] = Query(default=None),
    DATE_FROM: Optional[str] = Query(default=None),
    DATE_TO: Optional[str] = Query(default=None),
    DATE_RANGE: Optional[str] = Query(default=None),
    PROJECT_NAME: Optional[str] = Query(default=None),
    COMPANY: Optional[str] = Query(default=None),
):
    # —— 关键：参数规范化（空字符串 → None；去掉前后空格）
    PROJECT_NAME = (PROJECT_NAME.strip() or None) if PROJECT_NAME is not None else None
    COMPANY      = (COMPANY.strip() or None) if COMPANY is not None else None
    DATE_STR     = DATE_STR.strip() if DATE_STR else None
    DATE_FROM    = DATE_FROM.strip() if DATE_FROM else None
    DATE_TO      = DATE_TO.strip() if DATE_TO else None
    DATE_RANGE   = DATE_RANGE.strip() if DATE_RANGE else None

    # 构建 where
    where_sql, params, span, single_day = build_where_and_params(
        DATE_STR, DATE_FROM, DATE_TO, DATE_RANGE, PROJECT_NAME, COMPANY
    )

    # —— 默认 GIJ：只要没有“非空 PROJECT_NAME”，就强制限定 GIJ（无论是否有日期/公司条件）
    if not PROJECT_NAME:
        if where_sql:
            where_sql += " AND PROJECT_NAME = %s"
            params.append(DEFAULT_PROJECT)
        else:
            where_sql = " WHERE PROJECT_NAME = %s"
            params = [DEFAULT_PROJECT]
        project_for_title = DEFAULT_PROJECT
    else:
        project_for_title = PROJECT_NAME

    sql = f"SELECT * FROM `{DB_TABLE}`{where_sql} ORDER BY DATE_STR ASC, ID ASC"

    try:
        conn = connect_db()
        with conn.cursor() as cur:
            cur.execute(sql, params)
            rows = cur.fetchall() or []
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"DB error: {e}"})
    finally:
        try: conn.close()
        except Exception: pass

    if not rows:
        return JSONResponse(status_code=404, content={"error": "No data found for given filters"})

    excel_bytes = make_excel(rows, project_for_title, span, single_day)

    # 文件名（PROJECT + 日期/日期段；兼容中文）
    if span[0] and span[1]:
        date_tag = span[0].strftime("%Y%m%d") if single_day else f"{span[0].strftime('%Y%m%d')}-{span[1].strftime('%Y%m%d')}"
    else:
        date_tag = dt.datetime.now().strftime("%Y%m%d")
    safe_proj = re.sub(r"[\\/:*?\"<>|]+", "_", project_for_title)
    filename = f"{safe_proj}_AI_REPORT_{date_tag}.xlsx"
    ascii_fallback = re.sub(r"[^\x00-\x7F]+", "_", filename)
    utf8_filename = quote(filename)
    headers = {"Content-Disposition": f"attachment; filename={ascii_fallback}; filename*=UTF-8''{utf8_filename}"}

    return StreamingResponse(io.BytesIO(excel_bytes),
                             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers=headers)

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8000"))
    uvicorn_run("main:app", host="0.0.0.0", port=port, reload=True)
