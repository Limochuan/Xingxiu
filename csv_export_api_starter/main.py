#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import io
import re
import datetime as dt
from urllib.parse import quote
from typing import List, Dict, Any, Optional, Tuple

import pymysql
from fastapi import FastAPI, Query
from fastapi.responses import StreamingResponse, JSONResponse
from uvicorn import run as uvicorn_run
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================
# 配置（优先读环境变量）
# =========================
DB_HOST = os.getenv("DB_HOST", "rm-k1a5w7qk9cnm74r25wo.mysql.ap-southeast-5.rds.aliyuncs.com")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_USER = os.getenv("DB_USER", "script_xingxiu")
DB_PASS = os.getenv("DB_PASS", "Julong678678678")
DB_NAME = os.getenv("DB_NAME", "xingxiu_db")
DB_TABLE = os.getenv("DB_TABLE", "xingxiu_daily_report")

DEFAULT_PROJECT = "Kalteng GIJ 中加一园"

FONT_TITLE_NAME = "仿宋"
FONT_BODY_NAME = "仿宋"

app = FastAPI(title="Yaoguang Excel Download API", version="3.0.0")


# =========================
# 工具函数
# =========================
def connect_db():
    return pymysql.connect(
        host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASS,
        database=DB_NAME, charset="utf8mb4", cursorclass=pymysql.cursors.DictCursor
    )

def parse_date_like(s: Optional[str]) -> Optional[dt.date]:
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s.strip(), fmt).date()
        except Exception:
            pass
    m = re.match(r"^\s*(\d{4})[-./](\d{1,2})[-./](\d{1,2})\s*$", s or "")
    if m:
        y, mo, d = map(int, m.groups())
        return dt.date(y, mo, d)
    return None

def parse_date_range(raw: Optional[str]) -> Tuple[Optional[dt.date], Optional[dt.date]]:
    if not raw:
        return None, None
    s = raw.strip()
    s = s.replace("—", "-").replace("–", "-").replace("到", "-").replace("to", "-").replace("TO", "-")
    s = re.sub(r"-{2,}", "-", s)
    parts = [p for p in s.split("-") if p.strip()]
    if len(parts) >= 2:
        return parse_date_like(parts[0]), parse_date_like(parts[-1])
    return None, None

def detect_date_span_from_rows(rows: List[Dict[str, Any]]) -> Tuple[Optional[dt.date], Optional[dt.date]]:
    if not rows:
        return None, None
    dates = []
    for r in rows:
        v = r.get("DATE_STR")
        if v is None:
            continue
        if isinstance(v, dt.date):
            dates.append(v)
        elif isinstance(v, dt.datetime):
            dates.append(v.date())
        elif isinstance(v, str):
            d = parse_date_like(v)
            if d:
                dates.append(d)
    if not dates:
        return None, None
    return min(dates), max(dates)

def build_where_and_params(
    date_str: Optional[str],
    date_from: Optional[str],
    date_to: Optional[str],
    date_range: Optional[str],
    project_name: Optional[str],
    company: Optional[str],
) -> Tuple[str, List[Any], Tuple[Optional[dt.date], Optional[dt.date]], bool]:
    """
    返回 (where_sql, params, span, single_day_flag)
    """
    where = []
    params: List[Any] = []

    span: Tuple[Optional[dt.date], Optional[dt.date]] = (None, None)
    single_day = False

    # 1) 单日
    if date_str:
        d = parse_date_like(date_str)
        if d:
            where.append("DATE_STR = %s")
            params.append(d.strftime("%Y-%m-%d"))
            span = (d, d)
            single_day = True

    # 2) 时间段
    left = parse_date_like(date_from) if date_from else None
    right = parse_date_like(date_to) if date_to else None
    if not (left and right) and date_range:
        l2, r2 = parse_date_range(date_range)
        left, right = left or l2, right or r2

    if left and right:
        where.append("DATE_STR BETWEEN %s AND %s")
        params.extend([left.strftime("%Y-%m-%d"), right.strftime("%Y-%m-%d")])
        span = (left, right)
        single_day = (left == right)

    # 3) 维度过滤
    #   如果没有显式传 PROJECT_NAME，则默认限定 GIJ
    if project_name:
        where.append("PROJECT_NAME = %s")
        params.append(project_name)
    else:
        where.append("PROJECT_NAME = %s")
        params.append(DEFAULT_PROJECT)

    if company:
        where.append("COMPANY = %s")
        params.append(company)

    where_sql = (" WHERE " + " AND ".join(where)) if where else ""
    return where_sql, params, span, single_day

def auto_select_columns(cursor, limit: int = 31) -> List[str]:
    cursor.execute(f"SHOW COLUMNS FROM `{DB_TABLE}`")
    cols = [row["Field"] for row in cursor.fetchall()]
    return cols[:limit]

def col_widths_spec() -> List[float]:
    # A..AE 共 31 列（索引 1..31）
    return [
        9, 15, 28, 20, 23, 14, 14, 18, 18, 15, 12, 13, 13,
        21, 21, 13, 12, 15, 21, 20, 40, 18, 19, 16, 20, 13,
        21, 9, 200, 20, 15
    ]

def to_period_str(span: Tuple[Optional[dt.date], Optional[dt.date]], single_day: bool) -> str:
    if span[0] and span[1]:
        if single_day:
            return f"（{span[0].strftime('%Y.%m.%d')}）"
        else:
            return f"（{span[0].strftime('%Y.%m.%d')}-{span[1].strftime('%Y.%m.%d')}）"
    return "（）"

def make_excel(
    rows: List[Dict[str, Any]],
    ordered_cols: List[str],
    project_name_for_title: str,
    span: Tuple[Optional[dt.date], Optional[dt.date]],
    single_day: bool
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "AI Report"

    # 列宽
    widths = col_widths_spec()
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    # —— 标题行（第1行，合并 A1:AE1）——
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=31)
    title_cell = ws.cell(row=1, column=1)

    line1 = f"{project_name_for_title}CATATAN ANALISIS OTOMATIS AI DARI DATA KENDARAAN"
    line2 = "车联网数据AI自动分析记录"
    # 若用户没传日期，则尝试用数据里的最小/最大 DATE_STR
    if not (span[0] and span[1]) and rows:
        span = detect_date_span_from_rows(rows)
        single_day = (span[0] and span[1] and span[0] == span[1])
    line3 = to_period_str(span, single_day)

    title_cell.value = f"{line1}\n{line2}\n{line3}"
    title_cell.font = Font(name=FONT_TITLE_NAME, size=26, bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 100

    # —— 表头（第2行）——
    ws.row_dimensions[2].height = 40
    header_font = Font(name=FONT_BODY_NAME, size=10, bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 用 ordered_cols 填表头（两行：字段名 + 空行）
    for col_idx, col_name in enumerate(ordered_cols, start=1):
        c = ws.cell(row=2, column=col_idx)
        c.value = f"{col_name}\n"
        c.font = header_font
        c.alignment = header_align

    # —— 数据区（第3行开始）——
    body_font = Font(name=FONT_BODY_NAME, size=10, bold=False)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    top_left_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    thin = Side(style="thin", color="000000")
    thick = Side(style="thick", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    for r_idx, row in enumerate(rows, start=3):
        ws.row_dimensions[r_idx].height = 40
        for c_idx, col_name in enumerate(ordered_cols, start=1):
            val = row.get(col_name, "")
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = body_font
            # AC 列（第29列）从第3行开始顶端左对齐，其余居中
            if c_idx == 29:
                cell.alignment = top_left_align
            else:
                cell.alignment = center_align
            cell.border = border_all

    # —— 外层加粗边框（整个使用区域）——
    max_row = max(2, 2 + len(rows))
    max_col = len(ordered_cols)
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            left = thick if c == 1 else cell.border.left
            right = thick if c == max_col else cell.border.right
            top = thick if r == 1 else cell.border.top
            bottom = thick if r == max_row else cell.border.bottom
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)

    # —— Y 列（第25列）设置印尼卢比两位小数 —— 
    currency_col = 25
    for r in range(3, max_row + 1):
        ws.cell(row=r, column=currency_col).number_format = 'Rp#,##0.00'

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.read()


# =========================
# API 路由
# =========================
@app.get("/healthz")
def healthz():
    return {"ok": True, "table": DB_TABLE}

@app.get("/download_excel")
def download_excel(
    DATE_STR: Optional[str] = Query(default=None, description="单日，如 2025-10-07"),
    DATE_FROM: Optional[str] = Query(default=None, description="起始日，如 2025-10-01"),
    DATE_TO: Optional[str] = Query(default=None, description="结束日，如 2025-10-15"),
    DATE_RANGE: Optional[str] = Query(default=None, description="时间段，支持 '2025-10-01到2025-10-15'、'...to...'、'...-...'"),
    PROJECT_NAME: Optional[str] = Query(default=None),
    COMPANY: Optional[str] = Query(default=None),
):
    # WHERE 及日期段信息
    where_sql, params, span, single_day = build_where_and_params(
        DATE_STR, DATE_FROM, DATE_TO, DATE_RANGE, PROJECT_NAME, COMPANY
    )
    project_for_title = PROJECT_NAME or DEFAULT_PROJECT

    sql = f"SELECT * FROM `{DB_TABLE}`{where_sql} ORDER BY DATE_STR ASC, ID ASC"

    try:
        conn = connect_db()
        with conn.cursor() as cur:
            ordered_cols = auto_select_columns(cur, limit=31)
            cur.execute(sql, params)
            rows = cur.fetchall() or []
            # 仅保留前 31 列键值
            trimmed_rows = [{k: r.get(k, "") for k in ordered_cols} for r in rows]
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"DB error: {e}"})
    finally:
        try:
            conn.close()
        except Exception:
            pass

    if not trimmed_rows:
        return JSONResponse(status_code=404, content={"error": "No data found for given filters"})

    # 生成 Excel
    excel_bytes = make_excel(
        rows=trimmed_rows,
        ordered_cols=ordered_cols,
        project_name_for_title=project_for_title,
        span=span,
        single_day=single_day,
    )

    # 生成文件名（支持中文，RFC5987）
    if span[0] and span[1]:
        date_tag = span[0].strftime("%Y%m%d") if single_day else f"{span[0].strftime('%Y%m%d')}-{span[1].strftime('%Y%m%d')}"
    else:
        date_tag = dt.datetime.now().strftime("%Y%m%d")

    safe_proj = re.sub(r"[\\/:*?\"<>|]+", "_", project_for_title)
    filename = f"{safe_proj}_AI_REPORT_{date_tag}.xlsx"
    ascii_fallback = re.sub(r"[^\x00-\x7F]+", "_", filename)
    utf8_filename = quote(filename)
    headers = {
        "Content-Disposition": f"attachment; filename={ascii_fallback}; filename*=UTF-8''{utf8_filename}"
    }

    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


if __name__ == "__main__":
    port = int(os.getenv("PORT", "8000"))
    uvicorn_run("main:app", host="0.0.0.0", port=port, reload=True)
