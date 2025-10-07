#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import io
import re
import datetime as dt
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
FONT_TITLE_NAME = "仿宋"   # 没装字体时，Excel 会自动回退
FONT_BODY_NAME = "仿宋"

app = FastAPI(title="Yaoguang Excel Download API", version="1.1.0")


# =========================
# 工具函数
# =========================
def connect_db():
    return pymysql.connect(
        host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASS,
        database=DB_NAME, charset="utf8mb4", cursorclass=pymysql.cursors.DictCursor
    )

def parse_date_like(s: str) -> Optional[dt.date]:
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

def parse_date_range(raw: str) -> Tuple[Optional[dt.date], Optional[dt.date]]:
    if not raw:
        return None, None
    s = raw.strip().replace("—", "-").replace("–", "-").replace("到", "-").replace("to", "-").replace("TO", "-")
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

def to_period_str(span: Tuple[Optional[dt.date], Optional[dt.date]], single_day: bool) -> str:
    if span[0] and span[1]:
        if single_day:
            return f"（{span[0].strftime('%Y.%m.%d')}）"
        else:
            return f"（{span[0].strftime('%Y.%m.%d')}-{span[1].strftime('%Y.%m.%d')}）"
    return "（）"

def build_where_and_params(
    date_str: Optional[str],
    date_from: Optional[str],
    date_to: Optional[str],
    date_range: Optional[str],
    project_name: Optional[str],
    company: Optional[str],
) -> Tuple[str, List[Any], Tuple[Optional[dt.date], Optional[dt.date]], bool]:
    where, params = [], []
    span: Tuple[Optional[dt.date], Optional[dt.date]] = (None, None)
    single_day = False

    if date_str:
        d = parse_date_like(date_str)
        if d:
            where.append("DATE_STR = %s")
            params.append(d.strftime("%Y-%m-%d"))
            span = (d, d)
            single_day = True

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

    if project_name:
        where.append("PROJECT_NAME = %s")
        params.append(project_name)
    if company:
        where.append("COMPANY = %s")
        params.append(company)

    return (" WHERE " + " AND ".join(where)) if where else "", params, span, single_day

def col_widths_spec() -> List[float]:
    # A..AE 共 31 列
    return [
        9, 15, 28, 20, 23, 14, 14, 18, 18, 15, 12, 13, 13,
        21, 21, 13, 12, 15, 21, 20, 40, 18, 19, 16, 20, 13,
        21, 9, 200, 20, 15
    ]


# =========================
# Excel 按要求渲染
# =========================
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

    # 列宽 A..AE
    widths = col_widths_spec()
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    # ===== 第1行：标题（高度100，仿宋 26 加粗，三行文字，A1:AE1 合并）=====
    max_cols = 31  # A..AE
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_cols)
    title = ws.cell(row=1, column=1)

    # 若无传入日期段且有数据，自动用数据里的最小/最大 DATE_STR
    if not (span[0] and span[1]) and rows:
        span = detect_date_span_from_rows(rows)
        single_day = bool(span[0] and span[1] and span[0] == span[1])
    date_line = to_period_str(span, single_day)

    line1 = f"{project_name_for_title}CATATAN ANALISIS OTOMATIS AI DARI DATA KENDARAAN"
    line2 = "车联网数据AI自动分析记录"
    title.value = f"{line1}\n{line2}\n{date_line}"
    title.font = Font(name=FONT_TITLE_NAME, size=26, bold=True)
    title.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 100

    # ===== 第2行：表头（高度40，仿宋 10 加粗，每个单元格两行字）=====
    ws.row_dimensions[2].height = 40
    header_font = Font(name=FONT_BODY_NAME, size=10, bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    headers = ordered_cols
    for col_idx, col_name in enumerate(headers, start=1):
        c = ws.cell(row=2, column=col_idx)
        c.value = f"{col_name}\n"   # 强制两行
        c.font = header_font
        c.alignment = header_align

    # ===== 数据区（第3行起，行高40，仿宋10；AC列(第29)顶端左对齐，其他上下左右居中；全区域细边框）=====
    body_font = Font(name=FONT_BODY_NAME, size=10, bold=False)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    top_left_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    thin = Side(style="thin", color="000000")
    thick = Side(style="thick", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    for r_idx, row in enumerate(rows, start=3):
        ws.row_dimensions[r_idx].height = 40
        for c_idx, col_name in enumerate(headers, start=1):
            val = row.get(col_name, "")
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = body_font
            cell.alignment = top_left_align if c_idx == 29 else center_align  # AC=29 顶左，其余居中
            cell.border = border_all

    # ===== 最外层加粗外框线 =====
    max_row = max(2, 2 + len(rows))
    max_col = len(headers)
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            left = thick if c == 1 else cell.border.left
            right = thick if c == max_col else cell.border.right
            top = thick if r == 1 else cell.border.top
            bottom = thick if r == max_row else cell.border.bottom
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)

    # ===== Y 列（第25列）设为印尼卢比 Rp 两位小数 =====
    for r in range(3, max_row + 1):
        ws.cell(row=r, column=25).number_format = 'Rp#,##0.00'

    # 输出为字节
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.read()


# =========================
# API 路由
# =========================
@app.get("/healthz")
def healthz():
    # 简单 DB 探活（可选）
    try:
        conn = connect_db()
        conn.close()
        ok = True
    except Exception as e:
        return {"ok": False, "error": str(e), "table": DB_TABLE}
    return {"ok": True, "table": DB_TABLE}

@app.get("/download_excel")
def download_excel(
    DATE_STR: Optional[str] = Query(default=None, description="单日，如 2025-10-07"),
    DATE_FROM: Optional[str] = Query(default=None, description="起始日，如 2025-10-01"),
    DATE_TO: Optional[str] = Query(default=None, description="结束日，如 2025-10-15"),
    DATE_RANGE: Optional[str] = Query(default=None, description="时间段，'2025-10-01到2025-10-15' / '...to...' / '...-...'"),
    PROJECT_NAME: Optional[str] = Query(default=None),
    COMPANY: Optional[str] = Query(default=None),
):
    project_for_title = PROJECT_NAME or DEFAULT_PROJECT

    where_sql, params, span, single_day = build_where_and_params(
        DATE_STR, DATE_FROM, DATE_TO, DATE_RANGE, PROJECT_NAME, COMPANY
    )

    # 无任何过滤时，默认限定项目
    if not where_sql:
        where_sql = " WHERE PROJECT_NAME = %s"
        params = [DEFAULT_PROJECT]

    sql = f"SELECT * FROM `{DB_TABLE}`{where_sql} ORDER BY DATE_STR ASC, ID ASC"

    try:
        conn = connect_db()
        with conn.cursor() as cur:
            # 列顺序（前 31 列）
            cur.execute(f"SHOW COLUMNS FROM `{DB_TABLE}`")
            ordered_cols = [row["Field"] for row in cur.fetchall()][:31]

            # 查询数据
            cur.execute(sql, params)
            fetched = cur.fetchall() or []

            # 只保留 ordered_cols 中的列，确保 A..AE 一致
            rows = [{k: r.get(k, "") for k in ordered_cols} for r in fetched]

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"DB error: {e}"})
    finally:
        try:
            conn.close()
        except Exception:
            pass

    if not rows:
        return JSONResponse(status_code=404, content={"error": "No data found for given filters"})

    # 生成 Excel
    excel_bytes = make_excel(
        rows=rows,
        ordered_cols=ordered_cols,
        project_name_for_title=project_for_title,
        span=span,
        single_day=single_day,
    )

    # 文件名：PROJECT + 日期/日期段
    if span[0] and span[1]:
        date_tag = span[0].strftime("%Y%m%d") if single_day else f"{span[0].strftime('%Y%m%d')}-{span[1].strftime('%Y%m%d')}"
    else:
        date_tag = dt.datetime.now().strftime("%Y%m%d")
    safe_proj = re.sub(r"[\\/:*?\"<>|]+", "_", project_for_title)
    filename = f"{safe_proj}_AI_REPORT_{date_tag}.xlsx"

    # 兼容中文文件名（RFC 5987）
    ascii_fallback = re.sub(r"[^\x00-\x7F]+", "_", filename)
    utf8_filename = quote(filename)
    headers = {"Content-Disposition": f"attachment; filename={ascii_fallback}; filename*=UTF-8''{utf8_filename}"}

    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8000"))
    uvicorn_run("main:app", host="0.0.0.0", port=port, reload=True)
