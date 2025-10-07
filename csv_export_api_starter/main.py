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
# 配置（优先读环境变量，给默认值）
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

app = FastAPI(title="Yaoguang Excel Download API", version="1.0.1")


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
    s = raw.strip()
    s = s.replace("—", "-").replace("–", "-").replace("到", "-").replace("to", "-")
    s = re.sub(r"-{2,}", "-", s)
    parts = [p.strip() for p in s.split("-") if p.strip()]
    if len(parts) >= 2:
        return parse_date_like(parts[0]), parse_date_like(parts[-1])
    return None, None


def build_where_and_params(date_str, date_from, date_to, date_range, project_name, company):
    where = []
    params = []
    span = (None, None)
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

    where_sql = " WHERE " + " AND ".join(where) if where else ""
    return where_sql, params, span, single_day


def make_excel(rows: List[Dict[str, Any]], headers: List[str], project: str, span, single_day):
    wb = Workbook()
    ws = wb.active
    ws.title = "AI Report"

    widths = [9, 15, 28, 20, 23, 14, 14, 18, 18, 15, 12, 13, 13,
              21, 21, 13, 12, 15, 21, 20, 40, 18, 19, 16, 20, 13,
              21, 9, 200, 20, 15]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 标题
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = f"{project}CATATAN ANALISIS OTOMATIS AI DARI DATA KENDARAAN\n车联网数据AI自动分析记录"
    title_cell.font = Font(name=FONT_TITLE_NAME, size=24, bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 80

    # 表头
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=2, column=j, value=h)
        c.font = Font(name=FONT_BODY_NAME, size=10, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 数据
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, row in enumerate(rows, start=3):
        for j, h in enumerate(headers, start=1):
            v = row.get(h, "")
            c = ws.cell(row=i, column=j, value=v)
            c.font = Font(name=FONT_BODY_NAME, size=10)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border

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
    DATE_STR: Optional[str] = None,
    DATE_FROM: Optional[str] = None,
    DATE_TO: Optional[str] = None,
    DATE_RANGE: Optional[str] = None,
    PROJECT_NAME: Optional[str] = None,
    COMPANY: Optional[str] = None,
):
    project = PROJECT_NAME or DEFAULT_PROJECT
    where_sql, params, span, single_day = build_where_and_params(
        DATE_STR, DATE_FROM, DATE_TO, DATE_RANGE, PROJECT_NAME, COMPANY
    )

    if not where_sql:
        where_sql = " WHERE PROJECT_NAME = %s"
        params = [DEFAULT_PROJECT]

    sql = f"SELECT * FROM `{DB_TABLE}`{where_sql} ORDER BY DATE_STR ASC, ID ASC"

    try:
        conn = connect_db()
        with conn.cursor() as cur:
            cur.execute(f"SHOW COLUMNS FROM `{DB_TABLE}`")
            headers = [r["Field"] for r in cur.fetchall()][:31]
            cur.execute(sql, params)
            rows = cur.fetchall() or []
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"DB error: {e}"})
    finally:
        try:
            conn.close()
        except Exception:
            pass

    if not rows:
        return JSONResponse(status_code=404, content={"error": "No data found for given filters"})

    excel_bytes = make_excel(rows, headers, project, span, single_day)

    date_tag = (
        span[0].strftime("%Y%m%d") if single_day and span[0]
        else f"{span[0].strftime('%Y%m%d')}-{span[1].strftime('%Y%m%d')}" if span[0] and span[1]
        else dt.datetime.now().strftime("%Y%m%d")
    )

    safe_proj = re.sub(r"[\\/:*?\"<>|]+", "_", project)
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
