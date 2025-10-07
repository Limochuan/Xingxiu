#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Yaoguang Excel 导出服务
--------------------------------------------------------
功能说明：
1. 支持从 MySQL 表导出 Excel（最多 31 列）。
2. 支持日期过滤（DATE_STR、DATE_FROM/DATE_TO、DATE_RANGE）。
3. 支持 COMPANY、PROJECT_NAME 等维度。
4. 若未显式传 PROJECT_NAME，则自动限定为 GIJ。
5. Excel 输出格式：
   - 三行标题（项目名 + 印尼语标题 + 中文标题 + 日期区间）
   - 列宽固定（A~AE 共 31 列）
   - 第一行行高 100，第二行行高 40，数据行高 40
   - 表头加粗、自动换行、居中
   - 第 29 列 AC 顶端左对齐，其余居中
   - 第 25 列 Y 显示为印尼卢比 Rp#,##0.00
   - 全表细边框，外框加粗
   - 文件名支持中文（RFC5987 编码）
--------------------------------------------------------
部署环境：
- Python 3.11+
- FastAPI + Uvicorn + openpyxl + PyMySQL
--------------------------------------------------------
作者：Jimmy 张杰铭
版本：3.3.0 (2025-10-07)
"""

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

# ===========================
# 基本配置
# ===========================

DB_HOST = os.getenv("DB_HOST", "rm-k1a5w7qk9cnm74r25wo.mysql.ap-southeast-5.rds.aliyuncs.com")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_USER = os.getenv("DB_USER", "script_xingxiu")
DB_PASS = os.getenv("DB_PASS", "Julong678678678")
DB_NAME = os.getenv("DB_NAME", "xingxiu_db")
DB_TABLE = os.getenv("DB_TABLE", "xingxiu_daily_report")

DEFAULT_PROJECT = "Kalteng GIJ 中加一园"

app = FastAPI(title="Yaoguang Excel Download Service", version="3.3.0")


# ===========================
# 工具函数
# ===========================

def connect_db():
    """建立 MySQL 连接"""
    return pymysql.connect(
        host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASS,
        database=DB_NAME, charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor
    )


def parse_date_like(s: Optional[str]) -> Optional[dt.date]:
    """解析字符串为日期对象"""
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s.strip(), fmt).date()
        except Exception:
            continue
    return None


def parse_date_range(raw: Optional[str]) -> Tuple[Optional[dt.date], Optional[dt.date]]:
    """解析日期范围，如 '2025-10-01到2025-10-07'"""
    if not raw:
        return None, None
    s = re.sub(r"[—–\-到toTO]+", "-", raw)
    parts = [p.strip() for p in s.split("-") if p.strip()]
    if len(parts) >= 2:
        return parse_date_like(parts[0]), parse_date_like(parts[-1])
    return None, None


def detect_date_span_from_rows(rows: List[Dict[str, Any]]) -> Tuple[Optional[dt.date], Optional[dt.date]]:
    """自动从查询结果中提取日期范围"""
    if not rows:
        return None, None
    dates = []
    for r in rows:
        v = r.get("DATE_STR")
        if not v:
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


def col_widths_spec() -> List[int]:
    """返回列宽定义"""
    return [
        9, 15, 28, 20, 23, 14, 14, 18, 18, 15, 12, 13, 13,
        21, 21, 13, 12, 15, 21, 20, 40, 18, 19, 16, 20, 13,
        21, 9, 200, 20, 15
    ]


def to_period_str(span: Tuple[Optional[dt.date], Optional[dt.date]], single_day: bool) -> str:
    """生成日期区间文本"""
    if span[0] and span[1]:
        if single_day:
            return f"（{span[0].strftime('%Y.%m.%d')}）"
        else:
            return f"（{span[0].strftime('%Y.%m.%d')}-{span[1].strftime('%Y.%m.%d')}）"
    return "（）"


# ===========================
# 核心逻辑
# ===========================

def build_where_clause(
    date_str: Optional[str],
    date_from: Optional[str],
    date_to: Optional[str],
    date_range: Optional[str],
    project_name: Optional[str],
    company: Optional[str],
) -> Tuple[str, List[Any], Tuple[Optional[dt.date], Optional[dt.date]], bool, str]:
    """
    构造 SQL WHERE 语句 + 参数
    返回: where_sql, params, (span_from, span_to), single_day, project_for_title
    """
    where = []
    params: List[Any] = []
    span: Tuple[Optional[dt.date], Optional[dt.date]] = (None, None)
    single_day = False
    project_for_title = project_name or DEFAULT_PROJECT

    # 1) 日期过滤
    if date_str:
        d = parse_date_like(date_str)
        if d:
            where.append("DATE_STR = %s")
            params.append(d.strftime("%Y-%m-%d"))
            span = (d, d)
            single_day = True

    left, right = parse_date_like(date_from), parse_date_like(date_to)
    if not (left and right) and date_range:
        l2, r2 = parse_date_range(date_range)
        left, right = left or l2, right or r2
    if left and right:
        where.append("DATE_STR BETWEEN %s AND %s")
        params += [left.strftime("%Y-%m-%d"), right.strftime("%Y-%m-%d")]
        span = (left, right)
        single_day = (left == right)

    # 2) 默认 GIJ 限定
    if not project_name:
        where.append("PROJECT_NAME = %s")
        params.append(DEFAULT_PROJECT)
        project_for_title = DEFAULT_PROJECT
    else:
        where.append("PROJECT_NAME = %s")
        params.append(project_name)

    # 3) 公司过滤
    if company:
        where.append("COMPANY = %s")
        params.append(company)

    where_sql = " WHERE " + " AND ".join(where) if where else ""
    return where_sql, params, span, single_day, project_for_title


def make_excel(
    rows: List[Dict[str, Any]],
    ordered_cols: List[str],
    project_name_for_title: str,
    span: Tuple[Optional[dt.date], Optional[dt.date]],
    single_day: bool
) -> bytes:
    """生成 Excel 文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "AI Report"

    # 列宽
    for idx, w in enumerate(col_widths_spec(), start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    # 样式定义
    font_title = Font(name="仿宋", size=26, bold=True)
    font_header = Font(name="仿宋", size=10, bold=True)
    font_body = Font(name="仿宋", size=10)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)
    thin = Side(style="thin", color="000000")
    thick = Side(style="thick", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 第一行标题
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=31)
    title_cell = ws.cell(row=1, column=1)
    line1 = f"{project_name_for_title}CATATAN ANALISIS OTOMATIS AI DARI DATA KENDARAAN"
    line2 = "车联网数据AI自动分析记录"
    line3 = to_period_str(span, single_day)
    title_cell.value = f"{line1}\n{line2}\n{line3}"
    title_cell.font = font_title
    title_cell.alignment = align_center
    ws.row_dimensions[1].height = 100

    # 第二行表头
    ws.row_dimensions[2].height = 40
    for idx, col in enumerate(ordered_cols, start=1):
        cell = ws.cell(row=2, column=idx, value=f"{col}\n")
        cell.font = font_header
        cell.alignment = align_center
        cell.border = border_all

    # 数据部分
    for r_idx, row in enumerate(rows, start=3):
        ws.row_dimensions[r_idx].height = 40
        for c_idx, col in enumerate(ordered_cols, start=1):
            val = row.get(col, "")
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_body
            cell.alignment = align_left_top if c_idx == 29 else align_center
            cell.border = border_all

    # 外框加粗
    max_row, max_col = ws.max_row, ws.max_column
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(
                left=thick if c == 1 else cell.border.left,
                right=thick if c == max_col else cell.border.right,
                top=thick if r == 1 else cell.border.top,
                bottom=thick if r == max_row else cell.border.bottom,
            )

    # Y 列格式
    for r in range(3, max_row + 1):
        ws.cell(row=r, column=25).number_format = "Rp#,##0.00"

    # 保存
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.read()


# ===========================
# API 路由
# ===========================

@app.get("/healthz")
def healthz():
    """健康检查"""
    return {"ok": True, "table": DB_TABLE}


@app.get("/download_excel")
def download_excel(
    DATE_STR: Optional[str] = Query(None),
    DATE_FROM: Optional[str] = Query(None),
    DATE_TO: Optional[str] = Query(None),
    DATE_RANGE: Optional[str] = Query(None),
    PROJECT_NAME: Optional[str] = Query(None),
    COMPANY: Optional[str] = Query(None),
):
    """主下载接口"""

    # 构建 WHERE 条件
    where_sql, params, span, single_day, project_for_title = build_where_clause(
        DATE_STR, DATE_FROM, DATE_TO, DATE_RANGE, PROJECT_NAME, COMPANY
    )

    sql = f"SELECT * FROM `{DB_TABLE}`{where_sql} ORDER BY DATE_STR ASC, ID ASC"

    # 执行 SQL
    try:
        conn = connect_db()
        with conn.cursor() as cur:
            cur.execute(f"SHOW COLUMNS FROM `{DB_TABLE}`")
            cols = [row["Field"] for row in cur.fetchall()][:31]
            cur.execute(sql, params)
            rows = cur.fetchall() or []
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
    finally:
        try:
            conn.close()
        except Exception:
            pass

    if not rows:
        return JSONResponse(status_code=404, content={"error": "No data found for given filters"})

    excel_bytes = make_excel(rows, cols, project_for_title, span, single_day)

    # 生成文件名
    if span[0] and span[1]:
        date_tag = span[0].strftime("%Y%m%d") if single_day else f"{span[0].strftime('%Y%m%d')}-{span[1].strftime('%Y%m%d')}"
    else:
        date_tag = dt.datetime.now().strftime("%Y%m%d")

    filename = f"{project_for_title}_AI_REPORT_{date_tag}.xlsx"
    ascii_name = re.sub(r"[^\x00-\x7F]+", "_", filename)
    utf8_name = quote(filename)
    headers = {"Content-Disposition": f"attachment; filename={ascii_name}; filename*=UTF-8''{utf8_name}"}

    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


# ===========================
# 启动服务
# ===========================

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8000"))
    uvicorn_run("main:app", host="0.0.0.0", port=port, reload=True)
