#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, io, re, datetime as dt
from urllib.parse import quote
from typing import List, Dict, Any, Optional, Tuple

import pymysql
from fastapi import FastAPI, Query
from fastapi.responses import StreamingResponse, JSONResponse
from uvicorn import run as uvicorn_run
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===== 数据库配置 =====
DB_HOST = os.getenv("DB_HOST", "rm-k1a5w7qk9cnm74r25wo.mysql.ap-southeast-5.rds.aliyuncs.com")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_USER = os.getenv("DB_USER", "script_xingxiu")
DB_PASS = os.getenv("DB_PASS", "Julong678678678")
DB_NAME = os.getenv("DB_NAME", "xingxiu_db")
DB_TABLE = os.getenv("DB_TABLE", "xingxiu_daily_report")
DEFAULT_PROJECT = "Kalteng GIJ 中加一园"

app = FastAPI(title="Yaoguang Excel API", version="3.1.0")

# ===== 字体样式 =====
FONT_TITLE = Font(name="仿宋", size=26, bold=True)
FONT_HEADER = Font(name="仿宋", size=10, bold=True)
FONT_BODY = Font(name="仿宋", size=10)

# ===== 工具函数 =====
def connect_db():
    return pymysql.connect(
        host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASS,
        database=DB_NAME, charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor
    )

def parse_date(s: Optional[str]) -> Optional[dt.date]:
    if not s: return None
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try: return dt.datetime.strptime(s.strip(), fmt).date()
        except: pass
    return None

def parse_range(raw: Optional[str]) -> Tuple[Optional[dt.date], Optional[dt.date]]:
    if not raw: return None, None
    s = re.sub(r"[-–—到toTO]+", "-", raw)
    parts = [p.strip() for p in s.split("-") if p.strip()]
    if len(parts) >= 2:
        return parse_date(parts[0]), parse_date(parts[-1])
    return None, None

def col_widths() -> List[int]:
    return [9,15,28,20,23,14,14,18,18,15,12,13,13,21,21,13,12,15,21,20,
            40,18,19,16,20,13,21,9,200,20,15]

def period_str(span, single):
    if span[0] and span[1]:
        return f"（{span[0].strftime('%Y.%m.%d')}）" if single \
            else f"（{span[0].strftime('%Y.%m.%d')}-{span[1].strftime('%Y.%m.%d')}）"
    return "（）"

# ===== Excel 生成 =====
def make_excel(rows, headers, project, span, single):
    wb = Workbook(); ws = wb.active; ws.title = "AI Report"
    for i, w in enumerate(col_widths(), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    thin, thick = Side("thin","000000"), Side("thick","000000")
    border_all = Border(left=thin,right=thin,top=thin,bottom=thin)
    align_center = Alignment(horizontal="center",vertical="center",wrap_text=True)
    align_top_left = Alignment(horizontal="left",vertical="top",wrap_text=True)

    # 标题
    ws.merge_cells("A1:AE1")
    t = ws.cell(1,1)
    t.value = f"{project}CATATAN ANALISIS OTOMATIS AI DARI DATA KENDARAAN\n车联网数据AI自动分析记录\n{period_str(span,single)}"
    t.font, t.alignment = FONT_TITLE, align_center
    ws.row_dimensions[1].height = 100

    # 表头
    ws.row_dimensions[2].height = 40
    for j,h in enumerate(headers,1):
        c = ws.cell(2,j,h+"\n"); c.font,c.alignment,c.border = FONT_HEADER,align_center,border_all

    # 数据
    for i,row in enumerate(rows,3):
        ws.row_dimensions[i].height = 40
        for j,h in enumerate(headers,1):
            v = row.get(h,""); c = ws.cell(i,j,v)
            c.font = FONT_BODY
            c.alignment = align_top_left if j==29 else align_center
            c.border = border_all

    # 外框加粗
    maxr,maxc = ws.max_row,len(headers)
    for r in range(1,maxr+1):
        for c in range(1,maxc+1):
            cell = ws.cell(r,c)
            cell.border = Border(
                left=thick if c==1 else cell.border.left,
                right=thick if c==maxc else cell.border.right,
                top=thick if r==1 else cell.border.top,
                bottom=thick if r==maxr else cell.border.bottom)

    # Y列Rp格式
    for r in range(3,maxr+1):
        ws.cell(r,25).number_format='Rp#,##0.00'

    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()

# ===== API =====
@app.get("/healthz")
def healthz(): return {"ok":True,"table":DB_TABLE}

@app.get("/download_excel")
def download_excel(
    DATE_STR: Optional[str]=None, DATE_FROM:Optional[str]=None,
    DATE_TO:Optional[str]=None, DATE_RANGE:Optional[str]=None,
    PROJECT_NAME:Optional[str]=None, COMPANY:Optional[str]=None):

    # 时间条件
    span=(None,None); single=False; where=[]; params=[]
    if DATE_STR:
        d=parse_date(DATE_STR)
        if d: where.append("DATE_STR=%s"); params.append(d.strftime("%Y-%m-%d")); span=(d,d); single=True
    lf,rt=parse_date(DATE_FROM),parse_date(DATE_TO)
    if not (lf and rt) and DATE_RANGE: lf,rt=parse_range(DATE_RANGE)
    if lf and rt:
        where.append("DATE_STR BETWEEN %s AND %s")
        params+=[lf.strftime("%Y-%m-%d"),rt.strftime("%Y-%m-%d")]
        span=(lf,rt); single=(lf==rt)

    # 默认 GIJ 强制筛选
    if not PROJECT_NAME:
        where.append("PROJECT_NAME=%s"); params.append(DEFAULT_PROJECT)
        project=DEFAULT_PROJECT
    else:
        where.append("PROJECT_NAME=%s"); params.append(PROJECT_NAME)
        project=PROJECT_NAME

    if COMPANY: where.append("COMPANY=%s"); params.append(COMPANY)
    where_sql=" WHERE "+ " AND ".join(where)
    sql=f"SELECT * FROM `{DB_TABLE}`{where_sql} ORDER BY DATE_STR ASC,ID ASC"

    try:
        conn=connect_db()
        with conn.cursor() as cur:
            cur.execute(f"SHOW COLUMNS FROM `{DB_TABLE}`")
            headers=[r["Field"] for r in cur.fetchall()][:31]
            cur.execute(sql,params)
            rows=cur.fetchall() or []
    except Exception as e:
        return JSONResponse(status_code=500,content={"error":str(e)})
    finally:
        try: conn.close()
        except: pass

    if not rows:
        return JSONResponse(status_code=404,content={"error":"No data found for given filters"})

    excel=make_excel(rows,headers,project,span,single)

    # 文件名
    date_tag=(span[0].strftime("%Y%m%d") if single and span[0]
              else f"{span[0].strftime('%Y%m%d')}-{span[1].strftime('%Y%m%d')}" if span[0] and span[1]
              else dt.datetime.now().strftime("%Y%m%d"))
    name=f"{project}_AI_REPORT_{date_tag}.xlsx"
    ascii_name=re.sub(r"[^\x00-\x7F]+","_",name)
    utf8_name=quote(name)
    headers={"Content-Disposition":f"attachment; filename={ascii_name}; filename*=UTF-8''{utf8_name}"}

    return StreamingResponse(io.BytesIO(excel),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers)

if __name__=="__main__":
    uvicorn_run("main:app",host="0.0.0.0",port=int(os.getenv("PORT",8000)),reload=True)
