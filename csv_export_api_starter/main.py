#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, io, re, datetime as dt
from typing import List, Dict, Any, Optional, Tuple
from urllib.parse import quote, urlencode

import pymysql
from fastapi import FastAPI, Query, Request
from fastapi.responses import StreamingResponse, JSONResponse, RedirectResponse, HTMLResponse
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

# 错误跳转域名（支持环境变量覆盖）
REDIRECT_BEFORE_CUTOFF = os.getenv("REDIRECT_BEFORE_CUTOFF", "http://www.info.julongairoxy.com/")
REDIRECT_TODAY_FUTURE  = os.getenv("REDIRECT_TODAY_FUTURE",  "http://www.info2.julongairoxy.com/")
REDIRECT_OTHER_ERROR   = os.getenv("REDIRECT_OTHER_ERROR",   "http://www.info1.julongairoxy.com/")
CUTOFF_DATE_STR = os.getenv("CUTOFF_DATE", "2025-09-01")
CUTOFF_DATE = dt.datetime.strptime(CUTOFF_DATE_STR, "%Y-%m-%d").date()

app = FastAPI(title="Yaoguang Excel Download API", version="1.6.0")

# =========================
# 列定义
# =========================
COLUMN_SPECS = [
    ("ID","序号",["ID"]),
    ("DEVICE_ID","工时通编号",["DEVICE_ID"]),
    ("PROJECT_NAME","系统项目名称",["PROJECT_NAME"]),
    ("MECHANICAL_NO","系统编号",["MECHANICAL_NO"]),
    ("CAR_TYPE","系统类型",["CAR_TYPE"]),
    ("DATE_STR","数据日期",["DATE_STR"]),
    ("RENT_TYPE","系统计算类型",["RENT_TYPE"]),
    ("VALID_DURATION","有效工时",["VALID_DURATION"]),
    ("IDLING_DURATION","怠速工时",["IDLING_DURATION"]),
    ("VALID_PERCENT","工时有效比%",["VALID_PERCENT"]),
    ("DAY_OIL","油耗",["DAY_OIL"]),
    ("DAY_REFUEL","加油",["DAY_REFUEL"]),
    ("DAY_MILEAGE","里程",["DAY_MILEAGE"]),
    ("WORKHOUR_AVG_OIL","工时平均油耗",["WORKHOUR_AVG_OIL"]),
    ("TRANSPORT_AVG_OIL","运输平均油耗",["TRANSPORT_AVG_OIL"]),
    ("COMPANY","公司",["COMPANY"]),
    ("ESTATE","区域",["ESTATE"]),
    ("MACHINE_TYPE","机械类型",["MACHINE_TYPE"]),
    ("MACHINE_CATEGORY","机械具体类型",["MACHINE_CATEGORY"]),
    ("MACHINE_NO","园区编号",["MACHINE_NO"]),
    ("BRAND_SPEC","具体型号",["BRAND_SPEC"]),
    ("ORDER_NO SAP","SAP订单号",["ORDER_NO SAP","ORDER_NO_SAP","ORDER_NO"]),
    ("PURCHASE_DATE","购买日期",["PURCHASE_DATE"]),
    ("DRIVER_COUNT","司机数量",["DRIVER_COUNT"]),
    ("PURCH_PRICE","购买价格",["PURCH_PRICE","PURCHASE_PRICE"]),
    ("FUEL_DIFF","标准油耗差",["FUEL_DIFF"]),
    ("INSERT_TIME","数据插入时间",["INSERT_TIME","CREATE_TIME"]),
    ("SCORE","得分",["SCORE"]),
    ("SUMMARY","AI分析",["SUMMARY"]),
    ("ANALYSIS_TIME","分析时间",["ANALYSIS_TIME"]),
    ("MODEL_NAME","使用模型",["MODEL_NAME"]),
]
FIELD_TO_COLIDX = {f:i+1 for i,(f,_,_) in enumerate(COLUMN_SPECS)}
COLIDX_Y  = FIELD_TO_COLIDX["PURCH_PRICE"]
COLIDX_AC = FIELD_TO_COLIDX["SUMMARY"]

# =========================
# 工具函数（日期/跳转）
# =========================
def connect_db():
    return pymysql.connect(
        host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASS,
        database=DB_NAME, charset="utf8mb4", cursorclass=pymysql.cursors.DictCursor
    )

def parse_date_like(s: Optional[str]) -> Optional[dt.date]:
    if not s: return None
    for fmt in ("%Y-%m-%d","%Y.%m.%d","%Y/%m/%d"):
        try: return dt.datetime.strptime(s.strip(), fmt).date()
        except: pass
    m = re.match(r"^\s*(\d{4})[-./](\d{1,2})[-./](\d{1,2})\s*$", s or "")
    if m:
        y,mo,d = map(int, m.groups()); return dt.date(y,mo,d)
    return None

def parse_date_range(raw: Optional[str]) -> Tuple[Optional[dt.date], Optional[dt.date]]:
    if not raw: return None, None
    s = raw.strip().replace("—","-").replace("–","-").replace("到","-").replace("to","-").replace("TO","-")
    s = re.sub(r"-{2,}", "-", s)
    parts = [p for p in s.split("-") if p.strip()]
    if len(parts) >= 2:
        return parse_date_like(parts[0]), parse_date_like(parts[-1])
    return None, None

def span_from_params(DATE_STR: Optional[str], DATE_FROM: Optional[str], DATE_TO: Optional[str], DATE_RANGE: Optional[str]) -> Tuple[Optional[dt.date], Optional[dt.date]]:
    """从参数推导[min_date, max_date]；单日返回(d,d)。解析失败返回(None,None)"""
    if DATE_STR:
        d = parse_date_like(DATE_STR)
        return (d, d) if d else (None, None)
    left = parse_date_like(DATE_FROM) if DATE_FROM else None
    right = parse_date_like(DATE_TO) if DATE_TO else None
    if not (left and right) and DATE_RANGE:
        l2, r2 = parse_date_range(DATE_RANGE)
        left, right = left or l2, right or r2
    if left and right:
        if left > right: left, right = right, left
        return left, right
    return None, None

def choose_redirect_by_span(span: Tuple[Optional[dt.date], Optional[dt.date]]) -> str:
    """仅用于错误时的跳转域名选择"""
    today = dt.date.today()
    s, e = span
    # 规则优先级：包含 < 2025-09-01 其一 ⇒ info；否则若包含 >= 今天 其一 ⇒ info2；否则 ⇒ info1
    if (s and s < CUTOFF_DATE) or (e and e < CUTOFF_DATE):
        return REDIRECT_BEFORE_CUTOFF
    if (s and s >= today) or (e and e >= today):
        return REDIRECT_TODAY_FUTURE
    return REDIRECT_OTHER_ERROR

def want_html(request: Request) -> bool:
    accept = (request.headers.get("accept") or "").lower()
    ua = (request.headers.get("user-agent") or "").lower()
    return "text/html" in accept or "mozilla" in ua

def redirect_with_error(msg: str, extra: dict, span: Tuple[Optional[dt.date], Optional[dt.date]], status_code: int = 303):
    params = {"err": msg}
    for k in ["DATE_STR","DATE_FROM","DATE_TO","DATE_RANGE","PROJECT_NAME"]:
        v = extra.get(k)
        if v: params[k] = v
    base = choose_redirect_by_span(span)
    return RedirectResponse(url=f"{base}?{urlencode(params, safe='')}", status_code=status_code)

# =========================
# Excel 渲染
# =========================
def detect_date_span_from_rows(rows: List[Dict[str, Any]]) -> Tuple[Optional[dt.date], Optional[dt.date]]:
    dates = []
    for r in rows:
        v = r.get("DATE_STR")
        if isinstance(v, dt.datetime): dates.append(v.date())
        elif isinstance(v, dt.date):   dates.append(v)
        elif isinstance(v, str):
            d = parse_date_like(v);  d and dates.append(d)
    return (min(dates), max(dates)) if dates else (None, None)

def col_widths_spec() -> List[float]:
    return [9,15,28,20,23,14,14,18,18,15,12,13,13,21,21,13,12,15,21,20,40,18,19,16,20,13,21,9,200,20,15]

def to_period_str(span: Tuple[Optional[dt.date], Optional[dt.date]], single_day: bool) -> str:
    if span[0] and span[1]:
        return f"（{span[0].strftime('%Y.%m.%d')}）" if single_day else f"（{span[0].strftime('%Y.%m.%d')}-{span[1].strftime('%Y.%m.%d')}）"
    return "（）"

def _first_present(row: Dict[str, Any], alts: list) -> Any:
    for k in alts:
        if k in row and row[k] is not None:
            return row[k]
    return ""

def make_excel(rows: List[Dict[str, Any]], project_name: str, span: Tuple[Optional[dt.date], Optional[dt.date]], single_day: bool) -> bytes:
    wb = Workbook(); ws = wb.active; ws.title = "AI Report"
    for idx, w in enumerate(col_widths_spec(), start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=31)
    if not (span[0] and span[1]) and rows:
        span = detect_date_span_from_rows(rows); single_day = (span[0] and span[1] and span[0]==span[1])
    c = ws.cell(row=1, column=1)
    c.value = f"{project_name}CATATAN ANALISIS OTOMATIS AI DARI DATA KENDARAAN\n车联网数据AI自动分析记录\n{to_period_str(span, single_day)}"
    c.font = Font(name=FONT_TITLE_NAME, size=26, bold=True)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 100

    ws.row_dimensions[2].height = 40
    head_font = Font(name=FONT_BODY_NAME, size=10, bold=True)
    head_align= Alignment(horizontal="center", vertical="center", wrap_text=True)
    for j,(f,zh,_) in enumerate(COLUMN_SPECS, start=1):
        cell = ws.cell(row=2, column=j, value=f"{f}\n{zh}")
        cell.font=head_font; cell.alignment=head_align

    body_font = Font(name=FONT_BODY_NAME, size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    top_left = Alignment(horizontal="left",   vertical="top",   wrap_text=True)
    thin, thick = Side(style="thin", color="000"), Side(style="thick", color="000")
    border_all = Border(left=thin,right=thin,top=thin,bottom=thin)

    for r_idx, row in enumerate(rows, start=3):
        ws.row_dimensions[r_idx].height = 40
        for c_idx, (_,_,alts) in enumerate(COLUMN_SPECS, start=1):
            v = _first_present(row, alts)
            cell = ws.cell(row=r_idx, column=c_idx, value=v)
            cell.font = body_font
            cell.alignment = top_left if c_idx==COLIDX_AC else center
            cell.border = border_all

    max_row = max(2, 2+len(rows)); max_col = 31
    for r in range(1, max_row+1):
        for c in range(1, max_col+1):
            cell = ws.cell(row=r, column=c)
            left   = thick if c==1       else cell.border.left
            right  = thick if c==max_col else cell.border.right
            top    = thick if r==1       else cell.border.top
            bottom = thick if r==max_row else cell.border.bottom
            cell.border = Border(left=left,right=right,top=top,bottom=bottom)

    for r in range(3, max_row+1):
        ws.cell(row=r, column=COLIDX_Y).number_format = 'Rp#,##0.00'

    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio.read()

# =========================
# 可选：本地简易表单
# =========================
@app.get("/", response_class=HTMLResponse)
@app.get("/ui", response_class=HTMLResponse)
def ui(error: Optional[str] = None):
    return """
<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>AI 报表下载</title></head>
<body style="font-family:Segoe UI,Roboto,Arial,'微软雅黑';max-width:880px;margin:40px auto;padding:0 16px">
<h1>车联网数据 AI 报表下载</h1>
<p>填写日期与项目名后点击下载。只下载单日时，仅填“单日”。</p>
<div>
<label>单日：</label><input id="DATE_STR" placeholder="YYYY-MM-DD">
<label style="margin-left:12px">开始：</label><input id="DATE_FROM" placeholder="YYYY-MM-DD">
<label style="margin-left:12px">结束：</label><input id="DATE_TO" placeholder="YYYY-MM-DD">
</div>
<div style="margin-top:8px">
<label>项目名：</label><input id="PROJECT_NAME" placeholder="默认 Kalteng GIJ 中加一园" style="width:60%">
</div>
<button onclick="go()" style="margin-top:10px">下载 Excel</button>
<script>
function enc(s){return encodeURIComponent((s||'').trim());}
function go(){
  const d=document.getElementById('DATE_STR').value;
  const f=document.getElementById('DATE_FROM').value;
  const t=document.getElementById('DATE_TO').value;
  const p=document.getElementById('PROJECT_NAME').value;
  let qs=[];
  if(d){qs.push('DATE_STR='+enc(d));}
  else if(f&&t){qs.push('DATE_FROM='+enc(f));qs.push('DATE_TO='+enc(t));}
  else{alert('请填写 单日 或 起止日期');return;}
  if(p) qs.push('PROJECT_NAME='+enc(p));
  location.href='/download_excel?'+qs.join('&');
}
</script></body></html>
"""

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
    request: Request,
    DATE_STR: Optional[str] = Query(default=None),
    DATE_FROM: Optional[str] = Query(default=None),
    DATE_TO: Optional[str]   = Query(default=None),
    DATE_RANGE: Optional[str]= Query(default=None),
    PROJECT_NAME: Optional[str] = Query(default=None),
):
    PROJECT_NAME = (PROJECT_NAME.strip() or None) if PROJECT_NAME is not None else None
    DATE_STR   = DATE_STR.strip() if DATE_STR else None
    DATE_FROM  = DATE_FROM.strip() if DATE_FROM else None
    DATE_TO    = DATE_TO.strip() if DATE_TO else None
    DATE_RANGE = DATE_RANGE.strip() if DATE_RANGE else None

    # 用于错误跳转决策的参数跨度
    input_span = span_from_params(DATE_STR, DATE_FROM, DATE_TO, DATE_RANGE)

    span = (None, None); single_day = False
    where, params = [], []

    if DATE_STR:
        d = parse_date_like(DATE_STR)
        if not d:
            msg = "DATE_STR 格式不合法，应为 YYYY-MM-DD / YYYY.MM.DD / YYYY/MM/DD"
            return redirect_with_error(msg,
                {"DATE_STR": DATE_STR, "PROJECT_NAME": PROJECT_NAME}, input_span) if want_html(request) \
                else JSONResponse(status_code=400, content={"error": msg})
        where.append("DATE_STR = %s"); params.append(d.strftime("%Y-%m-%d"))
        span, single_day = (d, d), True
    else:
        left = parse_date_like(DATE_FROM) if DATE_FROM else None
        right = parse_date_like(DATE_TO) if DATE_TO else None
        if not (left and right) and DATE_RANGE:
            l2, r2 = parse_date_range(DATE_RANGE)
            left, right = left or l2, right or r2
        if left and right:
            where.append("DATE_STR BETWEEN %s AND %s")
            params += [left.strftime("%Y-%m-%d"), right.strftime("%Y-%m-%d")]
            span, single_day = (left, right), (left == right)
        else:
            msg = "时间必填：传 DATE_STR 或 DATE_FROM+DATE_TO 或 DATE_RANGE"
            return redirect_with_error(msg,
                {"DATE_FROM": DATE_FROM, "DATE_TO": DATE_TO, "DATE_RANGE": DATE_RANGE, "PROJECT_NAME": PROJECT_NAME},
                input_span) if want_html(request) \
                else JSONResponse(status_code=400, content={"error": msg})

    project_for_title = PROJECT_NAME or DEFAULT_PROJECT
    where.append("PROJECT_NAME = %s"); params.append(project_for_title)

    sql = f"SELECT * FROM `{DB_TABLE}` WHERE " + " AND ".join(where) + " ORDER BY DATE_STR ASC, ID ASC"

    try:
        conn = connect_db()
        with conn.cursor() as cur:
            cur.execute(sql, params)
            rows = cur.fetchall() or []
    except Exception as e:
        msg = f"DB error: {e}"
        return redirect_with_error(msg,
            {"DATE_STR": DATE_STR, "DATE_FROM": DATE_FROM, "DATE_TO": DATE_TO, "DATE_RANGE": DATE_RANGE, "PROJECT_NAME": PROJECT_NAME},
            input_span) if want_html(request) \
            else JSONResponse(status_code=500, content={"error": msg})
    finally:
        try: conn.close()
        except: pass

    if not rows:
        msg = "没有匹配数据"
        return redirect_with_error(msg,
            {"DATE_STR": DATE_STR, "DATE_FROM": DATE_FROM, "DATE_TO": DATE_TO, "DATE_RANGE": DATE_RANGE, "PROJECT_NAME": PROJECT_NAME},
            input_span) if want_html(request) \
            else JSONResponse(status_code=404, content={"error": msg})

    excel_bytes = make_excel(rows, project_for_title, span, single_day)

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
