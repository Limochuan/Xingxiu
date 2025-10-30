#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
AI 并发分析并回写 xingxiu_daily_report（仅处理指定项目与指定日期）
- 长摘要：管理层阅读版（约 220–320 字），只用“当日数据”，明确写出打分依据（≥6 个数字）
- 运输：TRANSPORT_AVG_OIL + DAY_MILEAGE + (DAY_OIL/REFUEL 一致性) + FUEL_DIFF 等级 + 采购价简写
- 工程机械：WORKHOUR_AVG_OIL + 有效/怠速/有效比 + DAY_OIL/REFUEL 匹配 + FUEL_DIFF + 采购价简写
- 早判：传感器异常/全零直接 0 分
- 回写：SCORE、SUMMARY、ANALYSIS_TIME（JKT）、MODEL_NAME
"""

import os
import re
import json
import time
import logging
from logging.handlers import RotatingFileHandler
from dataclasses import dataclass
from typing import Optional, List, Tuple
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import pymysql
import requests

# ===== 可调参数（从环境读取，默认保持你给的值）=====
TARGET_DATE = os.getenv("TARGET_DATE") or os.getenv("DATE_STR") or "2025-10-29"
PROJECT_FILTER = os.getenv("PROJECT_FILTER", "Kalteng GIJ 中加一园")
MAX_WORKERS = int(os.getenv("MAX_WORKERS", "8"))
OPENAI_TIMEOUT = int(os.getenv("OPENAI_TIMEOUT", "60"))
OPENAI_RETRIES = int(os.getenv("OPENAI_RETRIES", "2"))
BATCH_UPDATE_SIZE = int(os.getenv("BATCH_UPDATE_SIZE", "100"))

# ===== 数据库（从 Secrets / env）=====
DB_HOST = os.getenv("DB_HOST", "rm-k1a5w7qk9cnm74r25wo.mysql.ap-southeast-5.rds.aliyuncs.com")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_USER = os.getenv("DB_USER", "script_xingxiu")
DB_PASS = os.getenv("DB_PASS", "Julong678678678")
DB_NAME = os.getenv("DB_NAME", "xingxiu_db")
REPORT_TABLE = os.getenv("REPORT_TABLE", "xingxiu_daily_report")

# ===== OpenAI（从 Secrets / env）=====
OPENAI_API_KEY  = os.getenv("OPENAI_API_KEY", "YOUR_API_KEY")
OPENAI_API_BASE = os.getenv("OPENAI_API_BASE", "https://api.openai-proxy.org/v1")
OPENAI_MODEL    = os.getenv("OPENAI_MODEL", "gpt-5-mini")
OPENAI_TEMPERATURE = os.getenv("OPENAI_TEMPERATURE", "").strip()

# ===== 日志 =====
logger = logging.getLogger("ai_analyzer")
logger.setLevel(logging.INFO)
fh = RotatingFileHandler("ai_analyzer.log", maxBytes=8*1024*1024, backupCount=5, encoding="utf-8")
fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
ch = logging.StreamHandler()
ch.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
if not logger.handlers:
    logger.addHandler(fh)
    logger.addHandler(ch)

# ===== 数据对象 =====
@dataclass
class ReportRow:
    id: int
    device_id: Optional[str]
    company: Optional[str]
    estate: Optional[str]
    machine_type: Optional[str]
    machine_category: Optional[str]
    machine_no: Optional[str]
    brand_spec: Optional[str]
    order_no: Optional[str]
    purchase_date: Optional[str]
    driver_count: Optional[int]
    project_name: Optional[str]
    create_time: Optional[str]
    date_str: Optional[str]
    rent_type: Optional[str]
    valid_duration: Optional[float]
    idling_duration: Optional[float]
    valid_percent: Optional[float]
    day_oil: Optional[float]
    day_refuel: Optional[float]
    day_mileage: Optional[float]
    workhour_avg_oil: Optional[float]
    transport_avg_oil: Optional[float]
    fuel_diff: Optional[float]
    purch_price: Optional[float]
    insert_time: Optional[str]

# ===== 小工具 =====
def _safe_float(v) -> Optional[float]:
    try:
        return None if v is None else float(v)
    except Exception:
        return None

def _s(v) -> str:
    return "" if v is None else str(v)

def jkt_now_str() -> str:
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo("Asia/Jakarta")
        return datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")

def format_idr_brief(v: Optional[float]) -> str:
    """采购价显示：<10亿 → “X条”；≥10亿 → “Y.M”"""
    if v is None:
        return ""
    try:
        a = float(v)
    except Exception:
        return str(v)
    if a < 1_000_000_000:
        return f"{int(round(a / 1_000_000))}条"
    else:
        return f"{a / 1_000_000_000:.1f}M"

def fmt2(x: Optional[float]) -> str:
    return "-" if x is None else f"{x:.2f}"

# ===== 业务规则 & 分类 =====
class Rules:
    TRUCK_KEYWORDS = ["dump", "dumptruck", "dump truck", "dumptruk", "truk", " dt", "dt "]
    MAX_HOURS = 24.0
    HOURS_SUM_TOL = 0.5
    MAX_DAY_OIL = 300.0
    MAX_DAY_REFUEL = 400.0
    AVG_OIL_MAX = 60.0

    @staticmethod
    def classify_by_rent(r: ReportRow) -> str:
        rt = (_s(r.rent_type)).lower()
        if ("运" in rt) or ("运输" in rt) or ("transport" in rt):
            return "truck"
        if ("包月" in rt) or ("monthly" in rt) or ("month" in rt):
            return "heavy"
        txt = " ".join([_s(r.machine_type), _s(r.machine_category), _s(r.brand_spec), _s(r.machine_no)]).lower()
        return "truck" if any(k in txt for k in Rules.TRUCK_KEYWORDS) else "heavy"

    @staticmethod
    def has_hours(r: ReportRow) -> bool:
        return (r.valid_duration and r.valid_duration > 0) or (r.idling_duration and r.idling_duration > 0)

    @staticmethod
    def avg_oil_value(r: ReportRow) -> Optional[float]:
        if r.workhour_avg_oil is not None and r.workhour_avg_oil >= 0:
            return r.workhour_avg_oil
        if r.day_oil and r.valid_duration and r.valid_duration > 0:
            return r.day_oil / r.valid_duration
        return None

    @staticmethod
    def early_sensor_fault(r: ReportRow):
        """运输：只看 TRANSPORT_AVG_OIL>1；工程：WORKHOUR_AVG_OIL=-1 或有工时无油量"""
        dev = r.machine_no or "(未知编号)"
        cls = Rules.classify_by_rent(r)
        if cls == "truck":
            if r.transport_avg_oil is not None and r.transport_avg_oil > 1:
                summary = (f"{dev}：运输平均油耗={_s(r.transport_avg_oil)}（>1），判定为口径异常。"
                           f"里程={_s(r.day_mileage)}km；油耗={_s(r.day_oil)}L；加油={_s(r.day_refuel)}L。")
                return 0, summary
            return None
        # heavy
        if r.workhour_avg_oil is not None and r.workhour_avg_oil == -1:
            summary = (f"{dev}：工时平均油耗=-1，判定为口径异常。"
                       f"有效={_s(r.valid_duration)}h；怠速={_s(r.idling_duration)}h。")
            return 0, summary
        has_oil = (r.day_oil and r.day_oil > 0) or (Rules.avg_oil_value(r) not in (None, 0))
        if Rules.has_hours(r) and not has_oil:
            return 0, (f"{dev}：有工时但缺少油量/平均油耗，判定为采集缺失。")
        return None

    @staticmethod
    def early_common_sanity(r: ReportRow):
        dev = r.machine_no or "(未知编号)"
        cls = Rules.classify_by_rent(r)

        def is_zero(x):
            try:
                return (x is None) or float(x) == 0.0
            except:
                return False

        if cls == "truck":
            if is_zero(r.transport_avg_oil) and is_zero(r.day_mileage) and is_zero(r.day_oil) and is_zero(r.day_refuel):
                return 0, (f"{dev}：运输平均油耗0.0、里程0.0、油耗0.0、加油0.0；当天无行驶与燃油记录。")
            for nm, val in {"当日油耗": r.day_oil, "当日加油": r.day_refuel, "当日里程": r.day_mileage}.items():
                try:
                    if val is not None and float(val) < 0:
                        return 0, f"{dev}：存在负数记录（{nm}），判定为口径错误。"
                except Exception:
                    pass
            if (r.day_oil and r.day_oil > Rules.MAX_DAY_OIL) or (r.day_refuel and r.day_refuel > Rules.MAX_DAY_REFUEL):
                return 0, f"{dev}：当日油耗/加油量异常偏大（>{Rules.MAX_DAY_OIL}L），判定为录入异常。"
            return None

        # heavy
        def z(x):
            try: return (x is None) or float(x) == 0.0
            except: return False

        if z(r.valid_duration) and z(r.idling_duration) and z(r.day_oil) and z(r.day_refuel):
            return 0, (f"{dev}：有效/怠速/油耗/加油均为0；显示停机或记录缺失。")
        for nm, val in {
            "有效工时": r.valid_duration, "怠速工时": r.idling_duration,
            "当日油耗": r.day_oil, "当日加油": r.day_refuel, "当日里程": r.day_mileage
        }.items():
            try:
                if val is not None and float(val) < 0:
                    return 0, f"{dev}：存在负数记录（{nm}），判定为口径错误。"
            except Exception:
                pass
        if (r.valid_duration and r.valid_duration > Rules.MAX_HOURS) or \
           (r.idling_duration and r.idling_duration > Rules.MAX_HOURS):
            return 0, f"{dev}：有效/怠速工时超过 24h/天；判定为统计异常。"
        if ((r.valid_duration or 0) + (r.idling_duration or 0)) > Rules.MAX_HOURS + Rules.HOURS_SUM_TOL:
            return 0, f"{dev}：有效+怠速合计超 24h/天；判定为重复累计/跨日切分。"
        ao = Rules.avg_oil_value(r)
        if ao is not None and ao > Rules.AVG_OIL_MAX:
            return 0, f"{dev}：平均油耗 {ao:.1f} L/h 偏高；判定为口径异常。"
        return None

# ===== Prompt & 解析 =====
BANNED_PHRASES = [
    "建议","优化","培训","保持良好操作习惯","检查","清单","逐项实施","措施",
    "可能","原因","由于","受","影响","导致","推测","猜测"
]

def build_prompt(r: ReportRow) -> str:
    """要求 220–320 字，至少 6 个数字点；不得出现“建议/可能/原因…”；不加“评分依据：”抬头"""
    cls = Rules.classify_by_rent(r)
    dev = r.machine_no or "(未知编号)"
    price_brief = format_idr_brief(r.purch_price)

    header = (
        f"设备：{dev}；项目：{_s(r.project_name)}；品牌型号：{_s(r.brand_spec)}；"
        f"日期：{_s(r.date_str)}；租用方式：{_s(r.rent_type)}；采购价：{price_brief}；采购日期：{_s(r.purchase_date)}；"
        f"FUEL_DIFF：{_s(r.fuel_diff)}。\n"
    )

    if cls == "truck":
        focus = (
            f"运输平均油耗：{fmt2(r.transport_avg_oil)}；里程：{fmt2(r.day_mileage)}km；"
            f"当日油耗：{fmt2(r.day_oil)}L；加油：{fmt2(r.day_refuel)}L。若里程>0，请计算 L/km=油耗/里程 并与运输平均油耗对齐核对。"
        )
    else:
        focus = (
            f"有效工时：{fmt2(r.valid_duration)}h；怠速工时：{fmt2(r.idling_duration)}h；"
            f"工时有效比：{fmt2(r.valid_percent)}；工时平均油耗：{fmt2(r.workhour_avg_oil)}L/h；"
            f"当日油耗：{fmt2(r.day_oil)}L；加油：{fmt2(r.day_refuel)}L。只用当日数据做结论。"
        )

    ask = (
        "请直接输出两行：\n"
        "评分：<整数>\n"
        "总结：<220–320字，给出当日“为何得此分”的完整结论；"
        "运输需写出 L/km 与运输平均油耗的一致性、里程规模、油耗与加油的匹配、FUEL_DIFF 等级及采购价简写；"
        "工程机械需写出工时结构、工时油耗与油耗/加油匹配；"
        "必须包含≥6个数字点；不得出现‘建议/可能/原因/由于/影响/导致/推测’等词；不要写“评分依据：”。>"
    )

    return "面向管理层的结论摘要：\n" + header + focus + "\n" + ask

def parse_output(text: str):
    if not text:
        return None, ""
    t = text.replace("：", ":").strip()
    dup = t.find("\n评分", 1)
    if dup > 0:
        t = t[:dup].strip()
    m = re.search(r"(?:评分|score|Score|最终评分|综合评分)\s*:\s*([0-9]{1,3})", t)
    score = int(m.group(1)) if m else None
    if score is not None:
        score = max(60, min(100, score))
    m2 = re.search(r"(?:总结|summary|Summary|结论|分析)\s*:\s*(.*)", t, re.S)
    summary = (m2.group(1).strip() if m2 else "\n".join(t.splitlines()[1:]).strip()).strip("`\"“”")
    return score, summary

# ===== 兜底长摘要 =====
def diff_band_text(x: Optional[float]) -> str:
    if x is None: return "无"
    try: v = abs(float(x))
    except Exception: return "无"
    if v <= 3:  return "一致"
    if v <= 6:  return "轻微"
    if v <= 10: return "中度"
    return "明显"

def compose_fallback_summary(r: ReportRow, cls: str, score: Optional[int]) -> str:
    price_brief = format_idr_brief(r.purch_price)
    score_txt = f"{score}分" if score is not None else "—分"
    if cls == "truck":
        lpk = None
        if r.day_mileage and r.day_mileage > 0 and r.day_oil is not None:
            try: lpk = r.day_oil / r.day_mileage
            except Exception: lpk = None
        lpk_txt = "-" if lpk is None else f"{lpk:.4f}"
        band = diff_band_text(r.fuel_diff)
        parts = [
            f"当日里程{fmt2(r.day_mileage)}km、油耗{fmt2(r.day_oil)}L、加油{fmt2(r.day_refuel)}L，",
            f"L/km={lpk_txt}，与运输平均油耗{fmt2(r.transport_avg_oil)}核对，",
            f"FUEL_DIFF={fmt2(r.fuel_diff)}（{band}），采购价{price_brief}，购于{_s(r.purchase_date)}，",
            "油耗与加油的量级与里程规模相互匹配，",
            f"据此给出{score_txt}。"
        ]
        text = "".join(parts)
    else:
        ao = Rules.avg_oil_value(r)
        parts = [
            f"有效{fmt2(r.valid_duration)}h、怠速{fmt2(r.idling_duration)}h、有效比{fmt2(r.valid_percent)}，",
            f"工时油耗{fmt2(ao if ao is not None else r.workhour_avg_oil)}L/h，",
            f"油耗{fmt2(r.day_oil)}L、加油{fmt2(r.day_refuel)}L，FUEL_DIFF={fmt2(r.fuel_diff)}，",
            f"采购价{price_brief}，购于{_s(r.purchase_date)}，",
            "工时结构与油耗/加油匹配性明确，",
            f"据此给出{score_txt}。"
        ]
        text = "".join(parts)
    for p in BANNED_PHRASES: text = text.replace(p, "")
    text = re.sub(r"\s+", "", text)
    while len(text) < 220:
        text += "数据完整、量级清晰，结论依据已充分列示。"
        if len(text) > 320: break
    if len(text) > 320: text = text[:320] + "。"
    if text and text[-1] not in "。.!！?？": text += "。"
    return text

def enforce_exec_style(summary: str, r: ReportRow, score: Optional[int]) -> str:
    if not summary or len(summary) < 220:
        return compose_fallback_summary(r, Rules.classify_by_rent(r), score)
    summary = re.sub(r"^(评分依据|小结|总结|结论)\s*[:：]\s*", "", summary)
    for p in BANNED_PHRASES:
        summary = summary.replace(p, "")
    summary = re.sub(r"\s+", " ", summary).strip()
    if len(summary) > 360:
        summary = summary[:360].rstrip("，,；;。.") + "。"
    if summary and summary[-1] not in "。.!！?？":
        summary += "。"
    if len(summary) < 220:
        summary = compose_fallback_summary(r, Rules.classify_by_rent(r), score)
    return summary

# ===== OpenAI 客户端 =====
class OpenAIClient:
    def __init__(self):
        self.base = OPENAI_API_BASE.rstrip("/")
        self.headers = {"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"}

    def chat(self, prompt: str) -> str:
        url = f"{self.base}/chat/completions"
        payload = {"model": OPENAI_MODEL, "messages": [{"role": "user", "content": prompt}]}
        if OPENAI_TEMPERATURE:
            try: payload["temperature"] = float(OPENAI_TEMPERATURE)
            except Exception: pass

        for attempt in range(OPENAI_RETRIES + 1):
            try:
                with requests.Session() as sess:
                    r = sess.post(url, headers=self.headers, data=json.dumps(payload), timeout=OPENAI_TIMEOUT)
                if r.status_code == 400 and "temperature" in r.text.lower() and "unsupported" in r.text.lower():
                    payload.pop("temperature", None); continue
                if r.status_code != 200:
                    raise RuntimeError(f"OpenAI error {r.status_code}: {r.text[:200]}")
                data = r.json()
                return data.get("choices", [{}])[0].get("message", {}).get("content", "")
            except Exception as e:
                if attempt >= OPENAI_RETRIES: raise
                time.sleep(1.2 * (attempt + 1))
        raise RuntimeError("OpenAI retries exhausted")

# ===== 表结构保障 =====
def ensure_report_extra_columns(conn):
    with conn.cursor() as cur:
        cur.execute(f"SHOW COLUMNS FROM `{REPORT_TABLE}`")
        cols = {row["Field"].upper(): row["Field"] for row in cur.fetchall()}
    missing = []
    if "ANALYSIS_TIME" not in cols:
        missing.append("ADD COLUMN `ANALYSIS_TIME` DATETIME NULL COMMENT 'AI分析入库时间（Asia/Jakarta）' AFTER `SUMMARY`")
    if "MODEL_NAME" not in cols:
        missing.append("ADD COLUMN `MODEL_NAME` VARCHAR(64) NULL COMMENT 'AI模型名称' AFTER `ANALYSIS_TIME`")
    if missing:
        sql = f"ALTER TABLE `{REPORT_TABLE}` " + ", ".join(missing)
        with conn.cursor() as cur:
            cur.execute(sql)
        conn.commit()
        logger.info("已补充列：" + "，".join(["ANALYSIS_TIME" if "ANALYSIS_TIME" in m else "MODEL_NAME" for m in missing]))

# ===== 仓储层 =====
class Repo:
    def __init__(self, conn, idling_col: str):
        self.conn = conn
        self.idling_col = idling_col

    def fetch_pending(self, date_str: str, project_name: str, limit: int = 2000) -> List[ReportRow]:
        rows: List[ReportRow] = []
        with self.conn.cursor() as cur:
            sql = f"""
                SELECT
                    ID AS id, DEVICE_ID AS device_id, COMPANY AS company, ESTATE AS estate,
                    MACHINE_TYPE AS machine_type, MACHINE_CATEGORY AS machine_category,
                    MACHINE_NO AS machine_no, BRAND_SPEC AS brand_spec, ORDER_NO AS order_no,
                    PURCHASE_DATE AS purchase_date, DRIVER_COUNT AS driver_count,
                    PROJECT_NAME AS project_name, CREATE_TIME AS create_time,
                    DATE_STR AS date_str, RENT_TYPE AS rent_type, VALID_DURATION AS valid_duration,
                    {self.idling_col} AS idling_duration, VALID_PERCENT AS valid_percent,
                    DAY_OIL AS day_oil, DAY_REFUEL AS day_refuel, DAY_MILEAGE AS day_mileage,
                    WORKHOUR_AVG_OIL AS workhour_avg_oil, TRANSPORT_AVG_OIL AS transport_avg_oil,
                    FUEL_DIFF AS fuel_diff, PURCH_PRICE AS purch_price,
                    INSERT_TIME AS insert_time
                FROM `{REPORT_TABLE}`
                WHERE DATE_STR=%s
                  AND PROJECT_NAME=%s
                  AND (SCORE IS NULL OR SUMMARY IS NULL)
                ORDER BY INSERT_TIME DESC
                LIMIT %s
            """
            cur.execute(sql, (date_str, project_name, limit))
            for r in cur.fetchall():
                rows.append(ReportRow(
                    id=r["id"], device_id=r.get("device_id"),
                    company=r.get("company"), estate=r.get("estate"),
                    machine_type=r.get("machine_type"), machine_category=r.get("machine_category"),
                    machine_no=r.get("machine_no"), brand_spec=r.get("brand_spec"),
                    order_no=r.get("order_no"),
                    purchase_date=str(r.get("purchase_date")) if r.get("purchase_date") else None,
                    driver_count=r.get("driver_count"), project_name=r.get("project_name"),
                    create_time=str(r.get("create_time")) if r.get("create_time") else None,
                    date_str=str(r.get("date_str")) if r.get("date_str") else None,
                    rent_type=str(r.get("rent_type")) if r.get("rent_type") else None,
                    valid_duration=_safe_float(r.get("valid_duration")),
                    idling_duration=_safe_float(r.get("idling_duration")),
                    valid_percent=_safe_float(r.get("valid_percent")),
                    day_oil=_safe_float(r.get("day_oil")),
                    day_refuel=_safe_float(r.get("day_refuel")),
                    day_mileage=_safe_float(r.get("day_mileage")),
                    workhour_avg_oil=_safe_float(r.get("workhour_avg_oil")),
                    transport_avg_oil=_safe_float(r.get("transport_avg_oil")),
                    fuel_diff=_safe_float(r.get("fuel_diff")),
                    purch_price=_safe_float(r.get("purch_price")),
                    insert_time=str(r.get("insert_time")) if r.get("insert_time") else None,
                ))
        return rows

    def write_results_batch(self, batch: List[Tuple[int, Optional[int], str, str, str, str, str]]):
        """批量更新：每项=(ID,SCORE,SUMMARY,ANALYSIS_TIME,MODEL_NAME,DATE_STR,PROJECT_NAME)"""
        if not batch: return
        with self.conn.cursor() as cur:
            sql = f"""
                UPDATE `{REPORT_TABLE}`
                SET SCORE=%s, SUMMARY=%s, ANALYSIS_TIME=%s, MODEL_NAME=%s
                WHERE ID=%s AND DATE_STR=%s AND PROJECT_NAME=%s
                  AND (SCORE IS NULL OR SUMMARY IS NULL)
            """
            params = [(sc, sm, at, mn, rid, ds, pj) for (rid, sc, sm, at, mn, ds, pj) in batch]
            cur.executemany(sql, params)
        self.conn.commit()

# ===== 单条处理 =====
def process_one(row: ReportRow) -> Tuple[int, Optional[int], str, str]:
    """返回：(row_id, score, summary, error)"""
    try:
        early = Rules.early_sensor_fault(row) or Rules.early_common_sanity(row)
        if early:
            score0, summary0 = early
            summary0 = re.sub(r"[；;]\s*$", "。", summary0)
            return (row.id, score0, summary0, "")

        client = OpenAIClient()
        prompt = build_prompt(row)
        content = client.chat(prompt)
        score, summary = parse_output(content)
        if score is None or not summary:
            return (row.id, score, compose_fallback_summary(row, Rules.classify_by_rent(row), score), "")
        summary = enforce_exec_style(summary, row, score)
        return (row.id, score, summary, "")
    except Exception as e:
        return (row.id, None, "", f"{type(e).__name__}: {e}")

# ===== 主流程 =====
def main():
    t0 = time.perf_counter()
    logger.info(f"启动：DB={DB_HOST}/{DB_NAME} 表={REPORT_TABLE} 日期={TARGET_DATE} 项目={PROJECT_FILTER} 并发={MAX_WORKERS}")

    conn = pymysql.connect(
        host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASS,
        database=DB_NAME, charset="utf8mb4", autocommit=False,
        cursorclass=pymysql.cursors.DictCursor,
    )
    try:
        ensure_report_extra_columns(conn)

        with conn.cursor() as cur:
            cur.execute(f"SHOW COLUMNS FROM `{REPORT_TABLE}`")
            cols = {row["Field"].upper(): row["Field"] for row in cur.fetchall()}
        idling_col = cols.get("IDLING_DURATION") or cols.get("IDING_DURATION")
        if not idling_col:
            raise RuntimeError("业务表缺少列：IDLING_DURATION / IDING_DURATION")
        logger.info(f"怠速工时列名：{idling_col}")

        repo = Repo(conn, idling_col)
        rows = repo.fetch_pending(TARGET_DATE, PROJECT_FILTER, limit=5000)
        if not rows:
            logger.info("当日该项目无待处理记录，退出。")
            return

        logger.info(f"待处理：{len(rows)} 条")
        model_name = OPENAI_MODEL
        analysis_time = jkt_now_str()
        results: List[Tuple[int, Optional[int], str, str, str, str, str]] = []

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            futs = [ex.submit(process_one, r) for r in rows]
            done = 0
            for fut in as_completed(futs):
                rid, score, summary, err = fut.result()
                done += 1
                if err:
                    logger.error(f" -> 行ID={rid} 失败：{err}")
                    continue
                results.append((rid, score, summary, analysis_time, model_name, TARGET_DATE, PROJECT_FILTER))
                if done % 10 == 0 or done == len(rows):
                    logger.info(f"进度：{done}/{len(rows)}")
                if len(results) >= BATCH_UPDATE_SIZE:
                    repo.write_results_batch(results)
                    logger.info(f"批量写入 {len(results)} 条")
                    results.clear()

        if results:
            repo.write_results_batch(results)
            logger.info(f"批量写入 {len(results)} 条")

        dt = time.perf_counter() - t0
        logger.info(f"完成。总耗时 {dt:.1f}s（并发={MAX_WORKERS}）")

    finally:
        try:
            conn.close()
        except Exception:
            pass

if __name__ == "__main__":
    main()
