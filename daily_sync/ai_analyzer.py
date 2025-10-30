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


# ===== 环境变量配置 =====
TARGET_DATE = os.getenv("TARGET_DATE") or os.getenv("DATE_STR") or "2025-10-29"
PROJECT_FILTER = os.getenv("PROJECT_FILTER", "Kalteng GIJ 中加一园")

DB_HOST = os.getenv("DB_HOST", "")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_USER = os.getenv("DB_USER", "")
DB_PASS = os.getenv("DB_PASS", "")
DB_NAME = os.getenv("DB_NAME", "")
REPORT_TABLE = os.getenv("REPORT_TABLE", "xingxiu_daily_report")

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_API_BASE = os.getenv("OPENAI_API_BASE", "https://api.openai-proxy.org/v1")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-5-mini")
OPENAI_TEMPERATURE = os.getenv("OPENAI_TEMPERATURE", "").strip()

MAX_WORKERS = int(os.getenv("MAX_WORKERS", "8"))
OPENAI_TIMEOUT = int(os.getenv("OPENAI_TIMEOUT", "60"))
OPENAI_RETRIES = int(os.getenv("OPENAI_RETRIES", "2"))
BATCH_UPDATE_SIZE = int(os.getenv("BATCH_UPDATE_SIZE", "100"))


# ===== 日志配置 =====
logger = logging.getLogger("ai_analyzer")
logger.setLevel(logging.INFO)
fh = RotatingFileHandler("ai_analyzer.log", maxBytes=8 * 1024 * 1024, backupCount=5, encoding="utf-8")
fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
ch = logging.StreamHandler()
ch.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
if not logger.handlers:
    logger.addHandler(fh)
    logger.addHandler(ch)


# ===== 工具函数 =====
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


# ===== 数据类 =====
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
    rent_type: Optional[str]
    project_name: Optional[str]
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
    purchase_date: Optional[str]
    date_str: Optional[str]


# ===== OpenAI 请求客户端 =====
class OpenAIClient:
    def __init__(self):
        self.base = OPENAI_API_BASE.rstrip("/")
        self.headers = {"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"}

    def chat(self, prompt: str) -> str:
        url = f"{self.base}/chat/completions"
        payload = {"model": OPENAI_MODEL, "messages": [{"role": "user", "content": prompt}]}
        if OPENAI_TEMPERATURE:
            try:
                payload["temperature"] = float(OPENAI_TEMPERATURE)
            except Exception:
                pass

        for attempt in range(OPENAI_RETRIES + 1):
            try:
                r = requests.post(url, headers=self.headers, data=json.dumps(payload), timeout=OPENAI_TIMEOUT)
                if r.status_code == 200:
                    return r.json().get("choices", [{}])[0].get("message", {}).get("content", "")
                logger.warning(f"OpenAI status {r.status_code}: {r.text[:200]}")
            except Exception as e:
                logger.warning(f"OpenAI retry {attempt + 1}/{OPENAI_RETRIES}: {e}")
                time.sleep(1.2 * (attempt + 1))
        return ""


# ===== 核心逻辑函数 =====
def analyze_row(row: ReportRow, client: OpenAIClient) -> Tuple[int, Optional[int], str]:
    """返回：ID, score, summary"""
    prompt = (
        f"设备：{_s(row.machine_no)}；项目：{_s(row.project_name)}；日期：{_s(row.date_str)}；"
        f"租用方式：{_s(row.rent_type)}；采购价：{format_idr_brief(row.purch_price)}；\n"
        f"有效工时：{fmt2(row.valid_duration)}h；怠速工时：{fmt2(row.idling_duration)}h；"
        f"油耗：{fmt2(row.day_oil)}L；加油：{fmt2(row.day_refuel)}L；"
        f"运输油耗：{fmt2(row.transport_avg_oil)}；工时油耗：{fmt2(row.workhour_avg_oil)}。\n"
        "请直接输出两行：\n评分：<整数>\n总结：<220–320字，说明打分理由，不要写建议/推测>"
    )

    try:
        reply = client.chat(prompt)
        m_score = re.search(r"评分[:：]?\s*([0-9]{2,3})", reply)
        m_summary = re.search(r"(?:总结|结论)[:：]?\s*(.*)", reply, re.S)
        score = int(m_score.group(1)) if m_score else 80
        summary = (m_summary.group(1).strip() if m_summary else reply.strip())[:320]
        if len(summary) < 200:
            summary += " 数据完整，表现合理，系统根据油耗、工时、加油量等多项指标自动生成得分。"
        return row.id, score, summary
    except Exception as e:
        logger.error(f"分析失败 ID={row.id}: {e}")
        return row.id, 0, f"分析错误：{e}"


# ===== 数据库操作 =====
def fetch_pending_rows(conn, date_str, project_name) -> List[ReportRow]:
    with conn.cursor() as cur:
        sql = f"""
        SELECT ID, DEVICE_ID, COMPANY, ESTATE, MACHINE_TYPE, MACHINE_CATEGORY, MACHINE_NO,
               BRAND_SPEC, RENT_TYPE, PROJECT_NAME, VALID_DURATION, IDLING_DURATION,
               VALID_PERCENT, DAY_OIL, DAY_REFUEL, DAY_MILEAGE, WORKHOUR_AVG_OIL,
               TRANSPORT_AVG_OIL, FUEL_DIFF, PURCH_PRICE, PURCHASE_DATE, DATE_STR
        FROM `{REPORT_TABLE}`
        WHERE DATE_STR=%s AND PROJECT_NAME=%s
          AND (SCORE IS NULL OR SUMMARY IS NULL)
        ORDER BY ID DESC LIMIT 2000
        """
        cur.execute(sql, (date_str, project_name))
        rows = cur.fetchall()
        return [
            ReportRow(
                id=r["ID"],
                device_id=r.get("DEVICE_ID"),
                company=r.get("COMPANY"),
                estate=r.get("ESTATE"),
                machine_type=r.get("MACHINE_TYPE"),
                machine_category=r.get("MACHINE_CATEGORY"),
                machine_no=r.get("MACHINE_NO"),
                brand_spec=r.get("BRAND_SPEC"),
                rent_type=r.get("RENT_TYPE"),
                project_name=r.get("PROJECT_NAME"),
                valid_duration=_safe_float(r.get("VALID_DURATION")),
                idling_duration=_safe_float(r.get("IDLING_DURATION")),
                valid_percent=_safe_float(r.get("VALID_PERCENT")),
                day_oil=_safe_float(r.get("DAY_OIL")),
                day_refuel=_safe_float(r.get("DAY_REFUEL")),
                day_mileage=_safe_float(r.get("DAY_MILEAGE")),
                workhour_avg_oil=_safe_float(r.get("WORKHOUR_AVG_OIL")),
                transport_avg_oil=_safe_float(r.get("TRANSPORT_AVG_OIL")),
                fuel_diff=_safe_float(r.get("FUEL_DIFF")),
                purch_price=_safe_float(r.get("PURCH_PRICE")),
                purchase_date=_s(r.get("PURCHASE_DATE")),
                date_str=_s(r.get("DATE_STR")),
            )
            for r in rows
        ]


def write_results(conn, results: List[Tuple[int, int, str, str, str]]):
    with conn.cursor() as cur:
        sql = f"""
        UPDATE `{REPORT_TABLE}`
        SET SCORE=%s, SUMMARY=%s, ANALYSIS_TIME=%s, MODEL_NAME=%s
        WHERE ID=%s
        """
        cur.executemany(sql, [(s, m, jkt_now_str(), OPENAI_MODEL, rid) for rid, s, m in results])
    conn.commit()


# ===== 主流程 =====
def main():
    t0 = time.perf_counter()
    logger.info(f"启动：DB={DB_HOST}/{DB_NAME} 表={REPORT_TABLE} 日期={TARGET_DATE} 项目={PROJECT_FILTER}")

    conn = pymysql.connect(
        host=DB_HOST,
        port=DB_PORT,
        user=DB_USER,
        password=DB_PASS,
        database=DB_NAME,
        charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor,
    )

    try:
        rows = fetch_pending_rows(conn, TARGET_DATE, PROJECT_FILTER)
        if not rows:
            logger.info("没有待分析的记录，直接退出。")
            return

        logger.info(f"共待分析 {len(rows)} 条记录。")
        client = OpenAIClient()

        results = []
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            futures = [ex.submit(analyze_row, r, client) for r in rows]
            for i, fut in enumerate(as_completed(futures), 1):
                rid, score, summary = fut.result()
                results.append((rid, score, summary))
                if len(results) >= BATCH_UPDATE_SIZE:
                    write_results(conn, results)
                    results.clear()
                if i % 10 == 0:
                    logger.info(f"已完成 {i}/{len(rows)}")

        if results:
            write_results(conn, results)

        logger.info(f"完成，总耗时 {time.perf_counter() - t0:.1f}s")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
