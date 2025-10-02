#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime as dt
import os
import sys
import time
from typing import Any, Dict, List

import pymysql
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ===== 配置（全部读环境变量，便于GitHub Secrets管理）=====
BASE_URL = os.getenv("API_BASE_URL", "http://119.47.88.14:81/admin/common/mechanical/ai_export")

DB_HOST = os.getenv("DB_HOST", "rm-k1a5w7qk9cnm74r25wo.mysql.ap-southeast-5.rds.aliyuncs.com")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_USER = os.getenv("DB_USER", "script_xingxiu")
DB_PASS = os.getenv("DB_PASS", "Julong678678678")
DB_NAME = os.getenv("DB_NAME", "xingxiu_db")
DB_TABLE = os.getenv("DB_TABLE", "xingxiu_daily")
# ==========================================================

COLUMN_ORDER = [
    "DEVICE_NO","PROJECT_NAME","MECHANICAL_NO",
    "TYPE_NAME","CAR_TYPE","RENT_TYPE",
    "VALID_DURATION","IDLING_DURATION","DAY_OIL","DAY_REFUEL","DAY_MILEAGE","VALID_PERCENT",
    "WORKHOUR_AVG_OIL","TRANSPORT_AVG_OIL",
    "DATE_STR","CREATE_TIME"
]

FIELD_MAP = {
    "deviceNo": "DEVICE_NO",
    "projectName": "PROJECT_NAME",
    "mechanicalNo": "MECHANICAL_NO",
    "typeName": "TYPE_NAME",
    "carType": "CAR_TYPE",
    "rentType": "RENT_TYPE",
    "validDuration": "VALID_DURATION",
    "idlingDuration": "IDLING_DURATION",
    "dayOil": "DAY_OIL",
    "dayRefuel": "DAY_REFUEL",
    "dayMileage": "DAY_MILEAGE",
    "validPercent": "VALID_PERCENT",
    "workhourAvgOil": "WORKHOUR_AVG_OIL",
    "transportAvgOil": "TRANSPORT_AVG_OIL",
}

def jakarta_yesterday(today_utc: dt.datetime | None = None) -> str:
    """按 Asia/Jakarta 计算昨天 (YYYY-MM-DD)"""
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo("Asia/Jakarta")
    except Exception:
        class _FixedTZ(dt.tzinfo):
            def utcoffset(self, _): return dt.timedelta(hours=7)
            def tzname(self, _): return "UTC+07"
            def dst(self, _): return dt.timedelta(0)
        tz = _FixedTZ()
    now_utc = today_utc or dt.datetime.utcnow().replace(tzinfo=dt.timezone.utc)
    return (now_utc.astimezone(tz).date() - dt.timedelta(days=1)).strftime("%Y-%m-%d")

def make_session() -> requests.Session:
    """带重试的 Session：对超时/连接错误等进行重试"""
    session = requests.Session()
    retry = Retry(
        total=5,
        connect=3,
        read=5,
        backoff_factor=1.5,
        status_forcelist=(502, 503, 504),
        allowed_methods=frozenset(["POST"])
    )
    adapter = HTTPAdapter(max_retries=retry, pool_maxsize=10)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update({"Connection": "close"})
    return session

def fetch_data(date_str: str) -> List[Dict[str, Any]]:
    url = f"{BASE_URL}?dateStr={date_str}"
    session = make_session()
    timeout = (10, 120)  # (连接超时, 读取超时)

    for attempt in range(1, 6):
        try:
            print(f"[INFO] 第{attempt}次请求: {url}")
            print(f"[INFO] 正在向接口发送 POST 请求，超时设置=连接{timeout[0]}秒, 读取{timeout[1]}秒...")

            start_time = dt.datetime.now()
            resp = session.post(url, timeout=timeout)
            elapsed = (dt.datetime.now() - start_time).total_seconds()

            print(f"[INFO] 接口已响应，用时 {elapsed:.1f} 秒")
            print(f"[INFO] HTTP {resp.status_code}, Content-Type={resp.headers.get('Content-Type')}")

            resp.raise_for_status()
            obj = resp.json()
            if not isinstance(obj, dict) or "dataList" not in obj or not isinstance(obj["dataList"], list):
                print("[WARN] 返回结构异常，前 300 字符：", str(obj)[:300], file=sys.stderr)
                return []
            print(f"[INFO] 成功获取到 {len(obj['dataList'])} 条记录")
            return obj["dataList"]

        except Exception as e:
            print(f"[WARN] 请求异常：{e}，准备重试...", file=sys.stderr)
        time.sleep(1.5 ** attempt)

    print("[ERROR] 多次重试后仍失败。", file=sys.stderr)
    return []

def to_db_record(src: Dict[str, Any], date_str: str) -> Dict[str, Any]:
    dst = {col: None for col in COLUMN_ORDER}
    for k, v in src.items():
        if k in FIELD_MAP:
            val = v if v != "" else None
            # 特殊：传感器缺失 → -1
            if FIELD_MAP[k] == "WORKHOUR_AVG_OIL" and (val is None):
                val = -1
            dst[FIELD_MAP[k]] = val
    dst["DATE_STR"] = date_str
    dst["CREATE_TIME"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return dst

def upsert(records: List[Dict[str, Any]]) -> int:
    """直接 UPSERT（如遇唯一键冲突即更新）。"""
    if not records:
        return 0
    print(f"[INFO] 正在写入数据库，共 {len(records)} 条数据...")
    conn = pymysql.connect(
        host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASS,
        database=DB_NAME, charset="utf8mb4", autocommit=False
    )
    try:
        with conn.cursor() as cur:
            cols = COLUMN_ORDER
            col_list = ", ".join(f"`{c}`" for c in cols)
            placeholders = ", ".join(["%s"] * len(cols))
            update_list = ", ".join([f"`{c}`=VALUES(`{c}`)" for c in cols if c not in ("DEVICE_NO","DATE_STR")])
            sql = f"INSERT INTO `{DB_TABLE}` ({col_list}) VALUES ({placeholders}) ON DUPLICATE KEY UPDATE {update_list}"
            data = [[rec.get(c) for c in cols] for rec in records]
            affected = cur.executemany(sql, data)
        conn.commit()
        print(f"[INFO] 数据库写入完成，受影响行数={affected}")
        return affected
    finally:
        conn.close()

def main():
    # 允许手动覆盖日期（例如 Actions 里传入），否则按 JKT 计算昨天
    manual_date = os.getenv("DATE_OVERRIDE", "").strip()
    date_str = manual_date if manual_date else jakarta_yesterday()
    print(f"[INFO] 开始任务，目标日期: {date_str}")

    rows = fetch_data(date_str)
    if not rows:
        print("[WARN] 没有数据需要写入。")
        return

    db_rows = [to_db_record(r, date_str) for r in rows if r.get("deviceNo")]
    affected = upsert(db_rows)
    print("[INFO] 任务结束。")

if __name__ == "__main__":
    main()
