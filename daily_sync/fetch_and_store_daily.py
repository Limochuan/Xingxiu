#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime as dt
import json
import os
import sys
import time
from typing import Any, Dict, List

import pymysql
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


def env(key: str, default: str | None = None) -> str | None:
    v = os.getenv(key)
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return default
    return v


# ===== 接口与数据库配置（可被 GitHub Actions 的 env 覆盖）=====
BASE_URL = env("API_BASE_URL", "http://119.47.88.14:81/admin/common/mechanical/ai_export")

DB_HOST = env("DB_HOST", "rm-k1a5w7qk9cnm74r25wo.mysql.ap-southeast-5.rds.aliyuncs.com")
_port_str = env("DB_PORT", "3306")
DB_PORT = int(_port_str) if (_port_str and _port_str.isdigit()) else 3306
DB_USER = env("DB_USER", "script_xingxiu")
DB_PASS = env("DB_PASS", "Julong678678678")
DB_NAME = env("DB_NAME", "xingxiu_db")
DB_TABLE = env("DB_TABLE", "xingxiu_daily")

# 手动指定跑哪天（形如 2025-10-01）；不指定则默认“雅加达昨天”
DATE_OVERRIDE = env("DATE_OVERRIDE")
# ============================================================


# 表字段顺序（与 INSERT/UPDATE 占位一一对应）
COLUMN_ORDER = [
    "DEVICE_NO", "PROJECT_NAME", "MECHANICAL_NO",
    "TYPE_NAME", "CAR_TYPE", "RENT_TYPE",
    "VALID_DURATION", "IDLING_DURATION", "DAY_OIL", "DAY_REFUEL", "DAY_MILEAGE", "VALID_PERCENT",
    "WORKHOUR_AVG_OIL", "TRANSPORT_AVG_OIL",
    "DATE_STR", "CREATE_TIME"
]

# 接口字段 -> 表字段
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


def make_session() -> requests.Session:
    """带重试的 Session，避免接口偶发超时."""
    session = requests.Session()
    retry = Retry(
        total=5, connect=3, read=5,
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
    """POST 请求接口，带上 dateStr，返回 dataList 列表."""
    # 兜底：确保 BASE_URL 带协议
    if not BASE_URL.lower().startswith(("http://", "https://")):
        print(f"[WARN] API_BASE_URL 无效：{BASE_URL!r}，改用默认。")
        base = "http://119.47.88.14:81/admin/common/mechanical/ai_export"
    else:
        base = BASE_URL

    url = f"{base}?dateStr={date_str}"
    session = make_session()
    timeout = (10, 120)  # (连接超时, 读取超时)

    for attempt in range(1, 6):
        try:
            print(f"[INFO] 第{attempt}次请求: ****?dateStr={date_str}")
            print(f"[INFO] 正在向接口发送 POST 请求（连接{timeout[0]}秒，读取{timeout[1]}秒）...")
            t0 = time.time()
            resp = session.post(url, timeout=timeout)
            elapsed = time.time() - t0
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


def _jakarta_now_str() -> str:
    """返回雅加达当前时间字符串，用于 CREATE_TIME."""
    try:
        from zoneinfo import ZoneInfo  # py3.9+
        tz = ZoneInfo("Asia/Jakarta")
        return dt.datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        # 极端情况下找不到 zoneinfo，则退回固定 +07:00
        class _FixedTZ(dt.tzinfo):
            def utcoffset(self, _): return dt.timedelta(hours=7)
            def tzname(self, _): return "UTC+07"
            def dst(self, _): return dt.timedelta(0)
        tz = _FixedTZ()
        return dt.datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")


def to_db_record(src: Dict[str, Any], date_str: str) -> Dict[str, Any]:
    """把接口字段映射成表记录，并补全 DATE_STR/CREATE_TIME."""
    dst = {col: None for col in COLUMN_ORDER}
    for k, v in src.items():
        if k in FIELD_MAP:
            val = v if v != "" else None
            # 业务约定：没有传感器数据时 WORKHOUR_AVG_OIL 记为 -1
            if FIELD_MAP[k] == "WORKHOUR_AVG_OIL" and (val is None):
                val = -1
            dst[FIELD_MAP[k]] = val

    dst["DATE_STR"] = date_str
    # ✅ 用雅加达时间
    dst["CREATE_TIME"] = _jakarta_now_str()
    return dst


def upsert(records: List[Dict[str, Any]]) -> int:
    """批量 upsert；唯一键: (DEVICE_NO, DATE_STR)."""
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
            update_list = ", ".join([f"`{c}`=VALUES(`{c}`)" for c in cols if c not in ("DEVICE_NO", "DATE_STR")])

            sql = (
                f"INSERT INTO `{DB_TABLE}` ({col_list}) VALUES ({placeholders}) "
                f"ON DUPLICATE KEY UPDATE {update_list}"
            )
            data = [[rec.get(c) for c in cols] for rec in records]
            affected = cur.executemany(sql, data)
        conn.commit()
        print(f"[INFO] 数据库写入完成，受影响行数={affected}")
        return affected
    finally:
        conn.close()


def jakarta_yesterday() -> str:
    """返回雅加达时区的“昨天” YYYY-MM-DD."""
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo("Asia/Jakarta")
    except Exception:
        class _FixedTZ(dt.tzinfo):
            def utcoffset(self, _): return dt.timedelta(hours=7)
            def tzname(self, _): return "UTC+07"
            def dst(self, _): return dt.timedelta(0)
        tz = _FixedTZ()

    now_utc = dt.datetime.utcnow().replace(tzinfo=dt.timezone.utc)
    return (now_utc.astimezone(tz).date() - dt.timedelta(days=1)).strftime("%Y-%m-%d")


def main():
    t0 = time.time()
    date_str = DATE_OVERRIDE or jakarta_yesterday()
    print(f"[INFO] 开始任务，目标日期: {date_str}")

    rows = fetch_data(date_str)
    if not rows:
        print("[WARN] 接口无数据或请求失败。")
        return

    db_rows = [to_db_record(r, date_str) for r in rows if r.get("deviceNo")]
    affected = upsert(db_rows)

    print(f"[INFO] 任务结束。总数={len(rows)}，入库={len(db_rows)}，受影响={affected}，耗时={time.time()-t0:.1f}s")


if __name__ == "__main__":
    main()
