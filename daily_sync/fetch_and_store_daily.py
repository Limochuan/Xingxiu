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

# ---------- 工具函数 ----------
def require_env(name: str) -> str:
    """必须的环境变量；缺失就退出（不回显敏感值）"""
    v = os.getenv(name)
    if v is None or str(v).strip() == "":
        print(f"[ERR] Missing required environment variable: {name}", file=sys.stderr)
        sys.exit(2)
    return v.strip()

def int_env(name: str, default: int) -> int:
    v = os.getenv(name)
    return int(v) if v not in (None, "") else default

# ---------- 配置（全部来自环境变量） ----------
# 接口地址（建议放到 GitHub Secrets / Repository variables 中）
API_BASE = require_env("API_BASE")  # 例：http://.../ai_export

DB_HOST  = require_env("DB_HOST")
DB_PORT  = int_env("DB_PORT", 3306)   # 端口可用安全默认
DB_USER  = require_env("DB_USER")
DB_PASS  = require_env("DB_PASS")
DB_NAME  = require_env("DB_NAME")
DB_TABLE = os.getenv("DB_TABLE", "xingxiu_daily")  # 表名非敏感，可给默认

# 列顺序/字段映射（非敏感）
COLUMN_ORDER = [
    "DEVICE_NO", "PROJECT_NAME", "MECHANICAL_NO",
    "TYPE_NAME", "CAR_TYPE", "RENT_TYPE",
    "VALID_DURATION", "IDLING_DURATION", "DAY_OIL", "DAY_REFUEL", "DAY_MILEAGE", "VALID_PERCENT",
    "WORKHOUR_AVG_OIL", "TRANSPORT_AVG_OIL",
    "DATE_STR", "CREATE_TIME"
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

# ---------- 时区/日期 ----------
def jakarta_yesterday(today_utc: dt.datetime | None = None) -> str:
    """按 Asia/Jakarta 计算昨天 (YYYY-MM-DD)"""
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo("Asia/Jakarta")
        now_utc = today_utc or dt.datetime.utcnow().replace(tzinfo=dt.timezone.utc)
        return (now_utc.astimezone(tz).date() - dt.timedelta(days=1)).strftime("%Y-%m-%d")
    except Exception:
        # 兜底：固定 UTC+7
        now_utc = today_utc or dt.datetime.utcnow()
        return (now_utc + dt.timedelta(hours=7) - dt.timedelta(days=1)).date().strftime("%Y-%m-%d")

def get_target_date() -> str:
    # 优先 DATE_STR；兼容旧变量 DATE_OVERRIDE
    date_str = (os.getenv("DATE_STR") or os.getenv("DATE_OVERRIDE") or "").strip()
    if not date_str:
        date_str = jakarta_yesterday()
        print(f"[INFO] 未提供 DATE_STR，按雅加达时区默认取前一天：{date_str}")
    # 校验
    try:
        dt.datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        print(f"[ERR] 非法日期：{date_str}，期望 YYYY-MM-DD", file=sys.stderr)
        sys.exit(2)
    return date_str

# ---------- HTTP ----------
def make_session() -> requests.Session:
    """带重试的 Session"""
    session = requests.Session()
    retry = Retry(
        total=5, connect=3, read=5,
        backoff_factor=1.5,
        status_forcelist=(502, 503, 504),
        allowed_methods=frozenset(["POST"]),
    )
    adapter = HTTPAdapter(max_retries=retry, pool_maxsize=10)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update({"Connection": "close"})
    return session

def fetch_data(date_str: str) -> List[Dict[str, Any]]:
    url = f"{API_BASE}?dateStr={date_str}"
    session = make_session()
    timeout = (10, 120)  # (连接超时, 读取超时)

    for attempt in range(1, 6):
        try:
            print(f"[INFO] 第{attempt}次请求: {url}")
            start_time = dt.datetime.now()
            resp = session.post(url, timeout=timeout)
            elapsed = (dt.datetime.now() - start_time).total_seconds()

            print(f"[INFO] 接口已响应，用时 {elapsed:.1f} 秒，HTTP {resp.status_code}")
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

# ---------- 转换/写库 ----------
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

    # 写入时间（雅加达）
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo("Asia/Jakarta")
        dst["CREATE_TIME"] = dt.datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        dst["CREATE_TIME"] = (dt.datetime.utcnow() + dt.timedelta(hours=7)).strftime("%Y-%m-%d %H:%M:%S")

    return dst

def upsert(records: List[Dict[str, Any]]) -> int:
    """直接 UPSERT（唯一键冲突即更新）。"""
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
            # 唯一键一般是 (DEVICE_NO, DATE_STR)；不要更新它们
            update_cols = [c for c in cols if c not in ("DEVICE_NO", "DATE_STR")]
            update_list = ", ".join(f"`{c}`=VALUES(`{c}`)" for c in update_cols)
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

# ---------- 主流程 ----------
def main():
    date_str = get_target_date()
    print(f"[INFO] 开始任务，目标日期: {date_str}")

    rows = fetch_data(date_str)
    if not rows:
        print("[WARN] 接口无数据，结束。")
        return

    db_rows = [to_db_record(r, date_str) for r in rows if r.get("deviceNo")]
    affected = upsert(db_rows)
    print(f"[INFO] 任务结束，写入/更新 {affected} 行。")

if __name__ == "__main__":
    main()
