#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
本地一条龙脚本：
1. 如果指定日期已经写入过 xingxiu_daily → 直接退出
2. 否则从接口拉取指定日期数据，写入 xingxiu_daily
3. 再将指定日期的数据与 xingxiu_device_info 合并，写入 xingxiu_daily_report
"""

import os
import datetime as dt
import sys
import time
from typing import Any, Dict, List

import pymysql
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ===== 配置 =====
BASE_URL = "http://119.47.88.14:81/admin/common/mechanical/ai_export"

DB_HOST = "rm-k1a5w7qk9cnm74r25wo.mysql.ap-southeast-5.rds.aliyuncs.com"
DB_PORT = 3306
DB_USER = "script_xingxiu"
DB_PASS = "Julong678678678"
DB_NAME = "xingxiu_db"

DAILY_TABLE  = "xingxiu_daily"
DEVICE_TABLE = "xingxiu_device_info"
REPORT_TABLE = "xingxiu_daily_report"

# ===== 日期逻辑 =====
env_target_date = os.getenv("TARGET_DATE", "").strip()
if env_target_date:
    FIXED_DATE = env_target_date
    print(f"[INFO] 使用环境变量 TARGET_DATE 作为日期: {FIXED_DATE}")
else:
    FIXED_DATE = (dt.date.today() - dt.timedelta(days=1)).strftime("%Y-%m-%d")
    print(f"[INFO] 未指定 TARGET_DATE，使用运行时间前一天: {FIXED_DATE}")

# ===== 字段映射 =====
COLUMN_ORDER = [
    "DEVICE_NO","PROJECT_NAME","MECHANICAL_NO",
    "TYPE_NAME","CAR_TYPE","RENT_TYPE",
    "VALID_DURATION","IDLING_DURATION","DAY_OIL","DAY_REFUEL",
    "DAY_MILEAGE","VALID_PERCENT",
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

# ===== HTTP =====
def make_session() -> requests.Session:
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
    timeout = (10, 120)

    for attempt in range(1, 6):
        try:
            print(f"[INFO] 第{attempt}次请求接口: {url}")
            resp = session.post(url, timeout=timeout)
            resp.raise_for_status()

            obj = resp.json()
            data = obj.get("dataList", [])
            print(f"[INFO] 成功获取 {len(data)} 条记录")
            return data

        except Exception as e:
            print(f"[WARN] 请求异常：{e}，准备重试...", file=sys.stderr)
            time.sleep(1.5 ** attempt)

    print("[ERROR] 多次重试后仍失败。", file=sys.stderr)
    return []

# ===== 判断当天是否已跑过 =====
def daily_exists(conn, date_str: str) -> bool:
    with conn.cursor() as cur:
        cur.execute(
            f"SELECT 1 FROM `{DAILY_TABLE}` WHERE DATE_STR=%s LIMIT 1",
            (date_str,)
        )
        return cur.fetchone() is not None

# ===== daily 表写入（保持你原来的 ON DUPLICATE KEY 逻辑）=====
def upsert_daily(conn, records: List[Dict[str, Any]]) -> int:
    if not records:
        return 0

    print(f"[INFO] 正在写入 {DAILY_TABLE}，共 {len(records)} 条数据...")
    with conn.cursor() as cur:
        cols = COLUMN_ORDER
        col_list = ", ".join(f"`{c}`" for c in cols)
        placeholders = ", ".join(["%s"] * len(cols))
        update_list = ", ".join(
            [f"`{c}`=VALUES(`{c}`)" for c in cols if c not in ("DEVICE_NO","DATE_STR")]
        )
        sql = (
            f"INSERT INTO `{DAILY_TABLE}` ({col_list}) VALUES ({placeholders}) "
            f"ON DUPLICATE KEY UPDATE {update_list}"
        )
        data = [[rec.get(c) for c in cols] for rec in records]
        affected = cur.executemany(sql, data)

    print(f"[INFO] {DAILY_TABLE} 写入完成，受影响行数={affected}")
    return affected

# ===== merge 到 report =====
def first_existing_col(cur, table: str, candidates: list[str]) -> str:
    cur.execute(f"SHOW COLUMNS FROM `{table}`")
    cols = {row[0].upper() for row in cur.fetchall()}
    for c in candidates:
        if c.upper() in cols:
            return c
    raise RuntimeError(f"表 `{table}` 缺少列：" + " / ".join(candidates))

def merge_daily_to_report(conn, target_date: str) -> int:
    with conn.cursor() as cur:
        daily_device_col = first_existing_col(cur, DAILY_TABLE, ["DEVICE_NO", "DEVICE_ID"])
        daily_idle_col   = first_existing_col(cur, DAILY_TABLE, ["IDLING_DURATION", "IDING_DURATION"])
        report_idle_col  = first_existing_col(cur, REPORT_TABLE, ["IDLING_DURATION", "IDING_DURATION"])

        cur.execute(f"SHOW INDEX FROM `{REPORT_TABLE}` WHERE Key_name='uniq_device_date'")
        if not cur.fetchone():
            cur.execute(
                f"ALTER TABLE `{REPORT_TABLE}` "
                f"ADD UNIQUE KEY `uniq_device_date` (`DEVICE_ID`,`DATE_STR`)"
            )

        sql = f"""
            INSERT INTO `{REPORT_TABLE}` (
                DEVICE_ID, PROJECT_NAME, MECHANICAL_NO, CAR_TYPE, DATE_STR, RENT_TYPE,
                VALID_DURATION, {report_idle_col}, VALID_PERCENT,
                DAY_OIL, DAY_REFUEL, DAY_MILEAGE,
                WORKHOUR_AVG_OIL, TRANSPORT_AVG_OIL,
                INSERT_TIME
            )
            SELECT
                d.{daily_device_col}, d.PROJECT_NAME, d.MECHANICAL_NO,
                d.CAR_TYPE, d.DATE_STR, d.RENT_TYPE,
                d.VALID_DURATION, d.{daily_idle_col}, d.VALID_PERCENT,
                d.DAY_OIL, d.DAY_REFUEL, d.DAY_MILEAGE,
                d.WORKHOUR_AVG_OIL, d.TRANSPORT_AVG_OIL,
                NOW()
            FROM `{DAILY_TABLE}` d
            WHERE d.DATE_STR=%s
            ON DUPLICATE KEY UPDATE
                VALID_DURATION=VALUES(VALID_DURATION),
                {report_idle_col}=VALUES({report_idle_col}),
                DAY_OIL=VALUES(DAY_OIL),
                DAY_REFUEL=VALUES(DAY_REFUEL),
                DAY_MILEAGE=VALUES(DAY_MILEAGE),
                VALID_PERCENT=VALUES(VALID_PERCENT),
                INSERT_TIME=NOW()
        """

        cur.execute(sql, (target_date,))
        print(f"[INFO] {REPORT_TABLE} 合并完成（DATE_STR={target_date}）")
        return cur.rowcount

# ===== 主流程 =====
def main():
    print(f"[INFO] 开始任务，目标日期: {FIXED_DATE}")

    conn = pymysql.connect(
        host=DB_HOST,
        port=DB_PORT,
        user=DB_USER,
        password=DB_PASS,
        database=DB_NAME,
        charset="utf8mb4",
        autocommit=False
    )

    try:
        # 🔴 关键判断：如果今天已经跑过，直接退出
        if daily_exists(conn, FIXED_DATE):
            print(f"[INFO] 日期 {FIXED_DATE} 已存在数据，直接退出，不再执行。")
            return

        rows = fetch_data(FIXED_DATE)
        if not rows:
            print("[WARN] 接口无数据，退出。")
            return

        db_rows = [to_db_record(r, FIXED_DATE) for r in rows if r.get("deviceNo")]

        upsert_daily(conn, db_rows)
        merge_daily_to_report(conn, FIXED_DATE)

        conn.commit()
        print("[INFO] 全流程成功完成。")

    except Exception as e:
        conn.rollback()
        print(f"[ERROR] 异常回滚：{e}", file=sys.stderr)
        raise
    finally:
        conn.close()

if __name__ == "__main__":
    main()
