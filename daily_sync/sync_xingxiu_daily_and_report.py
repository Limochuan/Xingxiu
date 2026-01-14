#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
本地一条龙脚本（幂等版，零浪费 AUTO_INCREMENT）：
1. 从接口拉取指定日期的数据，写入 xingxiu_daily
   - 先 UPDATE（不占 ID）
   - UPDATE 不命中才 INSERT（才占 ID）
2. 再将指定日期的数据与 xingxiu_device_info 合并，写入 xingxiu_daily_report
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
    print(f"[INFO] 使用环境变量 TARGET_DATE: {FIXED_DATE}")
else:
    FIXED_DATE = (dt.date.today() - dt.timedelta(days=1)).strftime("%Y-%m-%d")
    print(f"[INFO] 使用默认日期（昨天）: {FIXED_DATE}")

# ===== 字段映射 =====
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
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.headers.update({"Connection": "close"})
    return session

def fetch_data(date_str: str) -> List[Dict[str, Any]]:
    url = f"{BASE_URL}?dateStr={date_str}"
    session = make_session()

    for i in range(1, 6):
        try:
            print(f"[INFO] 请求接口（第{i}次）: {url}")
            resp = session.post(url, timeout=(10, 120))
            resp.raise_for_status()
            data = resp.json().get("dataList", [])
            print(f"[INFO] 获取到 {len(data)} 条记录")
            return data
        except Exception as e:
            print(f"[WARN] 请求失败: {e}")
            time.sleep(1.5 ** i)

    print("[ERROR] 接口多次失败，终止")
    return []

# ===== DAILY：先 UPDATE 再 INSERT =====
def upsert_daily_safe(conn, rows: List[Dict[str, Any]], date_str: str):
    if not rows:
        return

    update_sql = f"""
        UPDATE {DAILY_TABLE}
        SET
            PROJECT_NAME=%s,
            MECHANICAL_NO=%s,
            TYPE_NAME=%s,
            CAR_TYPE=%s,
            RENT_TYPE=%s,
            VALID_DURATION=%s,
            IDLING_DURATION=%s,
            DAY_OIL=%s,
            DAY_REFUEL=%s,
            DAY_MILEAGE=%s,
            VALID_PERCENT=%s,
            WORKHOUR_AVG_OIL=%s,
            TRANSPORT_AVG_OIL=%s,
            CREATE_TIME=%s
        WHERE DEVICE_NO=%s AND DATE_STR=%s
    """

    insert_sql = f"""
        INSERT INTO {DAILY_TABLE} (
            DEVICE_NO, PROJECT_NAME, MECHANICAL_NO,
            TYPE_NAME, CAR_TYPE, RENT_TYPE,
            VALID_DURATION, IDLING_DURATION,
            DAY_OIL, DAY_REFUEL, DAY_MILEAGE,
            VALID_PERCENT, WORKHOUR_AVG_OIL,
            TRANSPORT_AVG_OIL, DATE_STR, CREATE_TIME
        ) VALUES (
            %s,%s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s
        )
    """

    updated = inserted = 0
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with conn.cursor() as cur:
        for r in rows:
            if not r.get("deviceNo"):
                continue

            vals = {FIELD_MAP[k]: r.get(k) for k in FIELD_MAP if k in r}
            cur.execute(
                update_sql,
                (
                    vals.get("PROJECT_NAME"),
                    vals.get("MECHANICAL_NO"),
                    vals.get("TYPE_NAME"),
                    vals.get("CAR_TYPE"),
                    vals.get("RENT_TYPE"),
                    vals.get("VALID_DURATION"),
                    vals.get("IDLING_DURATION"),
                    vals.get("DAY_OIL"),
                    vals.get("DAY_REFUEL"),
                    vals.get("DAY_MILEAGE"),
                    vals.get("VALID_PERCENT"),
                    vals.get("WORKHOUR_AVG_OIL") or -1,
                    vals.get("TRANSPORT_AVG_OIL"),
                    now,
                    vals.get("DEVICE_NO"),
                    date_str,
                ),
            )

            if cur.rowcount == 0:
                cur.execute(
                    insert_sql,
                    (
                        vals.get("DEVICE_NO"),
                        vals.get("PROJECT_NAME"),
                        vals.get("MECHANICAL_NO"),
                        vals.get("TYPE_NAME"),
                        vals.get("CAR_TYPE"),
                        vals.get("RENT_TYPE"),
                        vals.get("VALID_DURATION"),
                        vals.get("IDLING_DURATION"),
                        vals.get("DAY_OIL"),
                        vals.get("DAY_REFUEL"),
                        vals.get("DAY_MILEAGE"),
                        vals.get("VALID_PERCENT"),
                        vals.get("WORKHOUR_AVG_OIL") or -1,
                        vals.get("TRANSPORT_AVG_OIL"),
                        date_str,
                        now,
                    ),
                )
                inserted += 1
            else:
                updated += 1

    print(f"[INFO] DAILY 完成：UPDATE={updated}, INSERT={inserted}")

# ===== REPORT 合并（保持你原逻辑）=====
def merge_daily_to_report(conn, date_str: str):
    with conn.cursor() as cur:
        cur.execute(f"""
            INSERT INTO {REPORT_TABLE}
            SELECT
                NULL,
                d.DEVICE_NO,
                d.PROJECT_NAME,
                d.MECHANICAL_NO,
                d.CAR_TYPE,
                d.DATE_STR,
                d.RENT_TYPE,
                d.VALID_DURATION,
                d.IDLING_DURATION,
                d.VALID_PERCENT,
                d.DAY_OIL,
                d.DAY_REFUEL,
                d.DAY_MILEAGE,
                d.WORKHOUR_AVG_OIL,
                d.TRANSPORT_AVG_OIL,
                NOW()
            FROM {DAILY_TABLE} d
            WHERE d.DATE_STR = %s
            ON DUPLICATE KEY UPDATE
                VALID_DURATION=VALUES(VALID_DURATION),
                IDLING_DURATION=VALUES(IDLING_DURATION),
                DAY_OIL=VALUES(DAY_OIL),
                DAY_REFUEL=VALUES(DAY_REFUEL),
                DAY_MILEAGE=VALUES(DAY_MILEAGE),
                VALID_PERCENT=VALUES(VALID_PERCENT),
                INSERT_TIME=NOW()
        """, (date_str,))
        print(f"[INFO] REPORT 合并完成（DATE_STR={date_str}）")

# ===== 主流程 =====
def main():
    print(f"[INFO] 开始执行，日期={FIXED_DATE}")
    rows = fetch_data(FIXED_DATE)
    if not rows:
        print("[WARN] 无数据，结束")
        return

    conn = pymysql.connect(
        host=DB_HOST,
        port=DB_PORT,
        user=DB_USER,
        password=DB_PASS,
        database=DB_NAME,
        charset="utf8mb4",
        autocommit=False,
    )

    try:
        upsert_daily_safe(conn, rows, FIXED_DATE)
        merge_daily_to_report(conn, FIXED_DATE)
        conn.commit()
        print("[INFO] 全流程成功完成")
    except Exception as e:
        conn.rollback()
        print(f"[ERROR] 异常回滚: {e}", file=sys.stderr)
        raise
    finally:
        conn.close()

if __name__ == "__main__":
    main()
