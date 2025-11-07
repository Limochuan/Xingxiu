#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
本地/Actions 一条龙脚本：
1. 从接口拉取目标日期 (TARGET_DATE) 的数据，写入 xingxiu_daily
2. 再将目标日期的数据与 xingxiu_device_info 合并，写入 xingxiu_daily_report

日期规则：
- 如果环境变量 TARGET_DATE 有值（例如手动 dispatch 传入）：直接用它
- 否则：取雅加达当前日期的前一天作为目标日期
"""

import os
import sys
import time
import datetime as dt
from typing import Any, Dict, List

import pymysql
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except ImportError:
    ZoneInfo = None  # GitHub Actions 用 3.11，这里只是兜底


# ===== 配置：从环境变量读取（配合 GitHub Actions secrets） =====
BASE_URL = os.getenv(
    "XINGXIU_BASE_URL",
    "http://119.47.88.14:81/admin/common/mechanical/ai_export",
)

DB_HOST = os.getenv("XINGXIU_DB_HOST", "localhost")
DB_PORT = int(os.getenv("XINGXIU_DB_PORT") or 3306)
DB_USER = os.getenv("XINGXIU_DB_USER", "root")
DB_PASS = os.getenv("XINGXIU_DB_PASS", "")
DB_NAME = os.getenv("XINGXIU_DB_NAME", "xingxiu_db")

DAILY_TABLE = "xingxiu_daily"
DEVICE_TABLE = "xingxiu_device_info"
REPORT_TABLE = "xingxiu_daily_report"

# ===== 公共字段映射配置（给 daily 表用） =====
COLUMN_ORDER = [
    "DEVICE_NO",
    "PROJECT_NAME",
    "MECHANICAL_NO",
    "TYPE_NAME",
    "CAR_TYPE",
    "RENT_TYPE",
    "VALID_DURATION",
    "IDLING_DURATION",
    "DAY_OIL",
    "DAY_REFUEL",
    "DAY_MILEAGE",
    "VALID_PERCENT",
    "WORKHOUR_AVG_OIL",
    "TRANSPORT_AVG_OIL",
    "DATE_STR",
    "CREATE_TIME",
]

FIELD_MAP: Dict[str, str] = {
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


# ===== 日期逻辑：优先 TARGET_DATE，否则用雅加达“昨天” =====
def get_target_date_str() -> str:
    """
    1) 如果环境变量 TARGET_DATE 有值，直接用（方便手动 backfill）
    2) 否则：取 Asia/Jakarta 当前日期的前一天
    """
    env_date = os.getenv("TARGET_DATE", "").strip()
    if env_date:
        print(f"[INFO] 使用手动指定的 TARGET_DATE={env_date}")
        return env_date

    if ZoneInfo is not None:
        now_jkt = dt.datetime.now(ZoneInfo("Asia/Jakarta"))
    else:
        # 简单兜底：按本地时间减去 7 小时近似 JKT，再取日期
        now_local = dt.datetime.now()
        now_jkt = now_local - dt.timedelta(hours=7)

    yesterday = now_jkt.date() - dt.timedelta(days=1)
    date_str = yesterday.strftime("%Y-%m-%d")
    print(f"[INFO] 未指定 TARGET_DATE，自动使用雅加达昨天: {date_str}")
    return date_str


# ===== HTTP 部分：从接口拉数据 =====
def make_session() -> requests.Session:
    """带重试的 Session：对超时/连接错误等进行重试"""
    session = requests.Session()
    retry = Retry(
        total=5,
        connect=3,
        read=5,
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
    url = f"{BASE_URL}?dateStr={date_str}"
    session = make_session()
    timeout = (10, 120)  # (连接超时, 读取超时)

    for attempt in range(1, 6):
        try:
            print(f"[INFO] 第{attempt}次请求: {url}")
            print(
                f"[INFO] 正在向接口发送 POST 请求，超时设置=连接{timeout[0]}秒, 读取{timeout[1]}秒..."
            )

            start_time = dt.datetime.now()
            resp = session.post(url, timeout=timeout)
            elapsed = (dt.datetime.now() - start_time).total_seconds()

            print(f"[INFO] 接口已响应，用时 {elapsed:.1f} 秒")
            print(
                f"[INFO] HTTP {resp.status_code}, Content-Type={resp.headers.get('Content-Type')}"
            )

            resp.raise_for_status()
            obj = resp.json()
            if (
                not isinstance(obj, dict)
                or "dataList" not in obj
                or not isinstance(obj["dataList"], list)
            ):
                print(
                    "[WARN] 返回结构异常，前 300 字符：",
                    str(obj)[:300],
                    file=sys.stderr,
                )
                return []

            print(f"[INFO] 成功获取到 {len(obj['dataList'])} 条记录")
            return obj["dataList"]

        except Exception as e:
            print(f"[WARN] 请求异常：{e}，准备重试...", file=sys.stderr)
            time.sleep(1.5**attempt)

    print("[ERROR] 多次重试后仍失败。", file=sys.stderr)
    return []


# ===== daily 表：转换 + 写库 =====
def to_db_record(src: Dict[str, Any], date_str: str) -> Dict[str, Any]:
    dst: Dict[str, Any] = {col: None for col in COLUMN_ORDER}
    for k, v in src.items():
        if k in FIELD_MAP:
            val = v if v != "" else None
            # workhourAvgOil 为空时，给默认 -1（你原来的逻辑）
            if FIELD_MAP[k] == "WORKHOUR_AVG_OIL" and val is None:
                val = -1
            dst[FIELD_MAP[k]] = val
    dst["DATE_STR"] = date_str
    dst["CREATE_TIME"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return dst


def upsert_daily(conn, records: List[Dict[str, Any]]) -> int:
    """写入/更新到 xingxiu_daily 表"""
    if not records:
        print("[INFO] 没有需要写入 daily 表的记录。")
        return 0

    print(f"[INFO] 正在写入 {DAILY_TABLE}，共 {len(records)} 条数据...")
    with conn.cursor() as cur:
        cols = COLUMN_ORDER
        col_list = ", ".join(f"`{c}`" for c in cols)
        placeholders = ", ".join(["%s"] * len(cols))
        update_list = ", ".join(
            [f"`{c}`=VALUES(`{c}`)" for c in cols if c not in ("DEVICE_NO", "DATE_STR")]
        )
        sql = (
            f"INSERT INTO `{DAILY_TABLE}` ({col_list}) VALUES ({placeholders}) "
            f"ON DUPLICATE KEY UPDATE {update_list}"
        )
        data = [[rec.get(c) for c in cols] for rec in records]
        affected = cur.executemany(sql, data)
    print(f"[INFO] {DAILY_TABLE} 写入完成，受影响行数={affected}")
    return affected


# ===== merge 部分：daily + device → report =====
def first_existing_col(cur, table: str, candidates: List[str]) -> str:
    """返回 table 中第一个存在的列名（大小写不敏感）；若都不存在则抛错。"""
    cur.execute(f"SHOW COLUMNS FROM `{table}`")
    cols = {row[0].upper() for row in cur.fetchall()}
    for c in candidates:
        if c.upper() in cols:
            return c
    raise RuntimeError(f"表 `{table}` 缺少列：" + " / ".join(candidates))


def merge_daily_to_report(conn, target_date: str) -> int:
    """
    将指定日期的日表数据与设备表合并写入业务表：
    - 目标表：xingxiu_daily_report
    - 唯一键：(DEVICE_ID, DATE_STR)，存在则更新
    """
    with conn.cursor() as cur:
        # 探测关键列名
        daily_device_col = first_existing_col(
            cur, DAILY_TABLE, ["DEVICE_NO", "DEVICE_ID"]
        )
        daily_idle_col = first_existing_col(
            cur, DAILY_TABLE, ["IDLING_DURATION", "IDING_DURATION"]
        )
        report_idle_col = first_existing_col(
            cur, REPORT_TABLE, ["IDLING_DURATION", "IDING_DURATION"]
        )

        print(f"[INFO] 日表设备列：{DAILY_TABLE}.{daily_device_col}")
        print(
            f"[INFO] 怠速列：{DAILY_TABLE}.{daily_idle_col} -> {REPORT_TABLE}.{report_idle_col}"
        )

        # 确保唯一索引 (DEVICE_ID, DATE_STR)
        cur.execute(
            f"SHOW INDEX FROM `{REPORT_TABLE}` WHERE Key_name='uniq_device_date'"
        )
        if not cur.fetchone():
            cur.execute(
                f"ALTER TABLE `{REPORT_TABLE}` "
                f"ADD UNIQUE KEY `uniq_device_date` (`DEVICE_ID`,`DATE_STR`)"
            )
            print("[INFO] 已添加唯一索引 uniq_device_date")

        upsert_sql = f"""
            INSERT INTO `{REPORT_TABLE}` (
                -- 日表字段
                DEVICE_ID, PROJECT_NAME, MECHANICAL_NO, CAR_TYPE, DATE_STR, RENT_TYPE,
                VALID_DURATION, {report_idle_col}, VALID_PERCENT, DAY_OIL, DAY_REFUEL,
                DAY_MILEAGE, WORKHOUR_AVG_OIL, TRANSPORT_AVG_OIL,

                -- 设备表字段
                COMPANY, ESTATE, MACHINE_TYPE, MACHINE_CATEGORY, MACHINE_NO,
                BRAND_SPEC, ORDER_NO, PURCHASE_DATE, DRIVER_COUNT, CREATE_TIME,
                PURCH_PRICE, FUEL_DIFF,

                -- 系统字段
                INSERT_TIME
            )
            SELECT
                d.{daily_device_col}, d.PROJECT_NAME, d.MECHANICAL_NO, d.CAR_TYPE, d.DATE_STR, d.RENT_TYPE,
                d.VALID_DURATION, d.{daily_idle_col}, d.VALID_PERCENT, d.DAY_OIL, d.DAY_REFUEL,
                d.DAY_MILEAGE, d.WORKHOUR_AVG_OIL, d.TRANSPORT_AVG_OIL,

                e.COMPANY, e.ESTATE, e.MACHINE_TYPE, e.MACHINE_CATEGORY, e.MACHINE_NO,
                e.BRAND_SPEC, e.ORDER_NO, e.PURCHASE_DATE, e.DRIVER_COUNT, e.CREATE_TIME,
                e.PURCH_PRICE, e.FUEL_DIFF,

                NOW()
            FROM `{DAILY_TABLE}` d
            LEFT JOIN `{DEVICE_TABLE}` e
                ON e.DEVICE_ID = d.{daily_device_col}
            WHERE d.DATE_STR = %s
            ON DUPLICATE KEY UPDATE
                -- 覆盖日表字段
                PROJECT_NAME       = VALUES(PROJECT_NAME),
                MECHANICAL_NO      = VALUES(MECHANICAL_NO),
                CAR_TYPE           = VALUES(CAR_TYPE),
                RENT_TYPE          = VALUES(RENT_TYPE),
                VALID_DURATION     = VALUES(VALID_DURATION),
                {report_idle_col}  = VALUES({report_idle_col}),
                VALID_PERCENT      = VALUES(VALID_PERCENT),
                DAY_OIL            = VALUES(DAY_OIL),
                DAY_REFUEL         = VALUES(DAY_REFUEL),
                DAY_MILEAGE        = VALUES(DAY_MILEAGE),
                WORKHOUR_AVG_OIL   = VALUES(WORKHOUR_AVG_OIL),
                TRANSPORT_AVG_OIL  = VALUES(TRANSPORT_AVG_OIL),

                -- 覆盖设备表字段
                COMPANY            = VALUES(COMPANY),
                ESTATE             = VALUES(ESTATE),
                MACHINE_TYPE       = VALUES(MACHINE_TYPE),
                MACHINE_CATEGORY   = VALUES(MACHINE_CATEGORY),
                MACHINE_NO         = VALUES(MACHINE_NO),
                BRAND_SPEC         = VALUES(BRAND_SPEC),
                ORDER_NO           = VALUES(ORDER_NO),
                PURCHASE_DATE      = VALUES(PURCHASE_DATE),
                DRIVER_COUNT       = VALUES(DRIVER_COUNT),
                CREATE_TIME        = VALUES(CREATE_TIME),
                PURCH_PRICE        = VALUES(PURCH_PRICE),
                FUEL_DIFF          = VALUES(FUEL_DIFF),

                -- 合并时间刷新
                INSERT_TIME        = NOW()
        """

        cur.execute(upsert_sql, (target_date,))
        affected = cur.rowcount
        print(
            f"[INFO] {REPORT_TABLE} 已写入/更新 {affected} 行 (DATE_STR={target_date})"
        )
        return affected


# ===== 主流程 =====
def main():
    target_date = get_target_date_str()
    print(f"[INFO] 开始任务，目标日期: {target_date}")
    print(f"[INFO] 使用 BASE_URL={BASE_URL}")
    print(
        f"[INFO] DB: host={DB_HOST}, port={DB_PORT}, user={DB_USER}, db={DB_NAME}"
    )

    # 1. 拉接口数据
    rows = fetch_data(target_date)
    if not rows:
        print("[WARN] 接口返回无数据/失败，跳过写入与合并。")
        return

    db_rows = [to_db_record(r, target_date) for r in rows if r.get("deviceNo")]

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
        upsert_daily(conn, db_rows)
        merge_daily_to_report(conn, target_date)

        conn.commit()
        print("[INFO] 全流程成功完成。")
    except Exception as e:
        conn.rollback()
        print(f"[ERROR] 发生异常，已回滚: {e}", file=sys.stderr)
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    main()
