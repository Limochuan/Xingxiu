#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
将日表 (xingxiu_daily) 指定日期的数据与设备表 (xingxiu_device_info) 合并写入业务表 (xingxiu_daily_report)。

特性：
- 目标日期来源优先级：环境变量 DATE_STR（Actions 手动输入） > 默认（按雅加达时区自动取“前一天”）
- DEVICE_ID 来自日表（优先 d.DEVICE_NO，若没有则用 d.DEVICE_ID）
- 怠速工时列名自动适配：IDLING_DURATION / IDING_DURATION
- 顺序：先日表字段 → 设备表字段 → 系统字段
- (DEVICE_ID, DATE_STR) 唯一；存在则覆盖更新并刷新 INSERT_TIME；AI字段（SCORE/SUMMARY/AVG_SPEED）不改
- 运行日志：控制台 + 本地文件 merge_daily.log
"""

import os
import sys
import pymysql
import logging
from datetime import datetime, timedelta

# Python 3.9+ 推荐使用 zoneinfo 计算时区
try:
    from zoneinfo import ZoneInfo
    _ZONEINFO_OK = True
except Exception:
    _ZONEINFO_OK = False

# ====== 日志配置 ======
logging.basicConfig(
    filename="merge_daily.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)
_console = logging.StreamHandler()
_console.setLevel(logging.INFO)
_console.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(_console)

def compute_default_jakarta_yesterday() -> str:
    """未指定 DATE_STR 时，按雅加达时间自动取“前一天”的 YYYY-MM-DD。"""
    if _ZONEINFO_OK:
        now_jkt = datetime.now(ZoneInfo("Asia/Jakarta"))
    else:
        # 兜底：近似用 UTC+7
        now_jkt = datetime.utcnow() + timedelta(hours=7)
    return (now_jkt.date() - timedelta(days=1)).isoformat()

def get_target_date() -> str:
    date_str = os.getenv("DATE_STR", "").strip()
    if not date_str:
        date_str = compute_default_jakarta_yesterday()
        logger.info(f"未提供 DATE_STR，按雅加达时区默认取前一天：{date_str}")
    # 校验格式
    try:
        datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        logger.error(f"DATE_STR 非法：{date_str}（期望 YYYY-MM-DD）")
        sys.exit(2)
    return date_str

def env(name, default=None):
    v = os.getenv(name)
    return v if v not in (None, "") else default

def first_existing_col(cur, table: str, candidates: list[str]) -> str:
    """返回 table 中第一个存在的列名（大小写不敏感）；若都不存在则抛错。"""
    cur.execute(f"SHOW COLUMNS FROM `{table}`")
    cols = {row[0].upper() for row in cur.fetchall()}
    for c in candidates:
        if c.upper() in cols:
            return c
    raise RuntimeError(f"表 `{table}` 缺少列：" + " / ".join(candidates))

def main():
    # --- 读取配置（全部来自环境变量，Actions 里用 secrets/env 注入） ---
    DB_HOST = env("DB_HOST")
    DB_PORT = int(env("DB_PORT", "3306"))
    DB_USER = env("DB_USER")
    DB_PASS = env("DB_PASS")
    DB_NAME = env("DB_NAME")

    DEVICE_TABLE = env("DEVICE_TABLE", "xingxiu_device_info")
    DAILY_TABLE  = env("DAILY_TABLE",  "xingxiu_daily")
    REPORT_TABLE = env("REPORT_TABLE", "xingxiu_daily_report")

    # 基本校验
    missing = [k for k in ["DB_HOST","DB_USER","DB_PASS","DB_NAME"] if not env(k)]
    if missing:
        logger.error(f"缺少必需的数据库环境变量：{', '.join(missing)}")
        sys.exit(2)

    TARGET_DATE = get_target_date()
    logger.info(f"目标日期: {TARGET_DATE}")

    conn = None
    try:
        conn = pymysql.connect(
            host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASS,
            database=DB_NAME, charset="utf8mb4", autocommit=False
        )
        with conn.cursor() as cur:
            # ---- 探测关键列名 ----
            # 日表中的设备列：优先 DEVICE_NO，没有就用 DEVICE_ID
            daily_device_col = first_existing_col(cur, DAILY_TABLE, ["DEVICE_NO", "DEVICE_ID"])
            # 日表/业务表中的 怠速工时 列：IDLING_DURATION 或 IDING_DURATION
            daily_idle_col   = first_existing_col(cur, DAILY_TABLE,  ["IDLING_DURATION", "IDING_DURATION"])
            report_idle_col  = first_existing_col(cur, REPORT_TABLE, ["IDLING_DURATION", "IDING_DURATION"])
            logger.info(f"日表设备列：{DAILY_TABLE}.{daily_device_col}")
            logger.info(f"怠速列：{DAILY_TABLE}.{daily_idle_col} -> {REPORT_TABLE}.{report_idle_col}")

            # ---- 确保唯一索引 (DEVICE_ID, DATE_STR) ----
            cur.execute(f"SHOW INDEX FROM `{REPORT_TABLE}` WHERE Key_name='uniq_device_date'")
            if not cur.fetchone():
                cur.execute(
                    f"ALTER TABLE `{REPORT_TABLE}` "
                    f"ADD UNIQUE KEY `uniq_device_date` (`DEVICE_ID`,`DATE_STR`)"
                )
                logger.info("已添加唯一索引 uniq_device_date")

            # ---- UPSERT（先日表 → 再设备表 → 系统字段）----
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
                    RENT_TYPE          = VALUES(REN T_TYPE),
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
            # 小修正：防止拼接时误插空格
            upsert_sql = upsert_sql.replace("REN T_TYPE", "RENT_TYPE")

            cur.execute(upsert_sql, (TARGET_DATE,))
            affected = cur.rowcount

        conn.commit()
        logger.info(f"[OK] {affected} 行已写入/更新 (DATE_STR={TARGET_DATE})")
        logger.info(f"[OK] 表：daily={DAILY_TABLE}, device={DEVICE_TABLE}, report={REPORT_TABLE}")
    except Exception:
        if conn:
            conn.rollback()
        logger.exception("执行失败")
        sys.exit(1)
    finally:
        if conn:
            conn.close()

if __name__ == "__main__":
    main()
