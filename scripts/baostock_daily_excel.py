#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Baostock 日线增量导出 Excel（可推送到企业微信 / 飞书群）。

做什么：
1. 用 Baostock（免费官方源）拉取指定股票的日线，只取近 N 天的增量；
2. 与历史数据按 (日期, 股票代码) 去重后追加。历史存在 data/daily.csv（长表，机器用），
   Excel 是它的"宽表"视图：一行一个日期，每只股票占 7 列（代码/名称/最高/最低/收盘/涨跌额/涨跌幅），
   股票之间用一条竖线隔开；
3. 涨跌额 / 涨跌幅 上色（涨红、跌绿、平灰）；
4. 有新数据时推送：配了 WECHAT_WEBHOOK_URL 发企业微信，配了 FEISHU_WEBHOOK_URL 发飞书卡片；
   半夜（默认 00:00–08:00，北京时间的 QUIET_UNTIL）只写表不打扰，攒到白天第一次运行再补发。

为什么这样设计：
GitHub 的定时任务可能被延迟、甚至整次丢弃（官方文档明确"部分排队任务会被丢弃"），
所以不跟它较劲——每次被唤醒都跑一次，靠"去重"保证幂等：有新数据就入库，没有就什么都不变。

示例：
    python scripts/baostock_daily_excel.py --days 10
    python scripts/baostock_daily_excel.py --stocks 300750,688256 --days 5 --no-feishu

环境变量：
    STOCK_LIST          逗号分隔的股票代码（--stocks 优先）
    WECHAT_WEBHOOK_URL  企业微信群机器人 webhook（可选）
    FEISHU_WEBHOOK_URL  飞书群自定义机器人 webhook（可选）
    EXCEL_PATH          输出路径，默认 excel/daily_stock.xlsx
    QUIET_UNTIL         北京时间几点之前不推送，默认 08:00
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import baostock as bs
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_EXCEL = str(REPO_ROOT / "excel" / "daily_stock.xlsx")
# 长表（机器可读的真正数据源）：一行 = 一只股票的一天
CANON_CSV = REPO_ROOT / "data" / "daily.csv"
# 攒着没推送的日期（半夜不打扰，白天补发）
PENDING_FILE = REPO_ROOT / "state" / "pending_notify.json"
CN_TZ = timezone(timedelta(hours=8))   # 北京时间；项目按北京日期组织数据

# 长表列（也是 CSV 的列顺序）
CANON_HEADERS = [
    "日期",
    "股票代码",
    "股票名称",
    "当日最高",
    "当日最低",
    "收盘",
    "涨跌额",
    "涨跌幅(%)",
]

# 宽表里每只股票的字段顺序（用户指定：代码 名称 最高 最低 收盘 涨跌额 涨跌幅）
FIELDS = ["代码", "名称", "最高", "最低", "收盘", "涨跌额", "涨跌幅"]
FIELD_TO_KEY = {
    "代码": "股票代码",
    "名称": "股票名称",
    "最高": "当日最高",
    "最低": "当日最低",
    "收盘": "收盘",
    "涨跌额": "涨跌额",
    "涨跌幅": "涨跌幅(%)",
}
NUMBER_KEYS = ("当日最高", "当日最低", "收盘", "涨跌额", "涨跌幅(%)")

# Baostock 日线字段：pctChg 是涨跌幅（%），preclose 是前收盘
BAOSTOCK_FIELDS = "date,code,open,high,low,close,preclose,volume,amount,pctChg"

UP_COLOR = "C00000"        # 涨：红字
UP_FILL = "FDE9E9"
DOWN_COLOR = "0B7A0B"      # 跌：绿字
DOWN_FILL = "E8F5E9"
FLAT_COLOR = "595959"      # 平：灰
FLAT_FILL = "F2F2F2"

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
SEP_FILL = PatternFill("solid", fgColor="BFBFBF")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def parse_stock_list(raw: str) -> list[str]:
    """把 '300750,688836; sh600519' 之类统统拆成裸 6 位代码。"""
    raw = (raw or "").replace("，", ",").replace("；", ",").replace(";", ",")
    codes: list[str] = []
    for part in raw.replace("\n", ",").split(","):
        code = part.strip().upper().replace("SH", "").replace("SZ", "").replace(".", "")
        if not code:
            continue
        if not code.isdigit() or len(code) != 6:
            print(f"[跳过] 无法识别的股票代码: {part!r}（本脚本只支持 A 股 6 位代码）")
            continue
        if code not in codes:
            codes.append(code)
    return codes


def to_baostock_code(code: str) -> str | None:
    """6 位裸码 -> baostock 的 sh./sz. 前缀格式。北交所 Baostock 不提供，返回 None。"""
    if len(code) != 6 or not code.isdigit():
        return None
    if code[0] in ("5", "6", "9"):        # 沪市 A/B 股、ETF
        return f"sh.{code}"
    if code[0] in ("0", "1", "2", "3"):   # 深市主板/创业板/ETF
        return f"sz.{code}"
    return None                            # 4/8 开头多为北交所，Baostock 无此数据


def fetch_name(bs_code: str) -> str:
    """查股票名称，失败就返回空串（不影响主流程）。"""
    try:
        rs = bs.query_stock_basic(code=bs_code)
        if rs.error_code != "0":
            return ""
        while rs.next():
            row = rs.get_row_data()
            if len(row) >= 2:
                return row[1]
    except Exception:  # noqa: BLE001 - 名称只是锦上添花
        pass
    return ""


def fetch_daily(bs_code: str, start: str, end: str) -> list[dict]:
    """拉取一只股票在 [start, end] 的日线，返回标准 dict 列表。"""
    rs = bs.query_history_k_data_plus(
        bs_code,
        BAOSTOCK_FIELDS,
        start_date=start,
        end_date=end,
        frequency="d",
        adjustflag="3",  # 3=不复权：最低/最高/收盘即真实成交价
    )
    if rs.error_code != "0":
        print(f"[错误] {bs_code} 拉取失败: {rs.error_code} {rs.error_msg}")
        return []

    rows: list[dict] = []
    while rs.next():
        r = rs.get_row_data()
        if len(r) < 10:
            continue
        date_s, _code, _open, high, low, close, preclose, _vol, _amt, pct = r[:10]
        # 停牌日 Baostock 返回空字符串，直接跳过
        if not date_s or not close or not high or not low:
            continue
        try:
            close_f = float(close)
            high_f = float(high)
            low_f = float(low)
            pre_f = float(preclose) if preclose else None
            pct_f = float(pct) if pct else None
        except ValueError:
            continue

        # 涨跌额优先 close-preclose（Baostock 的 preclose 已按除权口径处理）
        if pre_f is not None:
            change = round(close_f - pre_f, 2)
        elif pct_f is not None:
            change = round(close_f * pct_f / (100 + pct_f), 2)
        else:
            change = None

        # 涨跌幅优先用官方 pctChg，缺失时用 close/preclose 反算
        if pct_f is None and pre_f:
            pct_f = round((close_f - pre_f) / pre_f * 100, 2)
        if pct_f is None:
            continue

        rows.append(
            {
                "日期": date_s,
                "股票代码": bs_code.split(".")[-1],
                "股票名称": "",
                "当日最高": round(high_f, 2),
                "当日最低": round(low_f, 2),
                "收盘": round(close_f, 2),
                "涨跌额": change,
                "涨跌幅(%)": round(pct_f, 2),
                "涨/跌": "涨" if pct_f > 0 else ("跌" if pct_f < 0 else "平"),
            }
        )
    return rows


def normalize_date(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    return text[:10]


def to_float(value):
    try:
        if value is None or value == "":
            return None
        return round(float(value), 2)
    except (TypeError, ValueError):
        return None


def clean_row(item: dict) -> dict | None:
    """校验并规范化一条长表数据；缺日期或代码就丢弃。"""
    date_s = normalize_date(item.get("日期") or "")
    code = str(item.get("股票代码") or "").strip()
    if not date_s or not code:
        return None
    code = code.split(".")[-1].zfill(6)
    row = {
        "日期": date_s,
        "股票代码": code,
        "股票名称": str(item.get("股票名称") or "").strip(),
    }
    for key in NUMBER_KEYS:
        row[key] = to_float(item.get(key))
    pct = row["涨跌幅(%)"]
    row["涨/跌"] = "涨" if (pct or 0) > 0 else ("跌" if (pct or 0) < 0 else "平")
    return row


def read_canonical(path: Path = CANON_CSV) -> list[dict]:
    """读长表 CSV（真正的数据源）。不存在就返回空列表。"""
    if not path.exists():
        return []
    rows: list[dict] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        for item in csv.DictReader(f):
            row = clean_row(item)
            if row:
                rows.append(row)
    return rows


def write_canonical(rows: list[dict], path: Path = CANON_CSV) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CANON_HEADERS, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k) for k in CANON_HEADERS})


def read_legacy_excel(path: Path) -> list[dict]:
    """兼容老格式的老表：把 Excel 读成长表数据（用于一次性迁移）。

    支持三种表头：
      A. 「股票」合并列（9 列，代码+名称挤在一格）
      B. 股票代码 / 股票名称 分开（8~9 列，长表）
      C. 宽表（一行一个日期，每只股票 7 列）
    """
    if not path.exists():
        return []
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    first = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    second = [str(c.value).strip() if c.value is not None else "" for c in ws[2]] if ws.max_row >= 2 else []

    # C. 宽表：第 2 行是字段名
    if "日期" not in first and "代码" in second:
        return _read_wide(ws, first, second)

    # A/B. 长表
    header = first
    rows: list[dict] = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        item: dict = {}
        for idx, name in enumerate(header):
            if idx < len(raw):
                item[name] = raw[idx]
        if "股票" in item:                       # A：拆回代码 + 名称
            label = str(item.pop("股票") or "").strip()
            parts = label.split(None, 1)
            item["股票代码"] = parts[0] if parts else ""
            item["股票名称"] = item.get("股票名称") or (parts[1] if len(parts) > 1 else "")
        row = clean_row(item)
        if row:
            rows.append(row)
    return rows


def _read_wide(ws, first: list[str], second: list[str]) -> list[dict]:
    """把宽表还原成长表数据。"""
    rows: list[dict] = []
    starts = [i for i, name in enumerate(second) if name == "代码"]
    for raw in ws.iter_rows(min_row=3, values_only=True):
        if not raw:
            continue
        date_s = normalize_date(raw[0] if raw else "")
        if not date_s:
            continue
        for start in starts:
            code = raw[start] if start < len(raw) else None
            if not code:
                continue
            item = {"日期": date_s, "股票代码": code}
            for offset, field in enumerate(FIELDS):
                idx = start + offset
                if field == "代码":
                    continue
                item[FIELD_TO_KEY[field]] = raw[idx] if idx < len(raw) else None
            row = clean_row(item)
            if row:
                rows.append(row)
    return rows


def load_rows(excel_path: Path) -> list[dict]:
    """优先读长表；没有长表就从老 Excel 迁移一次。"""
    rows = read_canonical()
    if rows:
        return rows
    rows = read_legacy_excel(excel_path)
    if rows:
        print(f"[迁移] 从 {excel_path.name} 读到 {len(rows)} 行，转存到 data/daily.csv")
        write_canonical(rows)
    return rows


def write_excel(path: Path, rows: list[dict], codes: list[str]) -> None:
    """宽表重写：一行一个日期，每只股票 7 列，股票之间一条竖线。"""
    path.parent.mkdir(parents=True, exist_ok=True)

    names: dict[str, str] = {}
    for row in rows:
        if row.get("股票名称"):
            names[row["股票代码"]] = row["股票名称"]
    order = list(codes)
    for code in sorted({r["股票代码"] for r in rows}):
        if code not in order:
            order.append(code)

    dates = sorted({r["日期"] for r in rows})
    by_key = {(r["日期"], r["股票代码"]): r for r in rows}

    wb = Workbook()
    ws = wb.active
    ws.title = "日线数据"

    # ---- 两行表头：第一行按股票分组，第二行是字段名 ----
    ws.cell(row=1, column=1, value="日期")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    col = 2
    sep_cols: list[int] = []
    for i, code in enumerate(order):
        start = col
        label = f"{code} {names.get(code, '')}".strip()
        ws.cell(row=1, column=start, value=label)
        ws.merge_cells(
            start_row=1, start_column=start, end_row=1, end_column=start + len(FIELDS) - 1
        )
        for j, field in enumerate(FIELDS):
            ws.cell(row=2, column=start + j, value=field)
        col = start + len(FIELDS)
        if i < len(order) - 1:      # 股票之间插一条窄的竖线
            ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
            sep_cols.append(col)
            col += 1

    # ---- 表头样式 ----
    for r in (1, 2):
        for c in range(1, col):
            cell = ws.cell(row=r, column=c)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
    for c in sep_cols:
        for r in (1, 2):
            ws.cell(row=r, column=c).fill = SEP_FILL
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 18

    # ---- 列宽 ----
    ws.column_dimensions["A"].width = 12
    for i, _code in enumerate(order):
        start = 2 + i * (len(FIELDS) + 1)
        widths = [10, 13, 8, 8, 9, 8, 9]     # 代码 名称 最高 最低 收盘 涨跌额 涨跌幅
        for j, width in enumerate(widths):
            ws.column_dimensions[get_column_letter(start + j)].width = width
    for c in sep_cols:
        ws.column_dimensions[get_column_letter(c)].width = 2

    # ---- 数据 ----
    for r, date_s in enumerate(dates, start=3):
        date_cell = ws.cell(row=r, column=1, value=date_s)
        date_cell.alignment = Alignment(horizontal="center")
        date_cell.border = BORDER
        for i, code in enumerate(order):
            start = 2 + i * (len(FIELDS) + 1)
            item = by_key.get((date_s, code))
            for j, field in enumerate(FIELDS):
                key = FIELD_TO_KEY[field]
                cell = ws.cell(row=r, column=start + j)
                cell.value = item.get(key) if item else None
                cell.border = BORDER
                if field in ("代码", "名称"):
                    cell.alignment = Alignment(horizontal="left" if field == "名称" else "center")
                elif field in ("最高", "最低", "收盘"):
                    cell.number_format = "0.00"
                    cell.alignment = Alignment(horizontal="center")
                elif field == "涨跌额":
                    cell.number_format = "+0.00;-0.00;0.00"
                    cell.alignment = Alignment(horizontal="center")
                else:  # 涨跌幅
                    cell.number_format = '+0.00"%";-0.00"%";0.00"%"'
                    cell.alignment = Alignment(horizontal="center")

            if not item:
                continue
            pct = item.get("涨跌幅(%)") or 0
            if pct > 0:
                font_color, fill_color = UP_COLOR, UP_FILL
            elif pct < 0:
                font_color, fill_color = DOWN_COLOR, DOWN_FILL
            else:
                font_color, fill_color = FLAT_COLOR, FLAT_FILL
            fill = PatternFill("solid", fgColor=fill_color)
            for field in ("涨跌额", "涨跌幅"):
                cell = ws.cell(row=r, column=start + FIELDS.index(field))
                cell.font = Font(color=font_color)
                cell.fill = fill

    # 日期列冻结（滚到右边时始终能看到是哪天），表头两行也一起冻住
    ws.freeze_panes = "B3"
    wb.save(path)


def write_wide_csv(path: Path, rows: list[dict], codes: list[str]) -> None:
    """和 Excel 同布局的 CSV：GitHub 网页上直接能看，不用下载。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    names: dict[str, str] = {}
    for row in rows:
        if row.get("股票名称"):
            names[row["股票代码"]] = row["股票名称"]
    order = list(codes)
    for code in sorted({r["股票代码"] for r in rows}):
        if code not in order:
            order.append(code)

    header: list[str] = ["日期"]
    for code in order:
        header.append(f"{code} {names.get(code, '')}".strip())
        header.extend(FIELDS[1:])
        header.append("|")              # 竖线占位符，和 Excel 里的分隔列对应
    if len(order) > 1:
        header = header[:-1]

    by_key = {(r["日期"], r["股票代码"]): r for r in rows}
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        for date_s in sorted({r["日期"] for r in rows}):
            line: list = [date_s]
            for i, code in enumerate(order):
                item = by_key.get((date_s, code))
                line.append(item.get("股票代码") if item else "")
                for field in FIELDS[1:]:
                    value = item.get(FIELD_TO_KEY[field]) if item else ""
                    line.append("" if value is None else value)
                if i < len(order) - 1:
                    line.append("|")
            writer.writerow(line)


def send_wechat(webhook: str, text: str) -> None:
    """企业微信群机器人：markdown 单条上限 4096 字节，超了直接截断。"""
    if not webhook:
        print("[企业微信] 跳过：没有配置 WECHAT_WEBHOOK_URL")
        return
    payload_bytes = text.encode("utf-8")
    if len(payload_bytes) > 4000:
        text = payload_bytes[:3900].decode("utf-8", "ignore") + "\n…(已截断)"
    try:
        import requests

        resp = requests.post(
            webhook,
            json={"msgtype": "markdown", "markdown": {"content": text}},
            timeout=15,
        )
        print(f"[企业微信] HTTP {resp.status_code} {resp.text[:200]}")
    except Exception as exc:  # noqa: BLE001 - 推送失败不该让整个任务失败
        print(f"[企业微信] 推送失败（不影响数据）: {exc}")


def build_summary(date_str: str, rows: list[dict], total: int, excel_path: Path) -> str:
    ups = [r for r in rows if r["涨/跌"] == "涨"]
    downs = [r for r in rows if r["涨/跌"] == "跌"]
    flats = [r for r in rows if r["涨/跌"] == "平"]

    lines = [
        f"## 股票日线更新 {date_str}",
        f"> 本次 **{len(rows)}** 只 ｜ 表内共 **{total}** 行",
        f"> 涨 {len(ups)} ｜ 跌 {len(downs)} ｜ 平 {len(flats)}",
    ]
    if rows:
        ordered = sorted(rows, key=lambda x: x["涨跌幅(%)"] or 0, reverse=True)
        top_up = [r for r in ordered if (r["涨跌幅(%)"] or 0) > 0][:3]
        top_down = [r for r in ordered[::-1] if (r["涨跌幅(%)"] or 0) < 0][:3]
        if top_up:
            lines.append("**涨幅前三**")
            for r in top_up:
                lines.append(f"- {r['股票代码']} {r['股票名称']} {r['涨跌幅(%)']:+.2f}%（{r['涨跌额']:+.2f}）")
        if top_down:
            lines.append("**跌幅前三**")
            for r in top_down:
                lines.append(f"- {r['股票代码']} {r['股票名称']} {r['涨跌幅(%)']:+.2f}%（{r['涨跌额']:+.2f}）")
    else:
        lines.append("本次没有新数据（非交易日，或 Baostock 尚未更新当日行情）。")

    lines.append(f"<font color=\"comment\">文件：{excel_path.name}</font>")
    return "\n".join(lines)


def build_feishu_card(date_str: str, rows: list[dict], total: int, excel_path: Path) -> dict:
    """飞书群机器人卡片：一只股票一行，字段与 Excel 保持一致。"""
    ups = [r for r in rows if r["涨/跌"] == "涨"]
    downs = [r for r in rows if r["涨/跌"] == "跌"]
    flats = [r for r in rows if r["涨/跌"] == "平"]

    elements: list[dict] = []
    head = f"**{len(rows)} 只 ｜ 涨 {len(ups)} · 跌 {len(downs)} · 平 {len(flats)}**\n表内累计 **{total}** 行"
    elements.append({"tag": "div", "text": {"tag": "lark_md", "content": head}})
    elements.append({"tag": "hr"})

    if rows:
        ordered = sorted(rows, key=lambda r: r["涨跌幅(%)"] if r["涨跌幅(%)"] is not None else 0, reverse=True)
        for r in ordered:
            pct = r["涨跌幅(%)"] or 0.0
            icon = "🔴" if pct > 0 else ("🟢" if pct < 0 else "⚪")
            change = r["涨跌额"]
            change_s = f"{change:.2f}" if isinstance(change, (int, float)) else "—"
            elements.append(
                {
                    "tag": "div",
                    "text": {
                        "tag": "lark_md",
                        "content": (
                            f"{icon} **{r['股票代码']} {r['股票名称'] or ''}**　涨跌 {pct:+.2f}%（{change_s}）　"
                            f"高 {r['当日最高']:.2f} · 低 {r['当日最低']:.2f}　收盘 {r['收盘']:.2f}"
                        ),
                    },
                }
            )
    elements.append({"tag": "hr"})
    elements.append(
        {
            "tag": "note",
            "elements": [
                {"tag": "plain_text", "content": f"数据源 Baostock · {date_str} 收盘 · 完整表格见仓库 {excel_path.name}"}
            ],
        }
    )

    return {
        "msg_type": "interactive",
        "card": {
            "config": {"wide_screen_mode": True},
            "header": {
                "template": "blue",
                "title": {"tag": "plain_text", "content": f"📈 每日行情速览 · {date_str}"},
            },
            "elements": elements,
        },
    }


def send_feishu(webhook: str, payload: dict) -> None:
    """飞书群自定义机器人：卡片消息，失败只打印不抛异常。"""
    if not webhook:
        print("[飞书] 跳过：没有配置 FEISHU_WEBHOOK_URL（GitHub 里检查 Settings → Secrets and variables → Actions → Secrets）")
        return
    try:
        import requests

        resp = requests.post(webhook, json=payload, timeout=15)
        ok = resp.status_code == 200 and '"code":0' in resp.text.replace(" ", "")
        print(f"[飞书] {'✅' if ok else '❌'} HTTP {resp.status_code} {resp.text[:200]}")
    except Exception as exc:  # noqa: BLE001 - 推送失败不该让整个任务失败
        print(f"[飞书] 推送失败（不影响数据）: {exc}")


def load_pending(path: Path = PENDING_FILE) -> list[str]:
    """读取攒着没推送的日期。"""
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return [str(d) for d in data.get("dates", [])]
    except (OSError, ValueError):
        return []


def save_pending(dates: list[str], path: Path = PENDING_FILE) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not dates:
        path.unlink(missing_ok=True)
        return
    path.write_text(json.dumps({"dates": sorted(set(dates))}, ensure_ascii=False) + "\n", encoding="utf-8")


def quiet_until_hour(raw: str) -> int:
    """'08:00' -> 8；格式不对就退回 8。"""
    try:
        return int(str(raw).strip().split(":")[0])
    except (TypeError, ValueError, IndexError):
        return 8


def is_wide_excel(path: Path) -> bool:
    """现有 xlsx 是不是新的宽表布局（第二行是 代码/名称/最高...）。"""
    if not path.exists():
        return False
    try:
        wb = load_workbook(path, read_only=True)
        ws = wb.active
        row2 = [c.value for c in next(ws.iter_rows(min_row=2, max_row=2), [])]
        wb.close()
        return "代码" in row2 and "涨跌幅" in row2
    except Exception:  # noqa: BLE001 - 读不了就当需要重写
        return False


def main() -> int:
    parser = argparse.ArgumentParser(description="Baostock 日线增量导出 Excel")
    parser.add_argument("--stocks", default=os.getenv("STOCK_LIST", ""), help="逗号分隔的股票代码")
    parser.add_argument("--days", type=int, default=10, help="回溯天数（默认 10，足够覆盖周末与节假日）")
    parser.add_argument("--excel", default=os.getenv("EXCEL_PATH", DEFAULT_EXCEL), help="Excel 输出路径")
    parser.add_argument("--no-wechat", action="store_true", help="不推送企业微信")
    parser.add_argument("--no-feishu", action="store_true", help="不推送飞书")
    parser.add_argument("--notify-empty", action="store_true", help="本次没有新数据也推送（本地测试用）")
    parser.add_argument("--force-notify", action="store_true", help="忽略安静时段，立即推送")
    parser.add_argument("--rebuild", action="store_true", help="没有新数据也按当前布局重画表格（改样式后用）")
    args = parser.parse_args()

    codes = parse_stock_list(args.stocks)
    if not codes:
        print("没有可用的股票代码：请用 --stocks 或环境变量 STOCK_LIST 指定", file=sys.stderr)
        return 2

    excel_path = Path(args.excel)
    end_date = datetime.now().strftime("%Y-%m-%d")
    start_date = (datetime.now() - timedelta(days=max(args.days, 1))).strftime("%Y-%m-%d")

    print(f"[配置] 股票 {len(codes)} 只: {','.join(codes)}")
    print(f"[配置] 日期范围 {start_date} ~ {end_date}，Excel: {excel_path}")

    existing = load_rows(excel_path)
    seen = {(r["日期"], r["股票代码"]) for r in existing}
    print(f"[增量] 已有 {len(existing)} 行")

    login = bs.login()
    if login.error_code != "0":
        print(f"[错误] Baostock 登录失败: {login.error_code} {login.error_msg}", file=sys.stderr)
        return 3

    new_rows: list[dict] = []
    name_cache: dict[str, str] = {}
    try:
        for idx, code in enumerate(codes, start=1):
            bs_code = to_baostock_code(code)
            if bs_code is None:
                print(f"[跳过] {code}: Baostock 不支持该代码段（如北交所）")
                continue
            rows = fetch_daily(bs_code, start_date, end_date)
            if not rows:
                print(f"[{idx}/{len(codes)}] {code}: 无数据")
                continue
            name = name_cache.get(code) or fetch_name(bs_code)
            name_cache[code] = name
            added = 0
            for row in rows:
                row["股票名称"] = name
                key = (row["日期"], row["股票代码"])
                if key in seen:
                    continue
                seen.add(key)
                new_rows.append(row)
                added += 1
            print(f"[{idx}/{len(codes)}] {code} {name}: 拉取 {len(rows)} 行，新增 {added} 行")
            time.sleep(0.3)  # 轻微限速，Baostock 虽然宽松但没必要猛打
    finally:
        bs.logout()

    all_rows = existing + new_rows
    all_rows.sort(key=lambda r: (str(r["日期"]), str(r["股票代码"])))

    # 有新数据就重写；没新数据但表还是老布局（或文件不存在）时也重写一次，
    # 这样改表结构后不需要手动补跑。
    if new_rows or args.rebuild or not is_wide_excel(excel_path):
        write_canonical(all_rows)
        write_excel(excel_path, all_rows, codes)
        write_wide_csv(excel_path.with_suffix(".csv"), all_rows, codes)
        print(f"[写出] {excel_path} 共 {len(all_rows)} 行（本次新增 {len(new_rows)}）")
    else:
        print("[写出] 没有新增数据，跳过重写（避免无意义的提交）")

    # ---- 推送：只在有新数据时打扰；半夜先攒着，白天补发 ----
    pending = load_pending()
    new_dates = sorted({r["日期"] for r in new_rows})
    if new_dates:
        pending = sorted(set(pending) | set(new_dates))

    notify_rows: list[dict] = []
    if pending:
        wanted = set(pending)
        notify_rows = [r for r in all_rows if r["日期"] in wanted]

    hour = datetime.now(CN_TZ).hour
    limit = quiet_until_hour(os.getenv("QUIET_UNTIL", "08:00"))
    quiet = hour < limit and not args.force_notify

    should_notify = bool(notify_rows) or args.notify_empty
    if quiet and should_notify:
        save_pending(pending)
        print(f"[推送] 现在北京时间 {hour} 点，属于安静时段（{limit} 点前不打扰），{len(pending)} 个日期攒着白天再发")
        should_notify = False

    if should_notify:
        date_str = max(pending, default="") or max((r["日期"] for r in all_rows), default=end_date)
        summary = build_summary(date_str, notify_rows, len(all_rows), excel_path)
        print("\n" + summary + "\n")
        if not args.no_wechat:
            send_wechat(os.getenv("WECHAT_WEBHOOK_URL", "").strip(), summary)
        if not args.no_feishu:
            send_feishu(
                os.getenv("FEISHU_WEBHOOK_URL", "").strip(),
                build_feishu_card(date_str, notify_rows, len(all_rows), excel_path),
            )
        save_pending([])
    else:
        if not notify_rows:
            print("[推送] 没有新数据，不推送")

    return 0


if __name__ == "__main__":
    sys.exit(main())
