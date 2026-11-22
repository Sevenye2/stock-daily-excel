#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Baostock 日线增量导出 Excel（可推送到企业微信 / 飞书群）。

做什么：
1. 用 Baostock（免费官方源）拉取指定股票的日线，只用近 N 天的增量；
2. 与本地 Excel 里的历史行按 (日期, 股票代码) 去重后追加，重写整张表；
3. 涨跌额 / 涨跌幅 上色（涨红、跌绿、平灰）；
4. 配置了 WECHAT_WEBHOOK_URL 时往企业微信群发摘要，配置了
   FEISHU_WEBHOOK_URL 时往飞书群发卡片，两个渠道可同时开。

示例：
    python scripts/baostock_daily_excel.py --days 10
    python scripts/baostock_daily_excel.py --stocks 300750,688256 --days 5 --no-wechat

环境变量：
    STOCK_LIST          逗号分隔的股票代码（--stocks 优先）
    WECHAT_WEBHOOK_URL  企业微信群机器人 webhook（可选）
    FEISHU_WEBHOOK_URL  飞书群自定义机器人 webhook（可选）
    EXCEL_PATH          输出路径，默认 excel/daily_stock.xlsx
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import baostock as bs
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADERS = [
    "日期",
    "股票代码",
    "股票名称",
    "当日最低",
    "当日最高",
    "收盘",
    "涨跌额",
    "涨跌幅(%)",
    "涨/跌",
]

REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_EXCEL = str(REPO_ROOT / "excel" / "daily_stock.xlsx")

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
                "当日最低": round(low_f, 2),
                "当日最高": round(high_f, 2),
                "收盘": round(close_f, 2),
                "涨跌额": change,
                "涨跌幅(%)": round(pct_f, 2),
                "涨/跌": "涨" if pct_f > 0 else ("跌" if pct_f < 0 else "平"),
            }
        )
    return rows


def read_existing(path: Path) -> list[dict]:
    """读已有 Excel；表头若变过则按列名映射，缺列留空。"""
    if not path.exists():
        return []
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    header = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    rows: list[dict] = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        item: dict = {}
        for idx, name in enumerate(header):
            if name in HEADERS and idx < len(raw):
                item[name] = raw[idx]
        if not item.get("日期") or not item.get("股票代码"):
            continue
        item["日期"] = normalize_date(item["日期"])
        item["股票代码"] = str(item["股票代码"]).zfill(6)
        rows.append(item)
    return rows


def normalize_date(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    return str(value)[:10]


def write_excel(path: Path, rows: list[dict]) -> None:
    """整表重写：保证排序、格式、配色一致。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "日线数据"

    ws.append(HEADERS)
    for col_idx, name in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = (
            12 if name in ("日期", "股票代码", "股票名称", "涨/跌") else 11
        )
    ws.row_dimensions[1].height = 22

    for row in rows:
        ws.append([row.get(h) for h in HEADERS])

    for r in range(2, ws.max_row + 1):
        pct_cell = ws.cell(row=r, column=HEADERS.index("涨跌幅(%)") + 1)
        pct = pct_cell.value
        try:
            pct_val = float(pct) if pct is not None else 0.0
        except (TypeError, ValueError):
            pct_val = 0.0

        if pct_val > 0:
            font_color, fill_color = UP_COLOR, UP_FILL
        elif pct_val < 0:
            font_color, fill_color = DOWN_COLOR, DOWN_FILL
        else:
            font_color, fill_color = FLAT_COLOR, FLAT_FILL

        fill = PatternFill("solid", fgColor=fill_color)
        for col_name in ("涨跌额", "涨跌幅(%)", "涨/跌"):
            cell = ws.cell(row=r, column=HEADERS.index(col_name) + 1)
            cell.font = Font(color=font_color, bold=col_name == "涨/跌")
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center")

        ws.cell(row=r, column=HEADERS.index("日期") + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=HEADERS.index("股票代码") + 1).alignment = Alignment(horizontal="center")
        for col_name in ("当日最低", "当日最高", "收盘"):
            ws.cell(row=r, column=HEADERS.index(col_name) + 1).number_format = "0.00"
        ws.cell(row=r, column=HEADERS.index("涨跌额") + 1).number_format = "+0.00;-0.00;0.00"
        # 涨跌幅存的是 2.35 这种百分数，显示成 +2.35%
        ws.cell(row=r, column=HEADERS.index("涨跌幅(%)") + 1).number_format = '+0.00"%";-0.00"%";0.00"%"'
        for col_idx in range(1, len(HEADERS) + 1):
            ws.cell(row=r, column=col_idx).border = BORDER

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{ws.max_row}"
    wb.save(path)


def write_csv(path: Path, rows: list[dict]) -> None:
    """同时导出一份 CSV：GitHub 网页上能直接点开看表格，不用下载 xlsx。"""
    import csv

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=HEADERS, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({h: row.get(h) for h in HEADERS})


def send_wechat(webhook: str, text: str) -> None:
    """企业微信群机器人：markdown 单条上限 4096 字节，超了直接截断。"""
    if not webhook:
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


def build_summary(date_str: str, new_rows: list[dict], total: int, excel_path: Path) -> str:
    ups = [r for r in new_rows if r["涨/跌"] == "涨"]
    downs = [r for r in new_rows if r["涨/跌"] == "跌"]
    flats = [r for r in new_rows if r["涨/跌"] == "平"]

    lines = [
        f"## 股票日线更新 {date_str}",
        f"> 新增 **{len(new_rows)}** 行 ｜ 表内共 **{total}** 行",
        f"> 涨 {len(ups)} ｜ 跌 {len(downs)} ｜ 平 {len(flats)}",
    ]
    if new_rows:
        top_up = sorted([r for r in new_rows if r["涨跌幅(%)"] is not None], key=lambda x: x["涨跌幅(%)"], reverse=True)[:3]
        top_down = sorted([r for r in new_rows if r["涨跌幅(%)"] is not None], key=lambda x: x["涨跌幅(%)"])[:3]
        if top_up and top_up[0]["涨跌幅(%)"] > 0:
            lines.append("**涨幅前三**")
            for r in top_up:
                if r["涨跌幅(%)"] <= 0:
                    break
                lines.append(f"- {r['股票代码']} {r['股票名称'] or ''} {r['涨跌幅(%)']:+.2f}%（{r['涨跌额']:+.2f}）")
        if top_down and top_down[0]["涨跌幅(%)"] < 0:
            lines.append("**跌幅前三**")
            for r in top_down:
                if r["涨跌幅(%)"] >= 0:
                    break
                lines.append(f"- {r['股票代码']} {r['股票名称'] or ''} {r['涨跌幅(%)']:+.2f}%（{r['涨跌额']:+.2f}）")
    else:
        lines.append("今日没有新数据（非交易日，或 Baostock 尚未更新当日行情）。")

    lines.append(f"<font color=\"comment\">文件：{excel_path.name}</font>")
    return "\n".join(lines)


def build_feishu_card(date_str: str, new_rows: list[dict], total: int, excel_path: Path) -> dict:
    """飞书群机器人卡片：一只股票一行，字段与 Excel 保持一致。"""
    ups = [r for r in new_rows if r["涨/跌"] == "涨"]
    downs = [r for r in new_rows if r["涨/跌"] == "跌"]
    flats = [r for r in new_rows if r["涨/跌"] == "平"]

    elements: list[dict] = []
    if new_rows:
        head = f"**{len(new_rows)} 只 ｜ 涨 {len(ups)} · 跌 {len(downs)} · 平 {len(flats)}**\n表内累计 **{total}** 行"
    else:
        head = f"**今日无新增**（非交易日，或 Baostock 尚未更新）\n最近数据日 {date_str} ｜ 表内累计 **{total}** 行"
    elements.append({"tag": "div", "text": {"tag": "lark_md", "content": head}})
    elements.append({"tag": "hr"})

    if new_rows:
        ordered = sorted(new_rows, key=lambda r: r["涨跌幅(%)"] if r["涨跌幅(%)"] is not None else 0, reverse=True)
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
        return
    try:
        import requests

        resp = requests.post(webhook, json=payload, timeout=15)
        ok = resp.status_code == 200 and '"code":0' in resp.text.replace(" ", "")
        print(f"[飞书] {'✅' if ok else '❌'} HTTP {resp.status_code} {resp.text[:200]}")
    except Exception as exc:  # noqa: BLE001 - 推送失败不该让整个任务失败
        print(f"[飞书] 推送失败（不影响数据）: {exc}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Baostock 日线增量导出 Excel")
    parser.add_argument("--stocks", default=os.getenv("STOCK_LIST", ""), help="逗号分隔的股票代码")
    parser.add_argument("--days", type=int, default=10, help="回溯天数（默认 10，足够覆盖周末与节假日）")
    parser.add_argument("--excel", default=os.getenv("EXCEL_PATH", DEFAULT_EXCEL), help="Excel 输出路径")
    parser.add_argument("--no-wechat", action="store_true", help="不推送企业微信")
    parser.add_argument("--no-feishu", action="store_true", help="不推送飞书")
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

    existing = read_existing(excel_path)
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
    all_rows.sort(key=lambda r: (str(r["日期"]), str(r["股票代码"])), reverse=True)
    write_excel(excel_path, all_rows)
    write_csv(excel_path.with_suffix(".csv"), all_rows)
    print(f"[写出] {excel_path} 共 {len(all_rows)} 行（本次新增 {len(new_rows)}）")

    # 没有新增行时（周末/节假日），标题里的日期回退到表内最近的数据日，避免显示成"今天"
    date_str = max((r["日期"] for r in new_rows), default="") or max(
        (r["日期"] for r in all_rows), default=end_date
    )
    summary = build_summary(date_str, new_rows, len(all_rows), excel_path)
    print("\n" + summary + "\n")

    if not args.no_wechat:
        send_wechat(os.getenv("WECHAT_WEBHOOK_URL", "").strip(), summary)

    if not args.no_feishu:
        send_feishu(
            os.getenv("FEISHU_WEBHOOK_URL", "").strip(),
            build_feishu_card(date_str, new_rows, len(all_rows), excel_path),
        )

    return 0


if __name__ == "__main__":
    sys.exit(main())
