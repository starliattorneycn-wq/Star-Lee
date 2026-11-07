#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""模拟富豪账本：虚拟消费、虚拟持仓、真实行情快照、周期盈亏报告。

所有操作都只写本地 JSON 账本，不产生任何真实交易。
账本默认路径 ~/.tycoon/ledger.json，可用 --ledger 或环境变量 TYCOON_LEDGER 覆盖。
"""

import argparse
import json
import os
import sys
import uuid
from datetime import datetime, timedelta, timezone

TZ = timezone(timedelta(hours=8))  # 固定按北京时间切分自然周/月
args_holder = None


def tycoon_home():
    """账本目录：默认 ~/.tycoon，可用 TYCOON_HOME 覆盖。"""
    return os.path.abspath(os.environ.get("TYCOON_HOME") or os.path.expanduser("~/.tycoon"))


def default_ledger_path():
    return os.path.join(tycoon_home(), "ledger.json")


def default_xlsx_path():
    return os.path.join(tycoon_home(), "账本.xlsx")


def config_path():
    return os.path.join(tycoon_home(), "config.json")


def memo_path():
    return os.path.join(tycoon_home(), "秘书备忘.md")


# ---------- 基础工具 ----------

def now():
    return datetime.now(TZ)


def now_iso():
    return now().isoformat(timespec="seconds")


def today():
    return now().date().isoformat()


def parse_date(s):
    return datetime.fromisoformat(s).astimezone(TZ)


def xlsx_path(args):
    return os.path.abspath(getattr(args, "excel", None) or os.environ.get("TYCOON_XLSX") or default_xlsx_path())


def week_summary(data):
    start = period_start("week")
    buys = [p for p in data["purchases"] if start <= parse_date(p["created_at"]) <= now()]
    order_spend = sum(float(p["amount"]) for p in buys if p["type"] == "order")
    trades = [x for x in data["trades"] if start <= parse_date(x["at"]) <= now()]
    realized = sum(float(x.get("realized_pnl", 0)) for x in trades if x["side"] == "sell")
    float_pnl = 0.0
    for pos in data["positions"].values():
        if float(pos["shares"]) <= 0:
            continue
        p0, p1, pnl = week_pnl_of(pos, start)
        float_pnl += pnl
    invest = realized + float_pnl
    return {"start": start, "end": now(), "order_spend": order_spend,
            "asset_spend": sum(float(p["amount"]) for p in buys if p["type"] == "asset"),
            "realized": realized, "float_pnl": float_pnl, "invest": invest,
            "change": invest - order_spend}


def export_xlsx(data, path):
    """把账本导出成本地 Excel（多 sheet），失败不影响记账。"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        print("⚠ 没有 openpyxl，先跳过 Excel 导出")
        return None

    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="1F2937")
    head_font = Font(color="FFFFFF", bold=True, size=11)
    title_font = Font(bold=True, size=14)

    def sheet(name, first=False):
        ws = wb.active if first else wb.create_sheet()
        ws.title = name
        return ws

    def write_table(ws, headers, rows, widths, money_cols=(), start_row=1):
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=start_row, column=i, value=h)
            c.fill, c.font = head_fill, head_font
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[c.column_letter].width = widths[i - 1]
        for r, row in enumerate(rows, start_row + 1):
            for i, v in enumerate(row, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in money_cols:
                    c.number_format = '#,##0.00'
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)

    w = week_summary(data)
    t = totals(data)
    p = data.get("profile", {})

    ws = sheet("总览", first=True)
    ws["A1"] = "富豪账本 · 总览"
    ws["A1"].font = title_font
    rows = [
        ("称呼", p.get("address") or "老板"),
        ("更新时间", now().strftime("%Y-%m-%d %H:%M")),
        ("身家", t["net_worth"]),
        ("现金", t["cash"]),
        ("持仓市值", t["stock"]),
        ("名下资产", t["assets"]),
        ("信用卡待还", t.get("debt", 0)),
        ("", ""),
        ("本周区间", "%s ~ %s" % (w["start"].strftime("%m-%d"), w["end"].strftime("%m-%d"))),
        ("本周消费", w["order_spend"] + w["asset_spend"]),
        ("其中购物", w["order_spend"]),
        ("本周投资盈亏", w["invest"]),
        ("其中已实现", w["realized"]),
        ("其中浮动", w["float_pnl"]),
        ("本周身家变化", w["change"]),
        ("", ""),
        ("开局本金", p.get("start_cash", 0)),
        ("账本文件", os.path.abspath(getattr(args_holder, "ledger", "") or default_ledger_path())),
    ]
    for i, (k, v) in enumerate(rows, 3):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True)
        c = ws.cell(row=i, column=2, value=v)
        if isinstance(v, (int, float)):
            c.number_format = '#,##0.00'
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 26

    # 消费记录
    ws = sheet("消费记录")
    rows = []
    for x in data["purchases"]:
        rows.append([x["created_at"][:16], "资产" if x["type"] == "asset" else "购物",
                     x["name"], x.get("category", ""), float(x["unit_price"]), float(x["qty"]),
                     float(x["amount"]), x.get("site", ""), x.get("card", ""),
                     x.get("url", ""), "已卖" if x.get("status") == "sold" else "在手"])
    write_table(ws, ["时间", "类别", "名称", "分类", "单价", "数量", "金额", "平台", "支付卡", "链接", "状态"],
                rows, [18, 8, 34, 12, 14, 8, 16, 14, 30, 46, 8], money_cols=(5, 7))

    # 持仓
    ws = sheet("持仓")
    rows = []
    for key, pos in data["positions"].items():
        shares = float(pos["shares"])
        if shares <= 0:
            continue
        p0, p1, wpnl = week_pnl_of(pos, w["start"])
        mv = shares * p1
        cost = shares * float(pos["avg_cost"])
        rows.append([key, pos.get("name", ""), pos.get("market", ""), shares,
                     float(pos["avg_cost"]), p1, mv, mv - cost, (mv - cost) / cost * 100 if cost else 0,
                     wpnl, pos.get("last_price_at", "")[:16]])
    write_table(ws, ["代码", "名称", "市场", "股数", "成本价", "现价", "市值", "浮盈亏", "浮盈亏%", "本周盈亏", "行情时间"],
                rows, [10, 16, 8, 12, 12, 12, 16, 14, 10, 14, 18], money_cols=(5, 6, 7, 8, 10))

    # 卡包
    ws = sheet("卡包")
    rows = []
    for c in data.get("cards", []):
        if c.get("type") == "credit":
            rows.append([c.get("issuer", ""), c["name"], c.get("last4", ""), "信用卡",
                         float(c.get("limit", 0)), float(c.get("outstanding", 0)),
                         float(c.get("limit", 0)) - float(c.get("outstanding", 0)),
                         c.get("billing_day", 0), c.get("due_day", 0), c.get("rewards", "")])
        else:
            rows.append([c.get("issuer", ""), c["name"], c.get("last4", ""), "借记卡",
                         0, 0, 0, 0, 0, c.get("rewards", "")])
    write_table(ws, ["银行", "卡名", "尾号", "类型", "额度", "待还", "可用", "账单日", "还款日", "权益"],
                rows, [14, 26, 8, 8, 16, 16, 16, 8, 8, 34], money_cols=(5, 6, 7))

    # 关注清单
    ws = sheet("关注清单")
    rows = [[x["category"], x["keyword"], x.get("region", ""),
             x.get("target_price") or "", x.get("last_result", "") or "待核查",
             x.get("last_check", "")[:16] or "从未核查", x.get("status", "open")]
            for x in data.get("watchlist", [])]
    write_table(ws, ["类别", "关键词", "地区", "目标价", "最近结论", "上次核查", "状态"],
                rows, [12, 34, 12, 14, 56, 18, 8], money_cols=(4,))

    # 流水
    ws = sheet("流水")
    rows = [[j["at"][:16], j["reason"], float(j["delta"]), float(j["balance"])]
            for j in data.get("journal", [])]
    write_table(ws, ["时间", "摘要", "金额变动", "余额"], rows,
                [18, 40, 18, 18], money_cols=(3, 4))

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path


def ledger_path(args):
    return os.path.abspath(args.ledger or os.environ.get("TYCOON_LEDGER") or default_ledger_path())


CLOUD_STATE_SHEET = "数据"


def sheet_url(args):
    """云端账本地址：--sheet / TYCOON_SHEET / config.json 里的 sheet_url。"""
    u = getattr(args, "sheet", None) or os.environ.get("TYCOON_SHEET")
    if u:
        return u
    try:
        with open(config_path(), "r", encoding="utf-8") as f:
            return json.load(f).get("sheet_url") or None
    except Exception:
        return None


def _cli(argv, stdin_text=None):
    """调用 lark-cli sheets，返回 data。"""
    import subprocess
    p = subprocess.run(["lark-cli", "sheets"] + argv, capture_output=True,
                       text=True, input=stdin_text)
    if p.returncode != 0:
        raise RuntimeError((p.stderr or p.stdout).strip()[:400])
    out = p.stdout.strip()
    at = out.find("{")
    if at < 0:
        raise RuntimeError(out[:400])
    r = json.loads(out[at:])
    if not r.get("ok"):
        raise RuntimeError(str(r.get("error"))[:400])
    return r.get("data", {})


def _csv_cell(text):
    import csv as _csv
    import io as _io
    buf = _io.StringIO()
    _csv.writer(buf, quoting=_csv.QUOTE_MINIMAL).writerow([text])
    return buf.getvalue()


def cloud_load(args):
    url = sheet_url(args)
    try:
        d = _cli(["+csv-get", "--url", url, "--sheet-name", CLOUD_STATE_SHEET,
                  "--range", "A1"])
    except RuntimeError:
        return None
    import csv as _csv
    import io as _io
    text = d.get("csv") or ""
    rows = list(_csv.reader(_io.StringIO(text)))
    if not rows or not rows[0]:
        return None
    raw = rows[0][0]
    return json.loads(raw) if raw.strip() else None


def cloud_ensure_state_sheet(url):
    try:
        _cli(["+csv-get", "--url", url, "--sheet-name", CLOUD_STATE_SHEET, "--range", "A1"])
    except RuntimeError:
        _cli(["+sheet-create", "--url", url, "--title", CLOUD_STATE_SHEET])


def view_tables(data):
    """云端可读页：名称 -> (表头, 行)。"""
    t = totals(data)
    w = week_summary(data)
    p = data.get("profile", {})
    out = {}

    rows = [["称呼", p.get("address") or "老板"],
            ["更新时间", now().strftime("%Y-%m-%d %H:%M")],
            ["身家", round(t["net_worth"], 2)], ["现金", round(t["cash"], 2)],
            ["持仓市值", round(t["stock"], 2)], ["名下资产", round(t["assets"], 2)],
            ["信用卡待还", round(t.get("debt", 0), 2)],
            ["", ""],
            ["本周区间", "%s ~ %s" % (w["start"].strftime("%m-%d"), w["end"].strftime("%m-%d"))],
            ["本周消费", round(w["order_spend"] + w["asset_spend"], 2)],
            ["本周投资盈亏", round(w["invest"], 2)],
            ["本周身家变化", round(w["change"], 2)]]
    out["总览"] = (["项目", "数值"], rows)

    rows = [[x["created_at"][:16], "资产" if x["type"] == "asset" else "购物",
             x["name"], x.get("category", ""), round(float(x["unit_price"]), 2),
             float(x["qty"]), round(float(x["amount"]), 2), x.get("site", ""),
             x.get("card", ""), x.get("url", ""),
             "已卖" if x.get("status") == "sold" else "在手"] for x in data["purchases"]]
    out["消费记录"] = (["时间", "类别", "名称", "分类", "单价", "数量", "金额",
                     "平台", "支付卡", "链接", "状态"], rows)

    rows = []
    for key, pos in data["positions"].items():
        shares = float(pos["shares"])
        if shares <= 0:
            continue
        p0, p1, wpnl = week_pnl_of(pos, w["start"])
        mv = shares * p1
        cost = shares * float(pos["avg_cost"])
        rows.append([key, pos.get("name", ""), pos.get("market", ""), shares,
                     round(float(pos["avg_cost"]), 4), round(p1, 4), round(mv, 2),
                     round(mv - cost, 2), round((mv - cost) / cost * 100 if cost else 0, 2),
                     round(wpnl, 2), pos.get("last_price_at", "")[:16]])
    out["持仓"] = (["代码", "名称", "市场", "股数", "成本价", "现价", "市值",
                 "浮盈亏", "浮盈亏%", "本周盈亏", "行情时间"], rows)

    rows = []
    for c in data.get("cards", []):
        credit = c.get("type") == "credit"
        rows.append([c.get("issuer", ""), c["name"], c.get("last4", ""),
                     "信用卡" if credit else "借记卡",
                     round(float(c.get("limit", 0)), 2) if credit else 0,
                     round(float(c.get("outstanding", 0)), 2) if credit else 0,
                     round(float(c.get("limit", 0)) - float(c.get("outstanding", 0)), 2) if credit else 0,
                     c.get("billing_day", 0), c.get("due_day", 0), c.get("rewards", "")])
    out["卡包"] = (["银行", "卡名", "尾号", "类型", "额度", "待还", "可用",
                 "账单日", "还款日", "权益"], rows)

    rows = [[x["category"], x["keyword"], x.get("region", ""), x.get("last_result") or "待核查",
             x.get("last_check", "")[:16] or "从未核查", x.get("status", "open")]
            for x in data.get("watchlist", [])]
    out["关注清单"] = (["类别", "关键词", "地区", "最近结论", "上次核查", "状态"], rows)

    rows = [[j["at"][:16], j["reason"], round(float(j["delta"]), 2), round(float(j["balance"]), 2)]
            for j in data.get("journal", [])]
    out["流水"] = (["时间", "摘要", "金额变动", "余额"], rows)
    return out


def _csv_text(headers, rows):
    import csv as _csv
    import io as _io
    buf = _io.StringIO()
    w = _csv.writer(buf, lineterminator="\n")
    w.writerow(headers)
    for r in rows:
        w.writerow(["" if v is None else v for v in r])
    return buf.getvalue()


def cloud_save(args, data):
    url = sheet_url(args)
    data["updated_at"] = now_iso()
    cloud_ensure_state_sheet(url)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    _cli(["+csv-put", "--url", url, "--sheet-name", CLOUD_STATE_SHEET,
          "--start-cell", "A1", "--csv", "-"], stdin_text=_csv_cell(payload))
    for name, (headers, rows) in view_tables(data).items():
        try:
            _cli(["+csv-put", "--url", url, "--sheet-name", name,
                  "--start-cell", "A1", "--csv", "-"], stdin_text=_csv_text(headers, rows))
        except RuntimeError as e:
            print("⚠「%s」页刷新失败：%s" % (name, e))
    return url


def load(args, must_exist=True):
    if sheet_url(args):
        data = cloud_load(args)
        if data is None:
            if must_exist:
                die("云端账本里还没有数据页，先跑 setup 初始化。地址：%s" % sheet_url(args))
            return None
        return ensure(data)
    path = ledger_path(args)
    if not os.path.exists(path):
        if must_exist:
            die("账本不存在，先运行 init 开局。路径：%s" % path)
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save(args, data):
    if sheet_url(args):
        return cloud_save(args, data)
    path = ledger_path(args)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    data["updated_at"] = now_iso()
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)
    try:
        global args_holder
        args_holder = args
        export_xlsx(data, xlsx_path(args))
    except Exception as e:  # 导出失败不影响记账
        print("⚠ Excel 导出失败：%s" % e)
    return path


def die(msg, code=2):
    print("错误：%s" % msg, file=sys.stderr)
    sys.exit(code)


def money(x, currency="CNY"):
    x = float(x)
    sign = "-" if x < 0 else ""
    x = abs(x)
    sym = {"CNY": "¥", "USD": "$", "HKD": "HK$", "EUR": "€"}.get(currency, "")
    if x >= 1e8:
        return "%s%s%.4f亿" % (sign, sym, x / 1e8)
    if x >= 1e4:
        return "%s%s%.2f万" % (sign, sym, x / 1e4)
    return "%s%s%.2f" % (sign, sym, x)


def speak_money(x):
    """口语化金额：3 亿、994 万、3299 元，不堆小数。"""
    x = float(x)
    sign = "-" if x < 0 else ""
    x = abs(x)
    if x >= 1e8:
        v = x / 1e8
        return "%s%s亿" % (sign, ("%.2f" % v).rstrip("0").rstrip("."))
    if x >= 1e4:
        v = x / 1e4
        return "%s%s万" % (sign, ("%.1f" % v).rstrip("0").rstrip("."))
    return "%s%s元" % (sign, "{:,.0f}".format(x))


def period_start(period):
    d = now()
    if period == "day":
        return d.replace(hour=0, minute=0, second=0, microsecond=0)
    if period == "week":
        monday = d - timedelta(days=d.weekday())
        return monday.replace(hour=0, minute=0, second=0, microsecond=0)
    if period == "month":
        return d.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    if period == "year":
        return d.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
    if period == "all":
        return datetime(1970, 1, 1, tzinfo=TZ)
    die("period 只支持 day/week/month/year/all")


def new_id(prefix):
    return "%s-%s" % (prefix, uuid.uuid4().hex[:8])


# ---------- 估值 ----------

def position_value(pos):
    return float(pos["shares"]) * float(pos.get("last_price") or pos["avg_cost"])


def price_at(pos, at):
    """取 at 时刻（或之前最近一次）的价格，没有则退回成本价。"""
    hist = pos.get("price_history") or []
    chosen = None
    for h in hist:
        t = parse_date(h["at"])
        if t <= at:
            chosen = float(h["price"])
    return chosen if chosen is not None else float(pos["avg_cost"])


def totals(data):
    cash = float(data["profile"]["cash"])
    stock = sum(position_value(p) for p in data["positions"].values())
    owned = [a for a in data["purchases"] if a.get("status", "owned") == "owned"
             and a.get("type") == "asset"]
    asset_val = sum(float(a.get("current_value", a["amount"])) for a in owned)
    debt = sum(float(c.get("outstanding", 0)) for c in data.get("cards", [])
               if c.get("type") == "credit")
    return {
        "cash": cash,
        "stock": stock,
        "assets": asset_val,
        "debt": round(debt, 2),
        "net_worth": round(cash + stock + asset_val - debt, 2),
    }


def snapshot(data, at=None):
    t = totals(data)
    snap = {"date": today(), "at": now_iso(), "cash": t["cash"],
            "stock": t["stock"], "assets": t["assets"], "net_worth": t["net_worth"]}
    hist = data.setdefault("snapshots", [])
    if at:
        snap["at"] = at
    hist.append(snap)
    data["snapshots"] = hist[-400:]
    return snap


def week_key(d=None):
    d = d or now()
    y, w, _ = d.isocalendar()
    return "%d-W%02d" % (y, w)


def touch_week(data):
    """本周第一次接触某持仓时，把当时的价格记为本周基准价。"""
    key = week_key()
    for pos in data.get("positions", {}).values():
        if pos.get("week") != key:
            pos["week"] = key
            pos["week_open_price"] = float(pos.get("last_price") or pos["avg_cost"])
    data["week"] = key
    return data


def begin(args):
    data = touch_week(ensure(load(args)))
    return data


def ensure(data):
    data.setdefault("positions", {})
    data.setdefault("purchases", [])
    data.setdefault("trades", [])
    data.setdefault("journal", [])
    data.setdefault("snapshots", [])
    data.setdefault("watchlist", [])
    data.setdefault("cards", [])
    return data


def cash_move(data, delta, reason, ref=None):
    data["profile"]["cash"] = round(float(data["profile"]["cash"]) + float(delta), 2)
    data["journal"].append({"at": now_iso(), "delta": round(float(delta), 2),
                            "reason": reason, "ref": ref,
                            "balance": data["profile"]["cash"]})


# ---------- 命令 ----------

DEFAULT_CARDS = [
    {"id": "card-vip", "name": "私人银行钻石借记卡", "issuer": "招商银行", "type": "debit",
     "last4": "8821", "limit": 0, "rewards": "消费 1 倍积分", "billing_day": 0, "due_day": 0,
     "tier": "黑金 · 借记", "note": "大额付款走这张，直接扣余额"},
    {"id": "card-centurion", "name": "百夫长黑金卡", "issuer": "美国运通", "type": "credit",
     "last4": "1005", "limit": 20000000, "rewards": "机场贵宾厅 / 酒店升舱 / 2 倍积分",
     "billing_day": 3, "due_day": 23, "tier": "黑金 · 信用卡", "note": "买游艇、拍艺术品刷这张最有面子"},
    {"id": "card-infiniti", "name": "长城国际 Visa 无限卡", "issuer": "中国银行", "type": "credit",
     "last4": "6603", "limit": 5000000, "rewards": "海外消费 3 倍积分 / 免货币转换费",
     "billing_day": 5, "due_day": 25, "tier": "无限卡 · 信用卡", "note": "境外房产、海外购物首选"},
    {"id": "card-hsbc", "name": "尚玉 Visa Infinite", "issuer": "汇丰银行", "type": "credit",
     "last4": "7788", "limit": 8000000, "rewards": "全球礼宾 / 高端医疗 / 1.5 倍积分",
     "billing_day": 8, "due_day": 28, "tier": "高端 · 信用卡", "note": "跨境消费和旅行支出"},
    {"id": "card-citic", "name": "白金信用卡", "issuer": "中信银行", "type": "credit",
     "last4": "4492", "limit": 1000000, "rewards": "餐饮 5 倍积分 / 每季送酒店权益",
     "billing_day": 12, "due_day": 30, "tier": "白金 · 信用卡", "note": "日常吃喝和买买买"},
]


def cmd_init(args):
    path = ledger_path(args)
    if os.path.exists(path) and not args.force:
        die("账本已存在：%s（要重开请加 --force）" % path)
    data = {
        "version": 1,
        "profile": {"owner": args.owner or "", "currency": "CNY",
                    "start_cash": float(args.cash), "cash": float(args.cash),
                    "created_at": now_iso()},
        "positions": {}, "purchases": [], "trades": [], "journal": [], "snapshots": [],
        "watchlist": [],
        "cards": [] if getattr(args, "no_cards", False) else [dict(c, outstanding=0.0) for c in DEFAULT_CARDS],
    }
    snapshot(data)
    save(args, data)
    write_memo(data, args)
    print("开局完成：身家 %s，账本 %s，台账 %s" % (money(args.cash), path, xlsx_path(args)))


def find_card(data, ref):
    """按 id 或卡号后四位找卡。"""
    if not ref:
        return None
    ref = ref.strip().lower()
    for c in data.get("cards", []):
        cid = c["id"].lower()
        if ref == str(c.get("last4", "")) or ref == cid or ref in cid or ref in c["name"].lower() \
                or ref in str(c.get("issuer", "")).lower():
            return c
    return None


def card_label(c):
    return "%s %s（尾号 %s）" % (c.get("issuer", ""), c["name"], c.get("last4", ""))


def cmd_buy(args):
    data = begin(args)
    amount = round(float(args.price) * float(args.qty), 2)

    card = find_card(data, args.card) if args.card else None
    if args.card and not card:
        die("卡包里没有这张卡：%s。先 `card list` 看看有哪些" % args.card)

    if card and card.get("type") == "credit":
        available = float(card.get("limit", 0)) - float(card.get("outstanding", 0))
        if amount > available:
            die("这张卡额度不够：可用 %s，这笔要 %s（可换卡或先还款）" % (
                money(available), money(amount)))
        card["outstanding"] = round(float(card.get("outstanding", 0)) + amount, 2)
        pay_desc = "刷卡（本期账单）"
        balance_desc = "本期待还 %s｜剩余额度 %s" % (
            money(card["outstanding"]),
            money(float(card["limit"]) - card["outstanding"]))
    else:
        if amount > float(data["profile"]["cash"]):
            die("现金不足：需要 %s，账户只有 %s" % (money(amount), money(data["profile"]["cash"])))
        cash_move(data, -amount, "虚拟下单：%s" % args.name, None)
        pay_desc = "借记卡/余额支付" if card else "余额支付"
        balance_desc = "余额 %s" % money(data["profile"]["cash"])

    item = {
        "id": new_id("buy"), "type": args.type, "name": args.name,
        "category": args.category or "", "qty": float(args.qty),
        "unit_price": float(args.price), "amount": amount,
        "currency": "CNY", "site": args.site or "", "url": args.url or "",
        "note": args.note or "", "status": "owned", "created_at": now_iso(),
        "card": card_label(card) if card else "余额支付",
    }
    data["purchases"].append(item)
    snapshot(data)
    save(args, data)
    print("已记入账本：%s ×%g = %s｜%s%s｜%s" % (
        args.name, float(args.qty), money(amount), pay_desc,
        (" · " + card_label(card)) if card else "",
        balance_desc))
    print("（虚拟记账，未真实扣款）")


def cmd_card(args):
    data = begin(args)
    cards = data["cards"]

    if args.action == "seed":
        existing = {c["id"] for c in cards}
        added = 0
        for c in DEFAULT_CARDS:
            if c["id"] not in existing:
                cards.append(dict(c, outstanding=0.0))
                added += 1
        save(args, data)
        print("卡包已补齐，新增 %d 张，共 %d 张" % (added, len(cards)))
        return

    if args.action == "add":
        if not args.name or not args.last4:
            die("加卡需要 --name 和 --last4（四位为虚构，别填真实卡号）")
        card = {"id": new_id("card"), "name": args.name, "issuer": args.issuer or "",
                "type": args.type or "credit", "last4": args.last4,
                "limit": float(args.limit or 0), "rewards": args.rewards or "",
                "billing_day": int(args.billing_day or 0), "due_day": int(args.due_day or 0),
                "tier": args.tier or "", "note": args.note or "", "outstanding": 0.0}
        cards.append(card)
        save(args, data)
        print("已加入卡包：%s" % card_label(card))
        return

    if args.action == "rm":
        card = find_card(data, args.id)
        if not card:
            die("找不到这张卡：%s" % args.id)
        cards.remove(card)
        save(args, data)
        print("已移除：%s" % card_label(card))
        return

    if args.action == "pay":
        card = find_card(data, args.id)
        if not card:
            die("找不到这张卡：%s" % args.id)
        if card.get("type") != "credit":
            die("只有信用卡需要还款")
        amount = float(args.amount or card.get("outstanding", 0))
        amount = min(amount, float(card.get("outstanding", 0)))
        if amount > float(data["profile"]["cash"]):
            die("现金不够还这笔账：需要 %s，只有 %s" % (money(amount), money(data["profile"]["cash"])))
        card["outstanding"] = round(float(card["outstanding"]) - amount, 2)
        cash_move(data, -amount, "还信用卡：%s" % card["name"], card["id"])
        snapshot(data)
        save(args, data)
        print("已还款 %s → %s｜剩余待还 %s｜余额 %s" % (
            card_label(card), money(amount), money(card["outstanding"]),
            money(data["profile"]["cash"])))
        return

    # list
    if not cards:
        print("卡包是空的，用 `card --action seed` 生成一套默认卡。")
        return
    print("卡包（%d 张）：" % len(cards))
    for c in cards:
        if c.get("type") == "credit":
            avail = float(c.get("limit", 0)) - float(c.get("outstanding", 0))
            print("  [%s] %s｜%s｜额度 %s｜待还 %s｜可用 %s｜账单日 %s 日 / 还款日 %s 日｜%s" % (
                c["id"], card_label(c), c.get("tier", ""), money(c.get("limit", 0)),
                money(c.get("outstanding", 0)), money(avail), c.get("billing_day", "-"),
                c.get("due_day", "-"), c.get("rewards", "")))
        else:
            print("  [%s] %s｜%s｜直连账户余额 %s｜%s" % (
                c["id"], card_label(c), c.get("tier", ""), money(data["profile"]["cash"]),
                c.get("rewards", "")))
    if args.json:
        print("\n" + json.dumps(cards, ensure_ascii=False, indent=2))


def cmd_sell_asset(args):
    data = begin(args)
    item = next((p for p in data["purchases"] if p["id"] == args.id), None)
    if not item:
        die("找不到这笔资产：%s" % args.id)
    if item.get("status") == "sold":
        die("这笔资产已经卖掉了")
    proceeds = round(float(args.price), 2)
    gain = round(proceeds - float(item["amount"]), 2)
    item["status"] = "sold"
    item["sold_price"] = proceeds
    item["sold_at"] = now_iso()
    item["realized_pnl"] = gain
    cash_move(data, proceeds, "卖出资产：%s" % item["name"], item["id"])
    snapshot(data)
    save(args, data)
    print("已卖出：%s → %s，实现盈亏 %s｜余额 %s" % (
        item["name"], money(proceeds), money(gain), money(data["profile"]["cash"])))


def cmd_revalue(args):
    data = begin(args)
    item = next((p for p in data["purchases"] if p["id"] == args.id), None)
    if not item:
        die("找不到这笔资产：%s" % args.id)
    item["current_value"] = float(args.value)
    item["revalued_at"] = now_iso()
    snapshot(data)
    save(args, data)
    print("已重估：%s 现估值 %s" % (item["name"], money(args.value)))


def cmd_stock_buy(args):
    data = begin(args)
    shares, price = float(args.shares), float(args.price)
    amount = round(shares * price, 2)
    fee = float(args.fee or 0)
    if amount + fee > float(data["profile"]["cash"]):
        die("现金不足：需要 %s，账户只有 %s" % (money(amount + fee), money(data["profile"]["cash"])))
    key = args.symbol.upper()
    pos = data["positions"].get(key)
    if pos:
        total_shares = float(pos["shares"]) + shares
        pos["avg_cost"] = round((float(pos["avg_cost"]) * float(pos["shares"]) + amount) / total_shares, 4)
        pos["shares"] = total_shares
        pos["last_price"] = price
        pos["last_price_at"] = now_iso()
        pos.setdefault("price_history", []).append({"at": now_iso(), "price": price})
        pos["price_history"] = pos["price_history"][-200:]
    else:
        pos = {"symbol": key, "name": args.name or key, "market": args.market or "",
               "shares": shares, "avg_cost": round(price, 4), "last_price": price,
               "last_price_at": now_iso(), "currency": args.currency or "CNY",
               "realized_pnl": 0.0, "opened_at": now_iso(),
               "price_history": [{"at": now_iso(), "price": price}]}
        data["positions"][key] = pos
    if args.note:
        pos["note"] = args.note
    data["trades"].append({"at": now_iso(), "symbol": key, "side": "buy",
                           "shares": shares, "price": price, "amount": amount, "fee": fee,
                           "note": args.note})
    cash_move(data, -(amount + fee), "买入 %s" % key, "stock")
    if (args.currency or "CNY").upper() != "CNY":
        print("⚠ 账本统一按人民币计价：已写入的 %.4f 将当作人民币价，请确认已按当日汇率换算（原币请写进名称/备注）。" % price)
    snapshot(data)
    save(args, data)
    print("已建仓/加仓：%s %g 股 @ %.4f，市值 %s｜余额 %s" % (
        pos["name"], shares, price, money(amount), money(data["profile"]["cash"])))


def cmd_stock_sell(args):
    data = begin(args)
    key = args.symbol.upper()
    pos = data["positions"].get(key)
    if not pos:
        die("没有持仓：%s" % key)
    shares = float(args.shares)
    if shares > float(pos["shares"]):
        die("持仓不足：只有 %g 股" % float(pos["shares"]))
    price = float(args.price)
    fee = float(args.fee or 0)
    proceeds = round(shares * price - fee, 2)
    cost = round(float(pos["avg_cost"]) * shares, 2)
    gain = round(proceeds - cost, 2)
    pos["shares"] = float(pos["shares"]) - shares
    pos["realized_pnl"] = round(float(pos.get("realized_pnl", 0)) + gain, 2)
    pos["last_price"] = price
    pos["last_price_at"] = now_iso()
    pos.setdefault("price_history", []).append({"at": now_iso(), "price": price})
    data["trades"].append({"at": now_iso(), "symbol": key, "side": "sell", "shares": shares,
                           "price": price, "amount": proceeds, "fee": fee,
                           "realized_pnl": gain})
    if float(pos["shares"]) <= 0:
        pos["shares"] = 0.0
        pos["closed_at"] = now_iso()
    cash_move(data, proceeds, "卖出 %s" % key, "stock")
    snapshot(data)
    save(args, data)
    print("已卖出：%s %g 股 @ %.4f，实现盈亏 %s｜余额 %s" % (
        key, shares, price, money(gain), money(data["profile"]["cash"])))


def cmd_mark(args):
    data = begin(args)
    updated, missing = [], []
    for pair in args.quote:
        if "=" not in pair:
            die("行情格式应为 代码=价格，例如 AAPL=180.5")
        sym, px = pair.split("=", 1)
        key = sym.strip().upper()
        try:
            price = float(px)
        except ValueError:
            die("价格不是数字：%s" % px)
        pos = data["positions"].get(key)
        if not pos:
            missing.append(key)
            continue
        pos["last_price"] = price
        pos["last_price_at"] = now_iso()
        pos.setdefault("price_history", []).append({"at": now_iso(), "price": price})
        pos["price_history"] = pos["price_history"][-200:]
        updated.append((key, price))
    if updated:
        snapshot(data)
        save(args, data)
    for key, price in updated:
        pos = data["positions"][key]
        pnl = (price - float(pos["avg_cost"])) * float(pos["shares"])
        print("行情更新：%s %s → %.4f｜持仓浮盈 %s（市值 %s）" % (
            pos.get("name", key), key, price, money(pnl), money(position_value(pos))))
    if missing:
        print("未持仓、已忽略：%s" % ", ".join(missing))


def days_since(at_iso):
    try:
        return max(0, (now() - parse_date(at_iso)).days)
    except Exception:
        return 0


def cmd_status(args):
    bootstrap(args)
    data = begin(args)
    t = totals(data)
    out = {"totals": t, "cash": t["cash"], "positions": [], "assets": [], "recent": []}
    print("身家：%s｜现金：%s｜持仓市值：%s｜收藏资产：%s%s" % (
        money(t["net_worth"]), money(t["cash"]), money(t["stock"]), money(t["assets"]),
        ("｜信用卡待还：%s" % money(t.get("debt", 0))) if t.get("debt") else ""))
    if data["positions"]:
        print("\n持仓：")
        for key, pos in data["positions"].items():
            if float(pos["shares"]) <= 0:
                continue
            mv = position_value(pos)
            pnl = mv - float(pos["avg_cost"]) * float(pos["shares"])
            pnl_pct = (pnl / (float(pos["avg_cost"]) * float(pos["shares"])) * 100
                       if float(pos["avg_cost"]) * float(pos["shares"]) else 0)
            print("  %-6s %-10s %10g 股  成本 %.4f  现价 %.4f  市值 %s  浮盈 %s (%+.2f%%)" % (
                key, pos.get("name", ""), float(pos["shares"]), float(pos["avg_cost"]),
                float(pos.get("last_price") or pos["avg_cost"]), money(mv), money(pnl), pnl_pct))
            out["positions"].append({"symbol": key, "name": pos.get("name"),
                                     "shares": float(pos["shares"]),
                                     "avg_cost": float(pos["avg_cost"]),
                                     "last_price": float(pos.get("last_price") or pos["avg_cost"]),
                                     "market_value": mv, "unrealized_pnl": pnl})
    owned = [p for p in data["purchases"] if p.get("status", "owned") == "owned"]
    if owned:
        print("\n在手资产 / 虚拟订单：")
        for p in owned[-12:]:
            print("  [%s] %-20s %s  %s" % (p["id"], p["name"], money(p["amount"]),
                                          p.get("site") or p.get("category") or ""))
            out["assets"].append({"id": p["id"], "name": p["name"], "type": p["type"],
                                  "amount": p["amount"], "site": p.get("site"),
                                  "url": p.get("url"), "created_at": p["created_at"]})
    if args.json:
        print("\n" + json.dumps(out, ensure_ascii=False, indent=2))


def cmd_report(args):
    bootstrap(args)
    data = begin(args)
    start = period_start(args.period)
    end = now()

    buys = [p for p in data["purchases"] if start <= parse_date(p["created_at"]) <= end]
    spend = sum(float(p["amount"]) for p in buys)
    asset_spend = sum(float(p["amount"]) for p in buys if p["type"] == "asset")
    order_spend = sum(float(p["amount"]) for p in buys if p["type"] == "order")

    trades = [t for t in data["trades"] if start <= parse_date(t["at"]) <= end]
    realized = sum(float(t.get("realized_pnl", 0)) for t in trades if t["side"] == "sell")
    bought = sum(float(t["amount"]) for t in trades if t["side"] == "buy")

    float_pnl = 0.0
    positions_detail = []
    for key, pos in data["positions"].items():
        if float(pos["shares"]) <= 0:
            continue
        if pos.get("week") == week_key() and pos.get("week_open_price") is not None:
            p0 = float(pos["week_open_price"])
        else:
            p0 = price_at(pos, start)
        p1 = float(pos.get("last_price") or pos["avg_cost"])
        shares = float(pos["shares"])
        week_pnl = (p1 - p0) * shares
        float_pnl += week_pnl
        positions_detail.append({"symbol": key, "name": pos.get("name"), "shares": shares,
                                 "price_start": p0, "price_now": p1,
                                 "week_pnl": week_pnl,
                                 "total_pnl": (p1 - float(pos["avg_cost"])) * shares})

    # 身家变化恒等式：投资盈亏 - 购物消费 - 交易费用 + 资产重估增值
    # （买入资产是现金换资产，按成本计价，不影响身家）
    fees = sum(float(t.get("fee", 0)) for t in trades)
    reval_gain = 0.0
    for p in data["purchases"]:
        rv = parse_date(p.get("revalued_at") or p["created_at"])
        if start <= rv <= end and p.get("current_value") is not None:
            reval_gain += float(p["current_value"]) - float(p["amount"])
    net_now = totals(data)["net_worth"]
    net_change = realized + float_pnl + reval_gain - order_spend - fees

    label = {"day": "本日", "week": "本周", "month": "本月", "year": "今年", "all": "累计"}[args.period]
    print("【%s账单】%s ~ %s" % (label, start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")))
    print("消费：%s（资产 %s｜购物 %s，共 %d 笔）" % (
        money(spend), money(asset_spend), money(order_spend), len(buys)))
    print("股票已实现盈亏：%s（期间买入 %s）" % (money(realized), money(bought)))
    print("股票浮动盈亏：%s" % money(float_pnl))
    print("合计盈亏：%s" % money(realized + float_pnl))
    print("当前身家：%s（现金 %s｜持仓 %s｜资产 %s）" % (
        money(net_now), money(totals(data)["cash"]), money(totals(data)["stock"]),
        money(totals(data)["assets"])))
    if net_change is not None:
        print("%s身家变化：%s" % (label, money(net_change)))

    if args.detail:
        if buys:
            print("\n消费明细：")
            for p in buys:
                print("  - %s %s  %s  %s" % (p["created_at"][:16], p["name"],
                                            money(p["amount"]), p.get("site", "")))
        if positions_detail:
            print("\n持仓明细（周期内变动）：")
            for d in positions_detail:
                print("  - %-6s %-10s %10g 股  %.4f→%.4f  周期 %s  累计 %s" % (
                    d["symbol"], d["name"] or "", d["shares"], d["price_start"],
                    d["price_now"], money(d["week_pnl"]), money(d["total_pnl"])))

    if args.json:
        print("\n" + json.dumps({
            "period": args.period, "start": start.isoformat(), "end": end.isoformat(),
            "spend": spend, "asset_spend": asset_spend, "order_spend": order_spend,
            "realized_pnl": realized, "float_pnl": float_pnl,
            "total_pnl": realized + float_pnl, "net_worth": net_now,
            "net_change": net_change, "totals": totals(data),
            "positions": positions_detail, "purchases": buys,
        }, ensure_ascii=False, indent=2))


def week_pnl_of(pos, start):
    if pos.get("week") == week_key() and pos.get("week_open_price") is not None:
        p0 = float(pos["week_open_price"])
    else:
        p0 = price_at(pos, start)
    p1 = float(pos.get("last_price") or pos["avg_cost"])
    shares = float(pos["shares"])
    return p0, p1, (p1 - p0) * shares


def profile_of(data):
    p = data.setdefault("profile", {})
    p.setdefault("address", "老板")
    return p


def print_chat(data, t, d, weekdays, movers, invest, realized, float_pnl,
               order_spend, change, buys, max_stale=2):
    """把简报说成人话，而不是贴表。"""
    p = profile_of(data)
    addr = p.get("address") or "老板"
    weekday = weekdays[d.weekday()]
    closing = d.weekday() >= 5 or (d.hour < 9 or d.hour > 15)
    print("%s，%s好。%d月%d日，周%s%s。" % (
        addr, "早上" if d.hour < 11 else ("下午" if d.hour < 18 else "晚上"),
        d.month, d.day, weekday, "，盘都收了" if closing else ""))

    print("您这边家底 %s：现金 %s，持仓 %s，名下资产 %s%s。" % (
        speak_money(t["net_worth"]), speak_money(t["cash"]), speak_money(t["stock"]), speak_money(t["assets"]),
        ("，另有卡账 %s 待还" % speak_money(t["debt"])) if t.get("debt") else ""))

    if abs(invest) < 100 and order_spend < 100:
        print("这周基本没什么波动，投资这块盈亏是零。")
    else:
        print("这周投资%s %s，其中浮动 %s；另外花出去 %s。" % (
            "赚了" if invest >= 0 else "亏了", speak_money(abs(invest)),
            speak_money(float_pnl), speak_money(order_spend)))

    if movers:
        big = movers[0]
        if abs(big["pnl"]) >= 100:
            print("%s今天%s，%s，%s。" % (
                big["name"], "涨了" if big["pct"] > 0 else "跌了",
                ("%.2f%%" % abs(big["pct"])),
                ("账面多出 " + speak_money(big["pnl"])) if big["pnl"] > 0 else ("账面缩水 " + speak_money(-big["pnl"]))))

    if buys:
        top = max(buys, key=lambda x: float(x["amount"]))
        card = top.get("card", "")
        print("本周最大一笔是%s，%s%s。" % (
            top["name"], speak_money(top["amount"]),
            ("，刷的" + card) if card and "余额" not in card else ""))

    cards = data.get("cards", [])
    credit = [c for c in cards if c.get("type") == "credit"]
    if credit:
        due = [c for c in credit if float(c.get("outstanding", 0)) > 0]
        if due:
            c = max(due, key=lambda x: float(x["outstanding"]))
            print("卡这边，%s还挂着 %s，%s号到期，可用额度还剩 %s。" % (
                c["name"], speak_money(c["outstanding"]), c.get("due_day", "-"),
                speak_money(float(c.get("limit", 0)) - float(c["outstanding"]))))
        else:
            print("几张卡都是干净的，没有待还。")

    wl = [w for w in data.get("watchlist", []) if w.get("status", "open") == "open"]
    if wl:
        w = wl[0]
        print("您让我盯的%s，%s。" % (w["keyword"], w.get("last_result") or "我还没去查"))

    hints = []
    if t["net_worth"] and float(t["cash"]) / float(t["net_worth"]) >= 0.9:
        hints.append("现金占了身家的 %d%%，钱躺得有点多" % int(float(t["cash"]) / float(t["net_worth"]) * 100))
    stale = [x for x in data["positions"].values()
             if float(x["shares"]) > 0 and days_since(x.get("last_price_at", "")) >= max_stale]
    if stale:
        hints.append("%d 只持股的行情我还没更新" % len(stale))
    if hints:
        print("有件事得提一句：%s。" % "；".join(hints))

    # 收尾：只给一个建议 + 一个问题
    if t.get("debt"):
        print("您看是先还掉这 %s，还是我再去给您找点别的？" % speak_money(t["debt"]))
    elif hints:
        print("要不要我先给您挑两件合适的？")
    else:
        print("今天想让我办点什么？")


def cmd_export(args):
    data = begin(args)
    path = export_xlsx(data, xlsx_path(args))
    print("已导出 Excel：%s" % path if path else "导出失败")


def cmd_profile(args):
    data = begin(args)
    p = profile_of(data)
    if args.address:
        p["address"] = args.address
    if args.name:
        p["owner"] = args.name
    save(args, data)
    print("已记下：称呼「%s」%s" % (p["address"], ("，名字 " + p["owner"]) if p.get("owner") else ""))


def cmd_watch(args):
    data = begin(args)
    wl = data["watchlist"]

    if args.action == "add":
        if not args.keyword:
            die("关注清单需要 --keyword，例如 --category 游艇 --keyword '50米以上 Sunseeker'")
        item = {"id": new_id("w"), "category": args.category or "其他",
                "keyword": args.keyword, "region": args.region or "",
                "target_price": args.target_price, "note": args.note or "",
                "created_at": now_iso(), "last_check": None, "last_result": "",
                "status": "open"}
        wl.append(item)
        save(args, data)
        print("已加入关注清单 [%s] %s / %s" % (item["id"], item["category"], item["keyword"]))
        return

    if args.action == "rm":
        if not args.id:
            die("需要 --id")
        before = len(wl)
        data["watchlist"] = [w for w in wl if w["id"] != args.id]
        save(args, data)
        print("已移除 %d 条关注项" % (before - len(data["watchlist"])))
        return

    if args.action == "check":
        item = next((w for w in wl if w["id"] == args.id), None)
        if not item:
            die("找不到关注项：%s" % args.id)
        item["last_check"] = now_iso()
        item["last_result"] = args.result or "已核查"
        if args.status:
            item["status"] = args.status
        save(args, data)
        print("已记录核查：[%s] %s → %s" % (item["id"], item["keyword"], item["last_result"]))
        return

    # list
    open_items = [w for w in wl if w.get("status", "open") == "open"]
    if not open_items:
        print("关注清单是空的。")
        return
    print("关注清单（%d 条）：" % len(open_items))
    for i, w in enumerate(open_items, 1):
        age = days_since(w["last_check"]) if w.get("last_check") else None
        age_s = "从未核查" if age is None else ("今天核查" if age == 0 else "%d 天前核查" % age)
        flag = " ⚠超期" if (age is None or age > args.max_age) else ""
        print("  %d. [%s] %s%s｜%s%s%s" % (
            i, w["category"], w["keyword"], ("（%s）" % w["region"]) if w.get("region") else "",
            w.get("last_result") or "待核查", "｜" + age_s, flag))
    if args.json:
        print("\n" + json.dumps(open_items, ensure_ascii=False, indent=2))


def setup(args, quiet=False):
    """技能激活即执行：幂等准备账本、卡包、Excel 台账、config 与秘书备忘。"""
    os.makedirs(tycoon_home(), exist_ok=True)
    created = []

    data = load(args, must_exist=False)
    if data is None:
        cash = float(getattr(args, "cash", None) or 3e8)
        data = {
            "version": 1,
            "profile": {"owner": "", "address": getattr(args, "address", "") or "老板",
                        "currency": "CNY", "start_cash": cash, "cash": cash,
                        "created_at": now_iso()},
            "positions": {}, "purchases": [], "trades": [], "journal": [], "snapshots": [],
            "watchlist": [], "cards": [dict(c, outstanding=0.0) for c in DEFAULT_CARDS],
        }
        snapshot(data)
        created.append(("账本", sheet_url(args) or ledger_path(args)))

    # 卡包补齐（用户删过的卡不会被动恢复，只补默认 id 缺失的）
    existing = {c.get("id") for c in data["cards"]}
    missing = [dict(c, outstanding=0.0) for c in DEFAULT_CARDS if c["id"] not in existing]
    if missing:
        data["cards"].extend(missing)
        created.append(("卡包", "%d 张" % len(missing)))

    save(args, data)
    created.append(("Excel 台账", xlsx_path(args)))

    # config：缺失则新建；指定了云端表则更新 sheet_url，其余字段不动
    cfg_exists = os.path.exists(config_path())
    if cfg_exists and getattr(args, "sheet", None):
        with open(config_path(), "r", encoding="utf-8") as f:
            cfg = json.load(f)
        if cfg.get("sheet_url") != args.sheet:
            cfg["sheet_url"] = args.sheet
            with open(config_path(), "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            created.append(("云端账本", args.sheet))
    elif not cfg_exists:
        cfg = {"address": data["profile"].get("address", "老板"),
               "owner": data["profile"].get("owner", ""),
               "currency": "CNY", "timezone": "Asia/Shanghai",
               "excel": xlsx_path(args), "ledger": ledger_path(args),
               "report": "weekly", "created_at": now_iso()}
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        created.append(("适配配置", config_path()))

    write_memo(data, args)
    created.append(("秘书备忘", memo_path()))

    if not quiet:
        fresh = [c for c in created if c[0] in ("账本", "卡包", "适配配置")]
        who = data["profile"].get("address", "老板")
        if fresh:
            print("已为「%s」准备好：" % who)
            for k, v in created:
                print("  · %s：%s" % (k, v))
        else:
            print("「%s」的账本、卡包、配置都在，台账与备忘已刷新。" % who)
    return data


def write_memo(data, args):
    """秘书备忘：新会话第一眼就看这份，省得重新认识用户。"""
    t = totals(data)
    p = data.get("profile", {})
    lines = ["# 秘书备忘", "",
             "- 称呼：%s" % p.get("address", "老板"),
             "- 身家：%s（现金 %s｜持仓 %s｜资产 %s｜卡账 %s）" % (
                 speak_money(t["net_worth"]), speak_money(t["cash"]),
                 speak_money(t["stock"]), speak_money(t["assets"]), speak_money(t.get("debt", 0))),
             "- 开局本金：%s" % speak_money(p.get("start_cash", 0)),
             "", "## 卡包"]
    for c in data.get("cards", []):
        if c.get("type") == "credit":
            lines.append("- %s %s（尾号 %s）｜额度 %s｜待还 %s｜还款日 %s 日" % (
                c.get("issuer", ""), c["name"], c.get("last4", ""),
                speak_money(c.get("limit", 0)), speak_money(c.get("outstanding", 0)),
                c.get("due_day", "-")))
        else:
            lines.append("- %s %s（尾号 %s）｜借记卡" % (
                c.get("issuer", ""), c["name"], c.get("last4", "")))
    owned = [x for x in data["purchases"] if x.get("status", "owned") == "owned"]
    if owned:
        lines += ["", "## 名下资产"]
        for x in owned:
            lines.append("- %s｜%s｜%s" % (x["name"], speak_money(x["amount"]),
                                           x.get("site") or x.get("category") or ""))
    wl = [w for w in data.get("watchlist", []) if w.get("status", "open") == "open"]
    if wl:
        lines += ["", "## 关注清单"]
        for w in wl:
            lines.append("- [%s] %s｜%s" % (w["category"], w["keyword"],
                                            w.get("last_result") or "待核查"))
    lines += ["", "## 规矩",
              "- 货源与行情必须真实，钱和交易必须虚构，不产生真实支付。",
              "- 下单前先问用哪张卡；信用卡进账单，借记卡扣余额。",
              "- 每次记账自动刷新 Excel；随时可问本周花了多少、赚了多少。",
              "- 不索取、不记录真实卡号、CVV、验证码。", ""]
    os.makedirs(os.path.dirname(memo_path()), exist_ok=True)
    with open(memo_path(), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def bootstrap(args):
    """兼容旧调用：账本不存在时按默认身家准备全套文件。"""
    if os.path.exists(ledger_path(args)):
        return False
    setup(args)
    print("（首次使用，已按默认身家开局，卡包与台账一并备好。）")
    return True


def cmd_newsheet(args):
    """在飞书里新建一张云端账本表，并绑定到本用户。"""
    import tempfile
    cash = float(getattr(args, "cash", None) or 3e8)
    data = {
        "version": 1,
        "profile": {"owner": "", "address": "老板", "currency": "CNY",
                    "start_cash": cash, "cash": cash, "created_at": now_iso()},
        "positions": {}, "purchases": [], "trades": [], "journal": [], "snapshots": [],
        "watchlist": [], "cards": [dict(c, outstanding=0.0) for c in DEFAULT_CARDS],
    }
    snapshot(data)
    seed = os.path.join(tempfile.mkdtemp(), "seed.xlsx")
    export_xlsx(data, seed)
    d = _cli(["+workbook-import", "--file", seed, "--name", args.name or "富豪账本"])
    url = d.get("url") or d.get("data", {}).get("url")
    if not url:
        die("建表失败：%s" % d)

    cfg = {}
    if os.path.exists(config_path()):
        with open(config_path(), "r", encoding="utf-8") as f:
            cfg = json.load(f)
    cfg.update({"address": "老板", "currency": "CNY", "timezone": "Asia/Shanghai",
                "sheet_url": url, "ledger": ledger_path(args),
                "report": "weekly"})
    os.makedirs(tycoon_home(), exist_ok=True)
    with open(config_path(), "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

    args.sheet = url
    setup(args)
    print("云端账本已就绪：%s" % url)


def cmd_brief(args):
    """秘书简报：资产重大变化 + 余额 + 能做什么 + 关注清单。"""
    bootstrap(args)
    data = begin(args)
    start = period_start("week")
    t = totals(data)
    d = now()
    weekdays = "一二三四五六日"

    # 本周变化
    buys = [p for p in data["purchases"] if start <= parse_date(p["created_at"]) <= now()]
    order_spend = sum(float(p["amount"]) for p in buys if p["type"] == "order")
    trades = [x for x in data["trades"] if start <= parse_date(x["at"]) <= now()]
    realized = sum(float(x.get("realized_pnl", 0)) for x in trades if x["side"] == "sell")
    float_pnl, movers = 0.0, []
    for key, pos in data["positions"].items():
        if float(pos["shares"]) <= 0:
            continue
        p0, p1, pnl = week_pnl_of(pos, start)
        float_pnl += pnl
        pct = ((p1 - p0) / p0 * 100) if p0 else 0.0
        movers.append({"symbol": key, "name": pos.get("name") or key, "pct": pct,
                       "pnl": pnl, "price": p1, "shares": float(pos["shares"])})
    invest = realized + float_pnl
    change = invest - order_spend

    if getattr(args, "chat", False):
        print_chat(data, t, d, weekdays, movers, invest, realized, float_pnl,
                   order_spend, change, buys, getattr(args, "max_stale", 2))
        if args.json:
            print(json.dumps({"net_worth": t["net_worth"], "cash": t["cash"], "debt": t.get("debt", 0),
                              "week_change": change}, ensure_ascii=False, indent=2))
        return

    print("【富豪秘书简报】%d月%d日 周%s" % (d.month, d.day, weekdays[d.weekday()]))
    print("身家 %s｜现金 %s｜持仓 %s｜资产 %s%s" % (
        money(t["net_worth"]), money(t["cash"]), money(t["stock"]), money(t["assets"]),
        ("｜待还 %s" % money(t["debt"])) if t.get("debt") else ""))
    print("\n本周 投资盈亏 %s（已实现 %s｜浮动 %s）｜购物消费 %s｜身家变化 %s" % (
        money(invest), money(realized), money(float_pnl), money(order_spend), money(change)))

    cards = data.get("cards", [])
    if cards:
        print("\n【卡包】")
        for c in cards:
            if c.get("type") == "credit":
                avail = float(c.get("limit", 0)) - float(c.get("outstanding", 0))
                print("  %s %s（尾号 %s）｜额度 %s｜%s｜可用 %s｜还款日 %s 日" % (
                    c.get("issuer", ""), c["name"], c.get("last4", ""),
                    money(c.get("limit", 0)),
                    ("待还 " + money(c.get("outstanding", 0))) if float(c.get("outstanding", 0)) > 0 else "未出账",
                    money(avail), c.get("due_day", "-")))
            else:
                print("  %s %s（尾号 %s）｜余额 %s｜%s" % (
                    c.get("issuer", ""), c["name"], c.get("last4", ""),
                    money(t["cash"]), c.get("rewards", "")))

    print("\n【资产重大变化】")
    movers.sort(key=lambda m: -abs(m["pnl"]))
    shown = [m for m in movers if abs(m["pnl"]) >= 5000][:3] or movers[:1]
    if shown:
        for m in shown:
            arrow = "↑" if m["pnl"] > 0 else ("↓" if m["pnl"] < 0 else "→")
            print("  %s %s %.4f（%+.2f%%）｜%s 股｜本周 %s" % (
                arrow, m["name"], m["price"], m["pct"], "%g" % m["shares"], money(m["pnl"])))
    else:
        print("  持仓暂无变化（行情未更新时请先刷新）")

    reval = [p for p in data["purchases"]
             if p.get("revalued_at") and 0 <= days_since(p["revalued_at"]) <= 7]
    for p in reval[:3]:
        print("  🏛 %s 重估至 %s（%s）" % (
            p["name"], money(p.get("current_value", p["amount"])),
            "升值" if float(p.get("current_value", 0)) >= float(p["amount"]) else "缩水"))

    recent_buys = [p for p in buys if float(p["amount"]) >= 5000]
    if recent_buys:
        big = max(recent_buys, key=lambda p: float(p["amount"]))
        print("  💸 本周最大支出：%s %s（%s）" % (
            big["name"], money(big["amount"]),
            big.get("site") or ("购入资产" if big["type"] == "asset" else "购物下单")))

    print("\n【可以做什么】")
    print("  • 找货：游艇 / 豪宅 / 油气区块 / 私人飞机 / 岛屿 / 艺术品（实时抓真实挂牌）")
    print("  • 虚拟下单：淘宝、京东、亚马逊，只记账不付款")
    print("  • 虚拟炒股：A股 / 港股 / 美股 / 黄金，按真实行情浮动")
    print("  • 看账：本周账单、消费流水、资产分布")
    print("  • 加关注：把想盯的标的记进清单，我按点核查并报变化")

    print("\n【关注清单】")
    open_items = [w for w in data["watchlist"] if w.get("status", "open") == "open"]
    if open_items:
        for i, w in enumerate(open_items, 1):
            age = days_since(w["last_check"]) if w.get("last_check") else None
            age_s = "从未核查" if age is None else ("今天核查" if age == 0 else "%d 天前" % age)
            flag = " ⚠" if (age is None or age > args.max_age) else ""
            print("  %d. [%s] %s｜%s｜%s%s" % (
                i, w["category"], w["keyword"], w.get("last_result") or "待核查", age_s, flag))
    else:
        print("  还没有关注项。跟我说「盯一下 XX」，我就记进清单。")

    # 提醒
    stale = [p for p in data["positions"].values()
             if float(p["shares"]) > 0 and p.get("last_price_at") and days_since(p["last_price_at"]) >= args.max_stale]
    overdue = [w for w in open_items
               if not w.get("last_check") or days_since(w["last_check"]) > args.max_age]
    hints = []
    if stale:
        hints.append("%d 个持仓超过 %d 天没更新行情，报表会失真，建议先刷新" % (len(stale), args.max_stale))
    if overdue:
        never = [w for w in overdue if not w.get("last_check")]
        old = [w for w in overdue if w.get("last_check")]
        if never:
            hints.append("%d 条关注项还没核查过，我可以现在去抓一轮" % len(never))
        if old:
            hints.append("%d 条关注项超 %d 天未复检" % (len(old), args.max_age))
    cash_ratio = float(t["cash"]) / float(t["net_worth"]) if t["net_worth"] else 0
    if cash_ratio >= 0.9 and t["net_worth"] > 0:
        hints.append("现金占身家 %d%%，钱趴在账上，可以配置点资产或建仓" % int(cash_ratio * 100))
    if hints:
        print("\n【提醒】")
        for h in hints:
            print("  · " + h)

    if args.json:
        print("\n" + json.dumps({
            "net_worth": t["net_worth"], "cash": t["cash"], "stock": t["stock"],
            "assets": t["assets"], "week_change": change, "movers": movers,
            "watchlist": open_items, "stale_positions": [p["symbol"] for p in stale],
            "overdue_watch": [w["id"] for w in overdue],
        }, ensure_ascii=False, indent=2))


def cmd_list(args):
    data = begin(args)
    if not data["purchases"]:
        print("还没有任何消费记录。")
        return
    for p in data["purchases"][-args.limit:]:
        flag = "已卖出" if p.get("status") == "sold" else "在手"
        print("%s [%s] %-4s %-24s %12s %-10s %s %s" % (
            p["created_at"][:16], p["id"], flag, p["name"], money(p["amount"]),
            p.get("category", ""), p.get("site", ""), p.get("url", "")))


def cmd_reset(args):
    if not args.force:
        die("重开需要加 --force，会清空全部记录")
    cmd_init(args)


def main():
    ap = argparse.ArgumentParser(description="模拟富豪账本（纯虚拟，不产生真实交易）")
    ap.add_argument("--ledger", help="账本路径，默认 ~/.tycoon/ledger.json")
    ap.add_argument("--excel", help="Excel 导出路径，默认 ~/.tycoon/账本.xlsx")
    ap.add_argument("--sheet", help="云端账本：飞书表格链接或 token")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("newsheet", help="在飞书新建云端账本表并绑定")
    p.add_argument("--name", default="富豪账本")
    p.add_argument("--cash", type=float, default=3e8)
    p.set_defaults(func=cmd_newsheet)

    p = sub.add_parser("setup", help="激活技能：幂等准备账本/卡包/Excel/备忘")
    p.add_argument("--cash", type=float, default=3e8)
    p.add_argument("--address", default="", help="称呼，默认 老板")
    p.set_defaults(func=lambda a: setup(a))

    p = sub.add_parser("init", help="开局/建账本")
    p.add_argument("--cash", type=float, default=3e8)
    p.add_argument("--owner", default="")
    p.add_argument("--no-cards", action="store_true", help="不生成默认卡包")
    p.add_argument("--force", action="store_true")
    p.set_defaults(func=cmd_init)

    p = sub.add_parser("buy", help="虚拟下单/买入资产（只记账）")
    p.add_argument("--name", required=True)
    p.add_argument("--type", choices=["asset", "order"], default="order",
                   help="asset=游艇/房/油田等资产，order=购物网站下单")
    p.add_argument("--category", default="")
    p.add_argument("--price", type=float, required=True)
    p.add_argument("--qty", type=float, default=1)
    p.add_argument("--site", default="")
    p.add_argument("--url", default="")
    p.add_argument("--note", default="")
    p.add_argument("--card", default="", help="用哪张卡：卡 id 或后四位，不填则走余额")
    p.set_defaults(func=cmd_buy)

    p = sub.add_parser("sell-asset", help="卖出在手资产")
    p.add_argument("--id", required=True)
    p.add_argument("--price", type=float, required=True)
    p.set_defaults(func=cmd_sell_asset)

    p = sub.add_parser("revalue", help="重估资产现值")
    p.add_argument("--id", required=True)
    p.add_argument("--value", type=float, required=True)
    p.set_defaults(func=cmd_revalue)

    p = sub.add_parser("stock-buy", help="虚拟买入股票")
    p.add_argument("--symbol", required=True)
    p.add_argument("--name", default="")
    p.add_argument("--market", default="", help="A股/港股/美股")
    p.add_argument("--shares", type=float, required=True)
    p.add_argument("--price", type=float, required=True)
    p.add_argument("--fee", type=float, default=0)
    p.add_argument("--currency", default="CNY")
    p.add_argument("--note", default="", help="原币价格/汇率等说明")
    p.set_defaults(func=cmd_stock_buy)

    p = sub.add_parser("stock-sell", help="虚拟卖出股票")
    p.add_argument("--symbol", required=True)
    p.add_argument("--shares", type=float, required=True)
    p.add_argument("--price", type=float, required=True)
    p.add_argument("--fee", type=float, default=0)
    p.set_defaults(func=cmd_stock_sell)

    p = sub.add_parser("mark", help="写入真实行情快照，格式 代码=价格")
    p.add_argument("quote", nargs="+", help="例如 AAPL=180.5 600519=1500")
    p.set_defaults(func=cmd_mark)

    p = sub.add_parser("status", help="查看身家、持仓、在手资产")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_status)

    p = sub.add_parser("report", help="周期账单：消费/盈亏/身家变化")
    p.add_argument("--period", choices=["day", "week", "month", "year", "all"], default="week")
    p.add_argument("--detail", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_report)

    p = sub.add_parser("export", help="手动导出/刷新本地 Excel")
    p.set_defaults(func=cmd_export)

    p = sub.add_parser("profile", help="记下称呼与名字，秘书会一直这么叫")
    p.add_argument("--address", default="", help="称呼，如 老板/李总/先生")
    p.add_argument("--name", default="")
    p.set_defaults(func=cmd_profile)

    p = sub.add_parser("card", help="虚拟卡包：seed/list/add/rm/pay")
    p.add_argument("--action", choices=["list", "seed", "add", "rm", "pay"], default="list")
    p.add_argument("--id", default="")
    p.add_argument("--name", default="")
    p.add_argument("--issuer", default="")
    p.add_argument("--type", choices=["credit", "debit"], default="credit")
    p.add_argument("--last4", default="")
    p.add_argument("--limit", type=float, default=0)
    p.add_argument("--rewards", default="")
    p.add_argument("--tier", default="")
    p.add_argument("--billing-day", type=int, default=0)
    p.add_argument("--due-day", type=int, default=0)
    p.add_argument("--amount", type=float, default=0)
    p.add_argument("--note", default="")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_card)

    p = sub.add_parser("brief", help="秘书简报：异动+余额+能做什么+关注清单")
    p.add_argument("--max-age", type=int, default=7, help="关注项几天未核查算超期")
    p.add_argument("--max-stale", type=int, default=2, help="持仓几天未更新行情算过期")
    p.add_argument("--chat", action="store_true", help="秘书口吻播报（默认贴表）")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_brief)

    p = sub.add_parser("watch", help="关注清单：add/list/rm/check")
    p.add_argument("--action", choices=["add", "list", "rm", "check"], default="list")
    p.add_argument("--id", default="")
    p.add_argument("--category", default="", help="游艇/房产/油气区块/股票/商品/其他")
    p.add_argument("--keyword", default="")
    p.add_argument("--region", default="")
    p.add_argument("--target-price", type=float, default=None)
    p.add_argument("--note", default="")
    p.add_argument("--result", default="", help="本次核查结果，如 '新挂牌3条，最低 820 万美元'")
    p.add_argument("--status", default="", help="open/closed")
    p.add_argument("--max-age", type=int, default=7)
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_watch)

    p = sub.add_parser("list", help="消费流水")
    p.add_argument("--limit", type=int, default=20)
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("reset", help="清空重开")
    p.add_argument("--cash", type=float, default=3e8)
    p.add_argument("--owner", default="")
    p.add_argument("--force", action="store_true")
    p.set_defaults(func=cmd_reset)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
