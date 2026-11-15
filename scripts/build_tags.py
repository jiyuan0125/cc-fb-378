#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_tags.py — 支付事件流 -> 渠道日标签宽表

输入 CSV（每行一个支付事件，列: pay_id, channel, event, amount, ts）:
    event 取值: success / refund / reverse
      - success: 支付成功，同一 pay_id 以最早一条为准
      - reverse: 冲正，把同一 pay_id 的 success 整个抹掉（净额扣回原单全额）
      - refund : 成功之后往外退的钱，同一单可分多次退，累计上限 = 原单金额

输出 XLSX（两个工作表）:
  1) 渠道日标签  每渠道每个成功日一行，列:
       渠道 日期 成功笔数 成功金额 净额 退款笔数 退款金额 冲正笔数 成功率 退款率
  2) 异常单登记  一个异常支付单一行（含行级脏数据）

口径（粒度: 渠道 × 支付单 success 所在自然日）:
  成功笔数  当日 success 的支付单笔数（毛额，含后来被冲正的单）
  成功金额  上述支付单 success 金额合计
  冲正笔数  被 reverse 抹掉的支付单笔数（跨天冲正回溯到成功日）
  冲正金额  被抹掉的 success 原单金额（用于净额，内部口径不单独出列）
  退款笔数  success 之后、且单据未被冲正的有效 refund 事件数
  退款金额  有效 refund 金额合计；累计超过原单的部分只登记异常，不入宽表
  净额      财务口径“最后落到手里的钱” = 成功金额 - 退款金额 - 冲正金额
  成功率    (成功笔数 - 冲正笔数) / 成功笔数
  退款率    退款金额 / 成功金额

确定性: 纯标准库、无随机数、无 wall-clock、固定排序与 zip 时间戳，
同一份数据重跑产物字节一致；整表重建（先写临时文件再原子替换）。

用法:
  python3 build_tags.py transactions.csv -o channel_daily.xlsx
"""

import argparse
import csv
import os
import zipfile
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation

VALID_EVENTS = ("success", "refund", "reverse")
TS_FORMATS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%d %H:%M:%S.%f",
    "%Y-%m-%dT%H:%M:%S.%f",
    "%Y-%m-%d",
)
MONEY_Q = Decimal("0.01")

DAILY_HEADER = ["渠道", "日期", "成功笔数", "成功金额", "净额",
                "退款笔数", "退款金额", "冲正笔数", "成功率", "退款率"]
ANOMALY_HEADER = ["渠道", "支付单号", "异常类型", "异常说明",
                  "原单金额", "累计退款金额", "超额金额", "CSV行号"]

# 异常代码 -> (中文类型, 说明模板)
ANOMALY_TYPES = {
    "BAD_ROW":                ("行数据非法",          "字段解析失败: {detail}"),
    "CHANNEL_CONFLICT":       ("同单渠道不一致",      "同一支付单出现多个渠道: {detail}"),
    "DUP_SUCCESS":            ("重复success",         "同一支付单多条 success，仅保留最早一条"),
    "DUP_REVERSE":            ("重复reverse",         "同一支付单多条 reverse，仅保留最早一条"),
    "ORPHAN_REVERSE":         ("冲正无对应success",   "reverse 找不到同一支付单的 success"),
    "ORPHAN_REFUND":          ("退款无对应success",   "refund 找不到同一支付单的 success"),
    "REVERSE_BEFORE_SUCCESS": ("冲正早于success",     "reverse 时间早于 success 时间"),
    "REFUND_BEFORE_SUCCESS":  ("退款早于success",     "refund 时间早于 success 时间，该退款无效"),
    "REFUND_ON_REVERSED":     ("已冲正单发生退款",    "单据已被冲正，退款不计入宽表"),
    "OVER_REFUND":            ("累计退款超过原单",    "累计退款 {total} 超过原单金额 {amount}，超额部分不入宽表"),
}


# --------------------------------------------------------------------------
# 1. 读取与行级校验
# --------------------------------------------------------------------------
def parse_amount(text):
    d = Decimal(text.strip())
    if not d.is_finite() or d < 0:
        raise ValueError("金额须为非负有限数: %r" % text)
    return d


def parse_ts(text):
    text = text.strip()
    for fmt in TS_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise ValueError("无法解析时间: %r" % text)


def read_events(csv_path):
    """返回 (orders, anomalies)。orders: pay_id -> 订单聚合结构。"""
    orders = {}
    anomalies = []

    def add_bad_row(line_no, raw, reason):
        anomalies.append({
            "code": "BAD_ROW",
            "channel": (raw[1].strip() if len(raw) > 1 else ""),
            "pay_id": (raw[0].strip() if raw and raw[0].strip()
                       else "(行%d)" % line_no),
            "amount": None, "total": None, "excess": None,
            "lines": [line_no], "detail": reason,
        })

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        for line_no, row in enumerate(reader, start=1):
            if not row or all(not c.strip() for c in row):
                continue
            if line_no == 1 and row[0].strip().lower() == "pay_id":
                continue
            if len(row) != 5:
                add_bad_row(line_no, row, "字段数=%d，应为5" % len(row))
                continue
            pay_id, channel, event, amount_s, ts_s = (c.strip() for c in row)
            if not pay_id or not channel:
                add_bad_row(line_no, row, "支付单号/渠道为空")
                continue
            if event not in VALID_EVENTS:
                add_bad_row(line_no, row, "未知事件 %r" % event)
                continue
            try:
                amount = parse_amount(amount_s)
                ts = parse_ts(ts_s)
            except (InvalidOperation, ValueError) as exc:
                add_bad_row(line_no, row, str(exc))
                continue

            order = orders.get(pay_id)
            if order is None:
                order = {"channel": channel, "channels": {channel},
                         "events": [], "lines": []}
                orders[pay_id] = order
            order["channels"].add(channel)
            order["events"].append((ts, line_no, event, amount))
            order["lines"].append(line_no)

    return orders, anomalies


# --------------------------------------------------------------------------
# 2. 单据级状态机：success / refund / reverse 语义落地
# --------------------------------------------------------------------------
def process_orders(orders, anomalies):
    """返回 daily: (channel, date) -> 累加器。异常追加进 anomalies。"""
    daily = defaultdict(lambda: {
        "succ_cnt": 0, "succ_amt": Decimal("0"),
        "ref_cnt": 0, "ref_amt": Decimal("0"),
        "rev_cnt": 0, "rev_amt": Decimal("0"),
    })

    def add_anom(code, pay_id, order, *, amount=None, total=None,
                 excess=None, detail=None, lines=None):
        anomalies.append({
            "code": code, "channel": order["channel"], "pay_id": pay_id,
            "amount": amount, "total": total, "excess": excess,
            "lines": sorted(set(lines if lines is not None
                                 else order["lines"])),
            "detail": detail or "",
        })

    for pay_id in sorted(orders):
        order = orders[pay_id]
        events = sorted(order["events"], key=lambda e: (e[0], e[1]))

        if len(order["channels"]) > 1:
            add_anom("CHANNEL_CONFLICT", pay_id, order,
                     detail="/".join(sorted(order["channels"])))

        success = reverse = None
        refunds = []
        for ts, line_no, event, amount in events:
            if event == "success":
                if success is None:
                    success = (ts, line_no, amount)
                else:
                    add_anom("DUP_SUCCESS", pay_id, order, lines=[line_no])
            elif event == "reverse":
                if reverse is None:
                    reverse = (ts, line_no, amount)
                else:
                    add_anom("DUP_REVERSE", pay_id, order, lines=[line_no])
            else:  # refund
                refunds.append((ts, line_no, amount))

        # 没有 success：所有事件都是孤儿，不进宽表
        if success is None:
            if reverse is not None:
                add_anom("ORPHAN_REVERSE", pay_id, order,
                         lines=[reverse[1]])
            if refunds:
                add_anom("ORPHAN_REFUND", pay_id, order,
                         total=sum((r[2] for r in refunds), Decimal("0")),
                         lines=[r[1] for r in refunds])
            continue

        succ_ts, succ_line, succ_amt = success
        key = (order["channel"], succ_ts.date())
        acc = daily[key]
        acc["succ_cnt"] += 1
        acc["succ_amt"] += succ_amt

        if reverse is not None:
            rev_ts, rev_line, _rev_amt = reverse
            if rev_ts < succ_ts:
                add_anom("REVERSE_BEFORE_SUCCESS", pay_id, order,
                         amount=succ_amt, lines=[succ_line, rev_line])
            # 冲正抹掉整单 success：冲正笔数+1，净额扣回原单全额
            acc["rev_cnt"] += 1
            acc["rev_amt"] += succ_amt
            if refunds:
                add_anom("REFUND_ON_REVERSED", pay_id, order,
                         amount=succ_amt,
                         total=sum((r[2] for r in refunds), Decimal("0")),
                         lines=[r[1] for r in refunds] + [rev_line])
            continue

        # 未冲正单据：逐条退款，累计封顶 = 原单金额
        refunded = Decimal("0")
        early_lines = []
        over = False
        for ts, line_no, amt in refunds:
            if ts < succ_ts:
                early_lines.append(line_no)
                continue
            room = succ_amt - refunded
            if amt > room:
                over = True
            valid = min(amt, room)
            if valid > 0:
                refunded += valid
                acc["ref_cnt"] += 1
                acc["ref_amt"] += valid
        if early_lines:
            add_anom("REFUND_BEFORE_SUCCESS", pay_id, order,
                     amount=succ_amt, lines=early_lines + [succ_line])
        if over:
            total_ref = sum((r[2] for r in refunds if r[0] >= succ_ts),
                            Decimal("0"))
            add_anom("OVER_REFUND", pay_id, order, amount=succ_amt,
                     total=total_ref, excess=total_ref - succ_amt,
                     lines=[r[1] for r in refunds] + [succ_line])

    return daily


# --------------------------------------------------------------------------
# 3. 汇总成行（含比率，Decimal 精确运算）
# --------------------------------------------------------------------------
def build_daily_rows(daily):
    rows = []
    for channel, day in sorted(daily):
        a = daily[(channel, day)]
        succ_cnt, rev_cnt = a["succ_cnt"], a["rev_cnt"]
        succ_amt, ref_amt, rev_amt = a["succ_amt"], a["ref_amt"], a["rev_amt"]
        net = succ_amt - ref_amt - rev_amt
        succ_rate = (Decimal(succ_cnt - rev_cnt) / Decimal(succ_cnt)
                     if succ_cnt else Decimal("0"))
        ref_rate = (ref_amt / succ_amt if succ_amt > 0 else Decimal("0"))
        rows.append([
            channel, day.isoformat(), succ_cnt,
            succ_amt.quantize(MONEY_Q), net.quantize(MONEY_Q),
            a["ref_cnt"], ref_amt.quantize(MONEY_Q), rev_cnt,
            succ_rate, ref_rate,
        ])
    return rows


def build_anomaly_rows(anomalies):
    """同一支付单的多个异常合并成一行，按 (渠道, 支付单号, 异常类型) 排序。"""
    grouped = {}
    for a in anomalies:
        key = (a["channel"], a["pay_id"])
        g = grouped.get(key)
        if g is None:
            g = {"codes": [], "amount": None, "total": None,
                 "excess": None, "lines": set(), "details": []}
            grouped[key] = g
        g["codes"].append(a["code"])
        g["lines"].update(a["lines"])
        for name in ("amount", "total", "excess"):
            if a[name] is not None:
                g[name] = a[name]
        label, template = ANOMALY_TYPES[a["code"]]
        detail = template.format(detail=a["detail"],
                                 amount=a["amount"], total=a["total"])
        if detail not in g["details"]:
            g["details"].append(detail)

    def q(v):
        return v.quantize(MONEY_Q) if isinstance(v, Decimal) else ""

    rows = []
    for channel, pay_id in sorted(grouped):
        g = grouped[(channel, pay_id)]
        # 去重保序，按 ANOMALY_TYPES 的固定顺序排，保证可复现
        seen = set()
        codes = [c for c in ANOMALY_TYPES
                 if c in g["codes"] and not (c in seen or seen.add(c))]
        labels = [ANOMALY_TYPES[c][0] for c in codes]
        rows.append([
            channel, pay_id, "；".join(labels), "；".join(g["details"]),
            q(g["amount"]), q(g["total"]), q(g["excess"]),
            ",".join(str(n) for n in sorted(g["lines"])),
        ])
    return rows


# --------------------------------------------------------------------------
# 4. 极简确定性 XLSX writer（仅标准库 zipfile + XML，inlineStr 免共享串表）
# --------------------------------------------------------------------------
def _col_letter(idx):
    letters = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _esc(text):
    return (str(text).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _sheet_xml(rows, styles):
    """rows: list[list[cell]]; cell = (value, style_idx or None)。"""
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/'
             'spreadsheetml/2006/main"><sheetData>']
    for r_idx, row in enumerate(rows, start=1):
        parts.append('<row r="%d">' % r_idx)
        for c_idx, (value, style) in enumerate(row):
            ref = "%s%d" % (_col_letter(c_idx), r_idx)
            s = (' s="%d"' % style) if style is not None else ""
            if value is None or value == "":
                continue
            if isinstance(value, bool):
                parts.append('<c r="%s"%s t="b"><v>%d</v></c>'
                             % (ref, s, int(value)))
            elif isinstance(value, int):
                parts.append('<c r="%s"%s><v>%d</v></c>' % (ref, s, value))
            elif isinstance(value, Decimal):
                parts.append('<c r="%s"%s><v>%s</v></c>'
                             % (ref, s, format(value, "f")))
            else:
                parts.append('<c r="%s"%s t="inlineStr"><is><t xml:space='
                             '"preserve">%s</t></is></c>'
                             % (ref, s, _esc(value)))
        parts.append("</row>")
    parts.append("</sheetData></worksheet>")
    return "".join(parts)


def _workbook_xml(sheet_names):
    sheets = "".join(
        '<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
        % (_esc(n), i + 1, i + 1) for i, n in enumerate(sheet_names))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/'
            'spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
            '2006/relationships"><sheets>%s</sheets></workbook>' % sheets)


def _rels_xml(n):
    rels = "".join(
        '<Relationship Id="rId%d" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relations'
        'hips/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1)
        for i in range(n))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships">%s</Relationships>' % rels)


def _styles_xml():
    # style 0: 默认; 1: 表头加粗; 2: 金额两位小数; 3: 百分比
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/'
            'spreadsheetml/2006/main">'
            '<numFmts count="2">'
            '<numFmt numFmtId="164" formatCode="#,##0.00"/>'
            '<numFmt numFmtId="165" formatCode="0.00%"/>'
            '</numFmts>'
            '<fonts count="2">'
            '<font><sz val="11"/><name val="Calibri"/></font>'
            '<font><b/><sz val="11"/><name val="Calibri"/></font>'
            '</fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/>'
            '</border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
            'borderId="0"/></cellStyleXfs>'
            '<cellXfs count="4">'
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" '
            'applyFont="1"/>'
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" '
            'applyNumberFormat="1"/>'
            '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" '
            'applyNumberFormat="1"/>'
            '</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" '
            'builtinId="0"/></cellStyles>'
            '</styleSheet>')


def write_xlsx(path, sheets):
    """sheets: [(sheet_name, rows)]，rows 元素为 (value, style_idx) 元组。"""
    n = len(sheets)
    members = [
        ("[Content_Types].xml",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
         'content-types">'
         '<Default Extension="rels" ContentType="application/vnd.'
         'openxmlformats-package.relationships+xml"/>'
         '<Default Extension="xml" ContentType="application/xml"/>'
         '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.'
         'openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
         '<Override PartName="/xl/styles.xml" ContentType="application/vnd.'
         'openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
         + "".join('<Override PartName="/xl/worksheets/sheet%d.xml" '
                   'ContentType="application/vnd.openxmlformats-officedocument.'
                   'spreadsheetml.worksheet+xml"/>' % (i + 1)
                   for i in range(n))
         + "</Types>"),
        ("_rels/.rels",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
         '2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
         'officeDocument/2006/relationships/officeDocument" '
         'Target="xl/workbook.xml"/></Relationships>'),
        ("xl/workbook.xml", _workbook_xml([name for name, _ in sheets])),
        ("xl/_rels/workbook.xml.rels", _rels_xml(n)),
        ("xl/styles.xml", _styles_xml()),
    ]
    for i, (_name, rows) in enumerate(sheets):
        members.append(("xl/worksheets/sheet%d.xml" % (i + 1),
                        _sheet_xml(rows, None)))

    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED,
                         compresslevel=9) as zf:
        for name, content in members:
            info = zipfile.ZipInfo(filename=name,
                                   date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = (0o644 & 0xFFFF) << 16
            zf.writestr(info, content)
    os.replace(tmp, path)


def style_rows(header, data_rows, money_cols=(), pct_cols=(), int_cols=()):
    """把纯值行包成 (value, style) 行；列下标为 0-based。"""
    out = [[(h, 1) for h in header]]
    for row in data_rows:
        styled = []
        for i, v in enumerate(row):
            if i in money_cols:
                styled.append((v, 2))
            elif i in pct_cols:
                styled.append((v, 3))
            elif i in int_cols:
                styled.append((v, 0))
            else:
                styled.append((v, 0))
        out.append(styled)
    return out


# --------------------------------------------------------------------------
# 5. 主流程
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(
        description="支付事件流 -> 渠道日标签宽表 (channel_daily.xlsx)")
    ap.add_argument("input", help="事件流 CSV 路径")
    ap.add_argument("-o", "--output", default="channel_daily.xlsx",
                    help="输出 xlsx 路径（默认 ./channel_daily.xlsx）")
    args = ap.parse_args()

    orders, anomalies = read_events(args.input)
    daily = process_orders(orders, anomalies)
    daily_rows = build_daily_rows(daily)
    anomaly_rows = build_anomaly_rows(anomalies)

    sheet1 = style_rows(DAILY_HEADER, daily_rows,
                        money_cols=(3, 4, 6), pct_cols=(8, 9),
                        int_cols=(2, 5, 7))
    sheet2 = style_rows(ANOMALY_HEADER, anomaly_rows,
                        money_cols=(4, 5, 6))
    write_xlsx(args.output, [("渠道日标签", sheet1),
                             ("异常单登记", sheet2)])

    total_succ = sum(r[2] for r in daily_rows)
    total_rev = sum(r[7] for r in daily_rows)
    print("输出: %s" % args.output)
    print("渠道日行数: %d，成功笔数: %d，冲正笔数: %d，异常单数: %d"
          % (len(daily_rows), total_succ, total_rev, len(anomaly_rows)))


if __name__ == "__main__":
    main()
