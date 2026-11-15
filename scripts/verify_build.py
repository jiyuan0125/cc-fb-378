#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_build.py — 回读 xlsx 并对 build_tags.py 的口径做独立断言。"""
import hashlib
import subprocess
import sys
import zipfile
from decimal import Decimal
from xml.etree import ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SCRIPT = "/workspace/scripts/build_tags.py"
CSV = "/workspace/testdata/sample_events.csv"
OUT = "/workspace/testdata/out.xlsx"


def read_sheet(path, sheet_idx):
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        assert bad is None, "zip 损坏: %s" % bad
        xml = zf.read("xl/worksheets/sheet%d.xml" % sheet_idx)
    root = ET.fromstring(xml)  #  XML 必须良构
    rows = []
    for row in root.iter(NS + "row"):
        vals = []
        for c in row.findall(NS + "c"):
            t = c.get("t")
            if t == "inlineStr":
                vals.append(c.find(NS + "is/" + NS + "t").text)
            else:
                vals.append(Decimal(c.find(NS + "v").text))
        rows.append(vals)
    return rows


def main():
    # 重跑两遍，字节级一致
    hashes = []
    for _ in range(2):
        subprocess.run([sys.executable, "-I", SCRIPT, CSV, "-o", OUT],
                       check=True, capture_output=True)
        with open(OUT, "rb") as fh:
            hashes.append(hashlib.sha256(fh.read()).hexdigest())
    assert hashes[0] == hashes[1], "重跑产物不一致: %s" % hashes
    print("[ok] 重跑字节级一致 sha256=%s" % hashes[0][:16])

    daily = read_sheet(OUT, 1)
    header, data = daily[0], daily[1:]
    assert header == ["渠道", "日期", "成功笔数", "成功金额", "净额",
                      "退款笔数", "退款金额", "冲正笔数", "成功率", "退款率"]

    # 手工推算的期望行: 渠道,日期, 成功笔数,成功金额,净额, 退款笔数,退款金额,
    #                  冲正笔数, 成功率, 退款率
    expect = [
        ["alipay", "2026-08-20", 2, "200.00", "50.00", 2, "70.00", 1,
         Decimal(1) / Decimal(2), Decimal("0.35")],
        ["ccbpay", "2026-08-20", 1, "45.00", "45.00", 0, "0.00", 0,
         Decimal(1), Decimal(0)],
        ["ccbpay", "2026-08-21", 1, "33.33", "0.00", 2, "33.33", 0,
         Decimal(1), Decimal(1)],
        ["cmbpay", "2026-08-21", 1, "100.00", "70.00", 1, "30.00", 0,
         Decimal(1), Decimal("0.3")],
        ["jdpay", "2026-08-20", 1, "70.00", "70.00", 0, "0.00", 0,
         Decimal(1), Decimal(0)],
        ["unionpay", "2026-08-20", 2, "110.00", "0.00", 1, "50.00", 1,
         Decimal(1) / Decimal(2), Decimal(50) / Decimal(110)],
        ["wxpay", "2026-08-20", 4, "365.00", "55.00", 2, "110.00", 1,
         Decimal(3) / Decimal(4), Decimal(110) / Decimal(365)],
    ]
    assert len(data) == len(expect), "行数 %d != %d" % (len(data), len(expect))
    for got, want in zip(data, expect):
        assert got[0] == want[0] and got[1] == want[1], "键错位: %s" % got[:2]
        for i in (2, 3, 4, 5, 6, 7):
            assert got[i] == Decimal(want[i]), \
                "%s %s 列%d: %s != %s" % (want[0], want[1], i, got[i], want[i])
        for i in (8, 9):
            assert abs(got[i] - want[i]) < Decimal("1e-30"), \
                "%s %s 比率列%d: %s != %s" % (want[0], want[1], i, got[i], want[i])
    print("[ok] 渠道日宽表 %d 行全部口径断言通过" % len(data))

    # 异常单登记
    anomalies = read_sheet(OUT, 2)
    assert anomalies[0] == ["渠道", "支付单号", "异常类型", "异常说明",
                            "原单金额", "累计退款金额", "超额金额", "CSV行号"]
    by_id = {r[1]: r for r in anomalies[1:]}
    def atype(pid): return by_id[pid][2]
    assert "重复success" in atype("W003")
    assert "累计退款超过原单" in atype("U001")
    assert by_id["U001"][6] == Decimal("20.00"), "U001 超额应为 20.00"
    assert "已冲正单发生退款" in atype("U002")
    assert "退款无对应success" in atype("C001")
    assert "冲正无对应success" in atype("J001")
    assert "退款早于success" in atype("J002")
    assert "同单渠道不一致" in atype("X001")
    assert "行数据非法" in atype("Z001")
    assert any("行数据非法" in r[2] for r in anomalies[1:]), "6 列脏行未登记"
    # 正常单不在异常表
    for ok_id in ("W001", "W002", "A001", "A002", "C002", "M002"):
        assert ok_id not in by_id, "%s 不应是异常单" % ok_id
    print("[ok] 异常单登记 %d 条，类型与超额金额断言通过" % (len(anomalies) - 1))

    # 结构完整性：Content_Types / rels 齐备且 XML 良构
    with zipfile.ZipFile(OUT) as zf:
        names = set(zf.namelist())
        for need in ("[Content_Types].xml", "_rels/.rels",
                     "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
                     "xl/styles.xml", "xl/worksheets/sheet1.xml",
                     "xl/worksheets/sheet2.xml"):
            assert need in names, "缺少 %s" % need
        for n in names:
            if n.endswith(".xml") or n.endswith(".rels"):
                ET.fromstring(zf.read(n))
    print("[ok] xlsx 包结构与全部 XML 良构")
    print("全部验证通过")


if __name__ == "__main__":
    main()
