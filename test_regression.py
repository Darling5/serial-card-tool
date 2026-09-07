# -*- coding: utf-8 -*-
"""回归测试：serial_card_tool 解析引擎

用法：python test_regression.py <串口日志TXT> [池卡明细.xlsx]
基准数值来自 26621 行真实生产日志（369 配对 / 3 失败 / 371 设备 / 129 缺失）。
"""
import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from serial_card_tool import ParserEngine, load_card_table

if len(sys.argv) < 2:
    print("用法: python test_regression.py <串口日志TXT> [池卡明细.xlsx]")
    sys.exit(2)
LOG = sys.argv[1]
POOL = sys.argv[2] if len(sys.argv) > 2 else None
EXCLUDE = ["898604981026D0054007", "89860402102670003351",
           "89860402102670003387", "89860402102670003541"]

fails = 0

def check(name, cond, detail=""):
    global fails
    mark = "PASS" if cond else "FAIL"
    if not cond:
        fails += 1
    print("[%s] %s %s" % (mark, name, detail))

# 1. 引擎解析
eng = ParserEngine()
with open(LOG, encoding="utf-8", errors="replace") as f:
    lines = [l.rstrip("\r\n") for l in f]
eng.feed_lines(lines)
eng.flush()
recs = eng.records()

paired = [r for r in recs if r.iccid]
fail_recs = [r for r in recs if not r.iccid]
check("配对记录数 = 369", len(paired) == 369, "实际 %d" % len(paired))
check("读取失败记录数 = 3", len(fail_recs) == 3, "实际 %d" % len(fail_recs))
check("失败设备 = 110/303/473",
      sorted(r.dev for r in fail_recs) == ["785000060110", "785000060303", "785000060473"])

# 2. 与基准 CSV 逐行 diff（基准文件含真实生产数据不入库，本地存在时才比对）
BASE = os.path.dirname(os.path.abspath(__file__))
got = sorted((r.dev, r.iccid) for r in paired)
ref_csv = os.path.join(BASE, "device_iccid_table.csv")
if os.path.exists(ref_csv):
    ref = []
    with open(ref_csv, encoding="utf-8-sig") as f:
        for i, r in enumerate(csv.reader(f)):
            if i == 0 or not (r and r[0] and r[1]):
                continue
            ref.append((r[0], r[1]))
    if got != sorted(ref):
        from collections import Counter
        cg, cr = Counter(got), Counter(ref)
        diff = {k: (cg.get(k, 0), cr.get(k, 0)) for k in set(cg) | set(cr) if cg.get(k, 0) != cr.get(k, 0)}
        print("    差异明细（工具次数, 基准次数）:", diff)
    check("与离线基准完全一致（369 行）", got == sorted(ref))
else:
    print("[SKIP] 未找到 device_iccid_table.csv，跳过基准逐行比对")

# 3. 流式喂入（模拟串口逐行）结果一致
eng2 = ParserEngine()
for l in lines:
    eng2.feed_lines([l])
eng2.flush()
got2 = sorted((r.dev, r.iccid) for r in eng2.records() if r.iccid)
check("流式逐行喂入与批量一致", got2 == got)

# 4. 统计口径
devs = {}
for r in recs:
    devs.setdefault(r.dev, []).append(r)
read = len(devs)
check("已读设备 = 371", read == 371, "实际 %d" % read)
check("剩余未读取 = 129（总数500）", 500 - read == 129)

start = 785000060001
missing_dev = [d for d in range(start, start + 500) if str(d) not in devs]
check("范围内缺失设备 = 129", len(missing_dev) == 129, "实际 %d" % len(missing_dev))

swap = [d for d, rs in devs.items() if len({r.iccid for r in rs if r.iccid}) >= 2]
check("换卡设备 = 785000060500", swap == ["785000060500"])

# 5. 卡号表 + 排除（池卡明细为可选参数，未提供时跳过）
if POOL:
    card_map = load_card_table(POOL)
    check("卡号表加载 500 张", len(card_map) == 500, "实际 %d" % len(card_map))
    excl = set(EXCLUDE)
    appeared = {r.iccid for r in recs if r.iccid}
    matched = {i for i in appeared if i in card_map and i not in excl}
    check("匹配卡号 = 367", len(matched) == 367, "实际 %d" % len(matched))
    nonpool = {i for i in appeared if i not in card_map}
    check("非池内卡 = 1", len(nonpool) == 1, "实际 %s" % sorted(nonpool))

    miss_cards = [(c, i) for i, c in card_map.items()
                  if i not in appeared and i not in excl]
    check("缺失卡号 = 129", len(miss_cards) == 129, "实际 %d" % len(miss_cards))

    # 与 missing_cards.csv 对比（基准文件本地存在时才比对）
    ref_miss_csv = os.path.join(BASE, "missing_cards.csv")
    if os.path.exists(ref_miss_csv):
        ref_miss = set()
        with open(ref_miss_csv, encoding="utf-8-sig") as f:
            for r in csv.reader(f):
                if r and r[0] and r[2].startswith("缺失"):
                    ref_miss.add((r[0], r[1]))
        check("缺失卡号清单与基准一致", set(miss_cards) == ref_miss)
else:
    print("[SKIP] 未提供池卡明细参数，跳过卡号匹配测试")

print()
print("结果：%s（%d 项失败）" % ("全部通过" if fails == 0 else "存在失败", fails))
sys.exit(1 if fails else 0)
