# -*- coding: utf-8 -*-
"""dip-advisor 统一查询脚本（零第三方依赖）。

用法：
  python search.py diag <关键词>      诊断编码查询，Top 5：编码 名称
  python search.py group <诊断编码>   该诊断可入的全部病组（核心/综合/兜底），按分值降序
  python search.py family <编码前3位> 该类目下所有核心病组键及其最高分（诊断调整建议用）

付费标准系数唯一改动点在此（单位：元 / 分值）：
"""
import json
import re
import sys
from pathlib import Path

EMPLOYEE_RATE = 5.69
RESIDENT_RATE = 4.36

DATA = Path(__file__).parent.parent / "data"

DIAGNOSES = None
CORE = None
COMP = None


def load():
    global DIAGNOSES, CORE, COMP
    if DIAGNOSES is None:
        DIAGNOSES = json.loads((DATA / "diagnoses.json").read_text(encoding="utf-8"))
        CORE = json.loads((DATA / "core.json").read_text(encoding="utf-8"))
        COMP = json.loads((DATA / "comp.json").read_text(encoding="utf-8"))


def norm(code):
    return re.sub(r"[^A-Za-z0-9]", "", code).upper()


def fmt_score(score):
    return f"{score:.2f}"


def cmd_diag(keyword):
    """诊断查询：名称子串 + 编码前缀，合并去重，Top 5。"""
    kw = keyword.strip()
    kw_norm = norm(kw)
    hits = []
    for code, name in DIAGNOSES:
        if len(hits) >= 5:
            break
        if kw in name or (kw_norm and norm(code).startswith(kw_norm)):
            hits.append(f"{code} {name}")
    if not hits:
        print("(无匹配)")
    else:
        print("\n".join(hits))


def fmt_row(kind, code, name, score):
    emp = float(score) * EMPLOYEE_RATE
    res = float(score) * RESIDENT_RATE
    return (f"[{kind}] {code} | {name} | 分值 {fmt_score(score)}"
            f" | 职工 {emp:.2f} 元 | 居民 {res:.2f} 元")


def cmd_group(code):
    """查某诊断可入的全部病组：前4位核心 + 前3位综合（未命中首字母兜底），分值降序。"""
    n = norm(code)
    if len(n) < 3:
        print("(编码过短，无法匹配)")
        return
    rows = []
    core_key = n[:4]
    if core_key in CORE:
        for c, nm, s in CORE[core_key]:
            rows.append(("核心", c, nm, s))
    comp_key = n[:3]
    fallback = False
    if comp_key in COMP:
        comp_rows = COMP[comp_key]
    else:
        comp_rows = COMP.get(n[0], [])
        fallback = bool(comp_rows)
    for c, nm, s in comp_rows:
        rows.append(("综合(兜底)" if fallback else "综合", c, nm, s))
    if not rows:
        print("(该病组目录中未找到匹配病组)")
        return
    rows.sort(key=lambda r: -float(r[3]))
    for kind, c, nm, s in rows:
        print(fmt_row(kind, c, nm, s))


def cmd_family(prefix):
    """列出该前3位类目下所有核心病组键及最高分，供诊断调整建议。"""
    p = norm(prefix)[:3]
    keys = sorted(k for k in CORE if k.startswith(p))
    if not keys:
        print("(该类目下无核心病组)")
        return
    for k in keys:
        best = max(CORE[k], key=lambda r: float(r[2]))
        print(f"{k} 最高分 {fmt_score(best[2])} | {best[0]} | {best[1]}")


def main(argv):
    load()
    if len(argv) < 3:
        print(__doc__)
        return
    cmd, arg = argv[1], argv[2]
    if cmd == "diag":
        cmd_diag(arg)
    elif cmd == "group":
        cmd_group(arg)
    elif cmd == "family":
        cmd_family(arg)
    else:
        print(__doc__)


if __name__ == "__main__":
    main(sys.argv)
