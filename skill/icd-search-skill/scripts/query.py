#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ICD-10 & ICD-9-CM-3 (Hubei 2.0) Search Engine.
Pure Python standard library. Cross-platform compatible (Windows, Linux, macOS).
"""

import sys
import re
import argparse
import json
import sqlite3
from pathlib import Path

# Ensure UTF-8 output on all platforms
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Surgical action indicators
SURGERY_KEYWORDS = {
    "术", "切除", "结扎", "修补", "置管", "吻合", "镜", "引流", "成形", 
    "植入", "再造", "清扫", "活检", "消融", "造瘘", "取石", "截骨", 
    "减压", "缝合", "截肢", "穿刺", "操作", "灌洗", "透析", "插管", 
    "移植", "剥脱", "抽吸", "旋切", "切开", "复位", "松解", "闭合",
    "扩创", "包扎", "探查", "离断", "游离", "缝扎", "减容", "切取"
}

DISEASE_KEYWORDS = {
    "病", "炎", "症", "综合征", "综合症", "癌", "瘤", "坏死", "衰竭", 
    "梗死", "结石", "损伤", "骨折", "裂伤", "穿孔", "积液", "血肿", 
    "畸形", "感染", "脱垂", "狭窄", "缺血", "休克", "疝", "瘘", "囊肿",
    "息肉", "溃疡", "出血", "栓塞", "血栓", "血肿", "脓肿", "水肿"
}

def get_db_path() -> Path:
    # 1. Look in root directory (parent of scripts/)
    root_dir = Path(__file__).resolve().parent.parent
    db_path = root_dir / "icd_hubei2.db"
    if db_path.exists():
        return db_path
    # 2. Look in same directory as script
    db_same_dir = Path(__file__).resolve().parent / "icd_hubei2.db"
    if db_same_dir.exists():
        return db_same_dir
    # 3. Look in current working directory
    db_cwd = Path.cwd() / "icd_hubei2.db"
    if db_cwd.exists():
        return db_cwd
    raise FileNotFoundError(f"Database 'icd_hubei2.db' not found in {root_dir}")

def detect_mixed_input(text: str) -> bool:
    """
    Check if user input simultaneously requests both disease and surgery.
    e.g., '高血压和阑尾切除术', '诊断: 胃癌, 手术: 胃大部切除'
    """
    t = text.strip()
    # Explicit dual intent markers
    mixed_patterns = [
        r"(?:诊断|疾病)[：:].+?(?:手术|操作)[：:]",
        r"(?:手术|操作)[：:].+?(?:诊断|疾病)[：:]",
        r".+?(?:和|及|并|同时查|以及).+?(?:术|切除|结扎|修补|引流|穿刺|消融)",
    ]
    for p in mixed_patterns:
        if re.search(p, t):
            return True

    # Check if there are distinct disease phrases AND surgical phrases connected
    # e.g., "高血压 阑尾切除"
    tokens = re.split(r"[\s,，、；;+＋/]+", t)
    if len(tokens) >= 2:
        has_pure_disease = False
        has_pure_surgery = False
        for token in tokens:
            if not token:
                continue
            is_surg = any(k in token for k in SURGERY_KEYWORDS)
            is_dis = any(k in token for k in DISEASE_KEYWORDS)
            if is_surg and not token.endswith(("病", "症", "综合征")):
                has_pure_surgery = True
            elif is_dis and not is_surg:
                has_pure_disease = True
        if has_pure_disease and has_pure_surgery:
            return True

    return False

def detect_intent(text: str) -> str:
    """
    Infer if query is for surgery or disease.
    """
    t = text.strip()
    # If ends with surgical suffix, almost certainly surgery (e.g. 乳腺癌根治术)
    if any(t.endswith(s) for s in ["术", "切除", "修补", "结扎", "成形", "吻合", "引流", "置管", "穿刺", "消融", "旋切", "剥脱"]):
        return "surgery"
    if any(k in t for k in SURGERY_KEYWORDS):
        return "surgery"
    return "disease"

def clean_tokens(query: str):
    cleaned = re.sub(r"[\s,，、。+＋\-_/\\()（）\[\]【】]+", " ", query).strip()
    tokens = [t for t in cleaned.split(" ") if t]
    return tokens if tokens else [query]

def search_table(conn: sqlite3.Connection, table: str, query: str, limit: int = 5):
    cur = conn.cursor()
    q = query.strip()
    is_surgery = (table == "surgeries")

    # 1. Exact code match
    sql_code = f"SELECT code, name{', level' if is_surgery else ''} FROM {table} WHERE code = ?"
    exact_codes = cur.execute(sql_code, (q,)).fetchall()
    if exact_codes:
        results = []
        for r in exact_codes:
            item = {"code": r[0], "name": r[1], "score": 1000.0, "type": "surgery" if is_surgery else "disease"}
            if is_surgery:
                item["level"] = r[2]
            results.append(item)
        return results

    # 2. Exact name match
    sql_name = f"SELECT code, name{', level' if is_surgery else ''} FROM {table} WHERE name = ?"
    exact_names = cur.execute(sql_name, (q,)).fetchall()
    if exact_names:
        results = []
        for r in exact_names:
            item = {"code": r[0], "name": r[1], "score": 800.0, "type": "surgery" if is_surgery else "disease"}
            if is_surgery:
                item["level"] = r[2]
            results.append(item)
        return results

    tokens = clean_tokens(q)
    
    # Generate candidate segments
    all_terms = list(tokens)
    for t in tokens:
        if len(t) >= 4:
            # generate bi-grams and tri-grams
            for i in range(len(t) - 1):
                all_terms.append(t[i:i+2])
            for i in range(len(t) - 2):
                all_terms.append(t[i:i+3])

    candidates = []
    seen_codes = set()

    # Step 2a: Try matching all original tokens (AND)
    token_clauses = ["name LIKE ?" for _ in tokens]
    token_params = [f"%{t}%" for t in tokens]
    sql_and = f"SELECT code, name{', level' if is_surgery else ''} FROM {table} WHERE {' AND '.join(token_clauses)} LIMIT 100"
    for r in cur.execute(sql_and, token_params).fetchall():
        if r[0] not in seen_codes:
            candidates.append(r)
            seen_codes.add(r[0])

    # Step 2b: If not enough results, use subset of keywords (OR)
    if len(candidates) < 50 and len(all_terms) > 1:
        # Sort terms by length descending, pick most specific
        unique_terms = sorted(list(set(all_terms)), key=lambda x: len(x), reverse=True)[:8]
        or_clauses = ["name LIKE ?" for _ in unique_terms]
        or_params = [f"%{k}%" for k in unique_terms]
        sql_or = f"SELECT code, name{', level' if is_surgery else ''} FROM {table} WHERE {' OR '.join(or_clauses)} LIMIT 150"
        for r in cur.execute(sql_or, or_params).fetchall():
            if r[0] not in seen_codes:
                candidates.append(r)
                seen_codes.add(r[0])

    # Step 2c: If still empty, search query as substring
    if not candidates:
        sql_sub = f"SELECT code, name{', level' if is_surgery else ''} FROM {table} WHERE name LIKE ? LIMIT 50"
        for r in cur.execute(sql_sub, (f"%{q}%",)).fetchall():
            if r[0] not in seen_codes:
                candidates.append(r)
                seen_codes.add(r[0])

    # 3. Score & Rank candidates
    scored = []
    q_chars = set(q)
    for row in candidates:
        code = row[0]
        name = row[1]
        level = row[2] if is_surgery else None

        score = 0.0
        if name == q:
            score += 500.0
        elif name.startswith(q):
            score += 300.0 + (len(q) / len(name)) * 60.0
        elif q in name:
            score += 200.0 + (len(q) / len(name)) * 60.0
        elif name in q:
            score += 150.0 + (len(name) / len(q)) * 40.0

        # Character overlap score
        matched_chars = sum(1 for ch in q_chars if ch in name)
        overlap_ratio = matched_chars / max(len(q_chars), 1)
        score += overlap_ratio * 100.0

        # Token match bonus
        for t in tokens:
            if t in name:
                score += len(t) * 15.0

        # Order preservation bonus
        last_idx = -1
        in_order = True
        for t in tokens:
            idx = name.find(t)
            if idx == -1 or idx < last_idx:
                in_order = False
                break
            last_idx = idx
        if in_order and len(tokens) > 1:
            score += 40.0

        # Length penalty for excessive clutter
        len_diff = max(len(name) - len(q), 0)
        score -= min(len_diff * 1.5, 35.0)

        item = {
            "code": code,
            "name": name,
            "score": round(score, 2),
            "type": "surgery" if is_surgery else "disease"
        }
        if is_surgery:
            item["level"] = level
        scored.append(item)

    scored.sort(key=lambda x: x["score"], reverse=True)
    return scored[:limit]

def format_item(item: dict) -> str:
    """Format single result per user requirement:
    - Surgery: 淋巴结扩大性区域性切除术（40.3x00x001） [3级]
    - Disease: 疾病名称（疾病编码）
    """
    code = item["code"]
    name = item["name"]
    if item.get("type") == "surgery":
        lvl = item.get("level")
        lvl_str = f" [{lvl}级]" if lvl is not None else ""
        return f"{name}（{code}）{lvl_str}"
    else:
        return f"{name}（{code}）"

def main():
    parser = argparse.ArgumentParser(description="Query ICD Hubei 2.0 database")
    parser.add_argument("query", nargs="*", help="Query term or ICD code")
    parser.add_argument("-t", "--type", choices=["surgery", "surg", "s", "disease", "dis", "d", "auto"],
                        default="auto", help="Search target: surgery, disease, or auto (default: auto)")
    parser.add_argument("-n", "--limit", type=int, default=5, help="Max candidates (default: 5)")
    parser.add_argument("--json", action="store_true", help="Output JSON format")
    args = parser.parse_args()

    query_str = " ".join(args.query).strip()
    if not query_str:
        print("Error: query string is empty.", file=sys.stderr)
        sys.exit(1)

    # Check for mixed input rejection only when type is auto
    if args.type == "auto" and detect_mixed_input(query_str):
        msg = "拒绝查询：检测到您的输入同时包含疾病与手术（或意图不明确），请明确本次是查询【疾病诊断】还是【手术操作】，或分开发送查询。"
        if args.json:
            print(json.dumps({"error": "mixed_input", "message": msg}, ensure_ascii=False, indent=2))
        else:
            print(msg)
        sys.exit(0)

    # Determine search type
    t = args.type.lower()
    if t in ("surgery", "surg", "s"):
        target_type = "surgery"
    elif t in ("disease", "dis", "d"):
        target_type = "disease"
    else:
        target_type = detect_intent(query_str)

    table = "surgeries" if target_type == "surgery" else "diseases"

    db_path = get_db_path()
    conn = sqlite3.connect(str(db_path))

    results = search_table(conn, table, query_str, limit=args.limit)
    conn.close()

    if args.json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        if not results:
            print(f"未在湖北2.0版{ '手术' if target_type == 'surgery' else '疾病' }库中找到相关结果。")
            return

        # If highest score >= 500 (exact match) and only 1 candidate, or user only wants top
        # Check score gap: if top match is clearly ahead (e.g. exact match), can show top 1
        if len(results) == 1 or results[0]["score"] >= 800.0:
            print(format_item(results[0]))
        else:
            for idx, r in enumerate(results, 1):
                print(f"{idx}. {format_item(r)}")

if __name__ == "__main__":
    main()
