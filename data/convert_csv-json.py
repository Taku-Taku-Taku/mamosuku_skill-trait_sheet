# convert_csv-json.py
import csv, json, re, unicodedata
from pathlib import Path

SRC_FILES = [
    ("magi_common.csv", "magi_common.json", "common"),
    ("magi_pc.csv",     "magi_pc.json",     "pc"),
    ("magi_clan.csv",   "magi_clan.json",   "clan"),
]

def normalize_condition(s: str):
    """条件文字列を機械可読に正規化して、(kind, min, max, values) を返す"""
    if s is None:
        s = ""
    s = s.strip()

    # 全角→半角、長音や波ダッシュのゆらぎ、区切りのゆらぎを吸収
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("〜", "-").replace("~", "-")
    s = s.replace("・", ",").replace("，", ",").replace("、", ",")
    s = re.sub(r"\s+", "", s)  # 余分な空白は削除

    if s == "" or s == "なし":
        return ("none", None, None, None)

    # 1-6, 3-10 のような範囲("~"->"-"に変換済み)
    m = re.fullmatch(r"(\d+)-(\d+)", s)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        if a > b:  # 念のため入れ替え
            a, b = b, a
        return ("range", a, b, None)

    # 1,3,5,7 のような列挙
    if re.fullmatch(r"\d+(,\d+)+", s):
        arr = [int(x) for x in s.split(",")]
        return ("list", None, None, arr)

    # 1～2体 など数値以外を含むケース → text として残す
    return ("text", None, None, None)

def row_to_json(row: dict, category: str, idx: int):
    # CSV列名に合わせて拾う（前に作ったCSVに準拠）
    name = row.get("name","").strip()
    timing = row.get("timing","").strip()
    target = row.get("target","").strip()
    condition_raw = row.get("condition","").strip()
    effect = row.get("effect","").strip()
    description = row.get("description","").strip()

    kind, cmin, cmax, cvals = normalize_condition(condition_raw)

    return {
        "id": f"{category}-{idx:03d}",
        "category": category,         # common / pc / clan
        "name": name,
        "timing": timing,
        "target": target,
        "condition_raw": condition_raw,
        "condition_kind": kind,       # none / range / list / text
        "cond_min": cmin,
        "cond_max": cmax,
        "cond_values": cvals,         # 例: [2,4,6,8,10,12]
        "effect": effect,
        "description": description,
        # 簡易検索用のまとめ文字列（必要なら）
        "keywords": " ".join([name, timing, target, condition_raw, effect, description]),
    }

def convert_csv_to_json(csv_path: Path, json_path: Path, category: str):
    out = []
    with csv_path.open(newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader, 1):
            out.append(row_to_json(row, category, i))
    json_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✔ {csv_path.name} → {json_path.name} ({len(out)}件)")

if __name__ == "__main__":
    for src, dst, cat in SRC_FILES:
        convert_csv_to_json(Path(src), Path(dst), cat)
