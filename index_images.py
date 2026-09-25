"""顕微鏡画像のファイル名を患者リストと突き合わせ、患者・組織・染色ごとの一覧 JSON を作る。

使い方:
    uv run index_images.py --csv 患者リスト.csv -o 一覧.json [--root 探すフォルダ]

ファイルの中身は読まない（OneDrive からダウンロードしないので速い）。ファイル名の規則は README の
「顕微鏡画像のファイル名と患者リストの対応」を参照。

出力には患者の番号・年齢・性別とファイル名が入るので、公開リポジトリには置かないこと。
氏名と ID 全体は出力しない。
"""

import argparse
import csv
import json
import re
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

IMAGE_EXTS = {".czi", ".jpg", ".jpeg", ".tif", ".tiff", ".png"}

# ファイル名 → 組織（先に書いたものほど優先）
TISSUES = [
    ("cot", r"coty?(?:loid)?"),
    ("cam", r"cam"),
    ("cap", r"cap"),
    ("lab", r"lab"),
    ("syno", r"syno"),
    ("cartilage", r"軟骨|car(?=\d|\s)"),
]
# ファイル名 → 染色（先に書いたものほど優先）
STAINS = [
    ("RXFP1+Vimentin", r"rx\s*_?vim|rxv"),
    ("negative control", r"neg|(?<![a-z])nc\d*(?![a-z])|(?:cap|lab|cam|coty?|syno)\s*ng|(?<![a-z])ng\d*(?![a-z])"),
    ("positive control", r"posi"),
    ("RXFP1", r"rxfp|rx"),
    ("MT", r"(?<![a-z])mt|masson"),
    ("HE", r"(?<![a-z])he\d|(?<![a-z])he(?![a-z])"),
    ("Alcian blue", r"アルシアン|alcian"),
    ("SOX9", r"sox9"),
    ("OXT/OXTR", r"(?<![a-z])ox"),
    ("PPIB", r"ppib"),
]
PROCEDURES = r"fai|tha|oa|shelf"
# ZEN の書き出し（元の CZI がある可能性が高い）
EXPORT_MARKERS = re.compile(r"エクスポート|image export", re.I)
# ApoTome で撮ったものは蛍光（DAB の免疫染色とは別物）
FLUORESCENCE = re.compile(r"apotome", re.I)
# ヒト以外（マウス・ラット・細胞）は対象外
NOT_HUMAN = re.compile(r"マウス|mouse|mice|(?<![a-z])ms|ovx|rat(?![a-z])|atdc5|uterus|子宮", re.I)


def load_patients(csv_path):
    patients = []
    with open(csv_path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if not row.get("number"):
                continue
            patients.append({
                "number": row["number"].strip(),
                "id": row["ID"].strip(),
                "age": row["age"].strip(),
                "sex": row["性別"].strip().upper(),
                "procedure": row["procedure"].strip().lower(),
                "list_rxfp": {
                    "cap": row.get("RXFP-1 capsule", "").strip(),
                    "cam": row.get("RXFP-1 CAM", "").strip(),
                    "lab": row.get("RXFP-1 labrum", "").strip(),
                    "cot": row.get("RXFP-1 cotyloidfossa", "").strip(),
                },
            })
    return patients


def find_patient(stem, patients):
    """(患者, 名前に書かれた符号, 注意) を返す。見つからなければ患者は None。"""
    s = stem.lower()
    # 名前に書かれた年齢・性別（例: fai16f, tha 80m, THA57f）
    demo = re.search(rf"(?:{PROCEDURES})\s*(\d{{2}})\s*([mf])(?![a-z])", s) or \
        re.search(r"(?<!\d)(\d{2})\s*([mf])(?![a-z])", s)
    age, sex = (demo.group(1), demo.group(2).upper()) if demo else (None, None)

    def check(p, code):
        if age and (p["age"] != age or p["sex"] != sex):
            return p, code, f"名前の年齢・性別 {age}{sex} がリスト（{p['age']}{p['sex']}）と違う"
        return p, code, None

    m = re.match(r"(\d{8})", s)
    if m:
        for p in patients:
            if p["id"] == m.group(1):
                return check(p, m.group(1))
    m = re.match(r"(\d{4})(?!\d)", s)
    if m:
        code = m.group(1)
        hits = [p for p in patients if p["id"].endswith(code)] or \
            [p for p in patients if p["id"].startswith(code)]
        if len(hits) == 1:
            return check(hits[0], code)
    m = re.match(rf"(\d{{1,2}})\s*(?:{PROCEDURES})", s)
    if m:
        for p in patients:
            if p["number"] == str(int(m.group(1))):
                return check(p, m.group(1))
    # 番号・スペース・年齢性別（例: 40 17fcap）
    m = re.match(r"(\d{1,2})\s+(\d{2})\s*([mf])(?![a-z])|(\d{1,2})\s+(\d{2})\s*([mf])(?=cap|cam|lab|cot|syno)", s)
    if m:
        number = m.group(1) or m.group(4)
        for p in patients:
            if p["number"] == str(int(number)):
                return check(p, number)
    # 年齢性別で始まる名前（例: 44F CAM rx3, 76F THA cap nc2）。cam は FAI の組織
    m = re.match(rf"(\d{{2}})\s*([mf])(?![a-z])\s*({PROCEDURES})?", s)
    if m:
        proc = m.group(3) or ("fai" if re.search(r"cam", s) else None)
        proc = "tha" if proc == "oa" else proc
        hits = [p for p in patients if p["age"] == m.group(1) and p["sex"] == m.group(2).upper()
                and (proc is None or p["procedure"] == proc)]
        if len(hits) == 1:
            return hits[0], None, "番号なし。年齢・性別（と組織）から推定"
        if hits:
            return None, None, "候補が複数: " + ", ".join("#" + p["number"] for p in hits)
    # 番号がなく「術式+年齢+性別」だけの名前（例: THA57fcapRX）
    m = re.match(rf"({PROCEDURES})\s*(\d{{2}})\s*([mf])", s)
    if m:
        proc = "tha" if m.group(1) == "oa" else m.group(1)
        hits = [p for p in patients
                if p["procedure"] == proc and p["age"] == m.group(2) and p["sex"] == m.group(3).upper()]
        if len(hits) == 1:
            return hits[0], None, "番号なし。術式・年齢・性別から推定"
        if hits:
            return None, None, "候補が複数: " + ", ".join("#" + p["number"] for p in hits)
    return None, None, None


def classify(stem, table):
    s = stem.lower()
    for name, pattern in table:
        if re.search(pattern, s):
            return name
    return None


def build_index(root, patients, skip):
    by_patient = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    notes = defaultdict(set)
    unmatched = []
    for path in sorted(root.rglob("*")):
        if path.suffix.lower() not in IMAGE_EXTS or path.name.startswith("."):
            continue
        rel = path.relative_to(root)
        if any(part in skip or part.endswith(".tif_files") or part.startswith(".") for part in rel.parts):
            continue
        if NOT_HUMAN.search(str(rel)):
            continue
        stem = path.stem
        tissue = classify(stem, TISSUES)
        if tissue is None:
            continue  # 組織名のないファイル（マウス・細胞・Western など）は対象外
        patient, code, note = find_patient(stem, patients)
        # 名前に染色がなければフォルダ名から（例: human_RXFP1/55THA81m-cap100）
        stain = classify(stem, STAINS) or classify(str(rel.parent), STAINS) or "不明"
        entry = {"path": str(rel), "type": path.suffix.lower().lstrip(".")}
        if EXPORT_MARKERS.search(stem):
            entry["export"] = True
        if FLUORESCENCE.search(stem):
            entry["fluorescence"] = True
        if patient is None:
            unmatched.append({**entry, "tissue": tissue, "stain": stain, **({"note": note} if note else {})})
            continue
        if note:
            entry["note"] = note
        by_patient[patient["number"]][tissue][stain].append(entry)
        if code:
            notes[patient["number"]].add(code)
    # PPIB（RNAscope のポジコン）がある、または名前に ish があるフォルダは in situ hybridization とみなす
    ish_folders = set()
    for tissues in by_patient.values():
        for stains in tissues.values():
            for stain, files in stains.items():
                for f in files:
                    folder = str(Path(f["path"]).parent)
                    if stain == "PPIB" or re.search(r"ish", f["path"], re.I):
                        ish_folders.add(folder)
    for tissues in by_patient.values():
        for stains in tissues.values():
            for files in stains.values():
                for f in files:
                    if str(Path(f["path"]).parent) in ish_folders:
                        f["ish"] = True
    return by_patient, notes, unmatched


def summarize(files):
    """DAB 免疫染色（明視野）の CZI があるか。蛍光・ISH しかない場合はそう書く。"""
    brightfield = [f for f in files if not f.get("fluorescence") and not f.get("ish")]
    if not brightfield:
        return "ISH のみ" if all(f.get("ish") for f in files) else "蛍光・ISH のみ"
    return "czi" if any(f["type"] == "czi" for f in brightfield) else "jpg/tif のみ"


def main():
    parser = argparse.ArgumentParser(description="顕微鏡画像を患者・組織・染色ごとに一覧にする")
    parser.add_argument("--csv", type=Path, required=True, help="患者リスト（relaxin_clinical の CSV）")
    parser.add_argument("--root", type=Path, default=Path.home() / "Library/CloudStorage/OneDrive-Personal/Desktop",
                        help="探すフォルダ（既定: OneDrive の Desktop）。JSON のパスはここからの相対パス")
    parser.add_argument("-o", "--output", type=Path, required=True, help="出力する JSON")
    parser.add_argument("--skip", nargs="*", default=["20261003産業医大学会", "relaxin_clinical"],
                        help="対象外にするフォルダ名")
    args = parser.parse_args()

    patients = load_patients(args.csv)
    by_patient, codes, unmatched = build_index(args.root.expanduser(), patients, set(args.skip))

    result = {
        "root": str(args.root),
        "generated": date.today().isoformat(),
        "note": ("パスは root からの相対パス。type が jpg/tif だけのものも、元の CZI が別の場所にある可能性がある。"
                 "fluorescence = ApoTome の蛍光、ish = in situ hybridization（PPIB のあるフォルダ）、export = ZEN の書き出し。"
                 "rxfp_coverage は DAB 免疫染色（蛍光・ISH を除く）の有無"),
        "patients": {},
        "rxfp_coverage": {},
        "unmatched": unmatched,
    }
    order = sorted(by_patient, key=int)
    for number in order:
        p = next(p for p in patients if p["number"] == number)
        tissues = by_patient[number]
        result["patients"][f"#{number}"] = {
            "age": p["age"], "sex": p["sex"], "procedure": p["procedure"],
            "file_codes": sorted(codes[number]),
            "list_rxfp": {k: v for k, v in p["list_rxfp"].items() if v},
            "images": {t: dict(sorted(s.items())) for t, s in sorted(tissues.items())},
        }
        # RXFP1 免疫染色がどの組織にあるか（リストの判定と並べる）
        coverage = {}
        for tissue in ("cam", "cap", "lab", "cot"):
            files = tissues.get(tissue, {}).get("RXFP1", [])
            coverage[tissue] = summarize(files) if files else "なし"
        result["rxfp_coverage"][f"#{number} {p['procedure']} {p['age']}{p['sex']}"] = coverage

    text = json.dumps(result, ensure_ascii=False, indent=2)
    args.output.write_text(text + "\n", encoding="utf-8")
    n_files = sum(len(f) for t in by_patient.values() for s in t.values() for f in s.values())
    print(f"{args.output} を書き出しました（患者 {len(by_patient)} 人、画像 {n_files} 件、患者不明 {len(unmatched)} 件）",
          file=sys.stderr)


if __name__ == "__main__":
    main()
