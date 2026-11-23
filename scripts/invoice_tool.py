#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""电脑硬盘发票归类整理 —— 发票识别、重命名、汇总工具。

子命令：
  scan    扫描目录，抽取发票要素并生成清单 manifest.json（不改动任何文件）
  apply   依据清单重命名 / 移动 / 复制发票（默认只预览，加 --execute 才落盘）
  report  依据清单生成 Excel 汇总工作簿

用法示例：
  python invoice_tool.py scan   --src "D:\\发票" --manifest ".\\_发票整理归档\\manifest.json"
  python invoice_tool.py apply  --manifest ".\\_发票整理归档\\manifest.json"
  python invoice_tool.py apply  --manifest "..." --execute --mode rename
  python invoice_tool.py report --manifest "..." --xlsx ".\\发票汇总.xlsx"
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import shutil
import sys
import unicodedata
import zipfile
from datetime import date, datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SKILL_DIR = Path(__file__).resolve().parent.parent
RULES_PATH = SKILL_DIR / "references" / "category-rules.json"
ARCHIVE_DIRNAME = "_发票整理归档"

INVOICE_EXTS = {
    ".pdf", ".ofd", ".xml", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp",
}
SKIP_DIRS = {"_发票整理归档", ".git", "node_modules", "__MACOSX", ".idea", ".vscode"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}

RE_NO = re.compile(r"发\s*票\s*号\s*码\s*[:：]\s*([0-9]{8,24})")
RE_CODE = re.compile(r"发\s*票\s*代\s*码\s*[:：]\s*([0-9]{10,12})")
RE_DATE = re.compile(
    r"开\s*票\s*日\s*期\s*[:：]\s*(\d{4})\s*[年\-/.]\s*(\d{1,2})\s*[月\-/.]\s*(\d{1,2})"
)
RE_DATE_ALT = re.compile(
    r"(?:乘\s*车|出\s*发|航\s*班|起\s*飞)\s*日\s*期\s*[:：]?\s*(\d{4})\s*[年\-/.]\s*(\d{1,2})\s*[月\-/.]\s*(\d{1,2})"
)
RE_SMALL = re.compile(r"小\s*写\s*[）)]?\s*[:：]?\s*[¥￥]?\s*(-?[\d,]+\.?\d{0,2})")
RE_TAX_TOTAL = re.compile(r"价\s*税\s*合\s*计[\s\S]{0,120}?[¥￥]\s*(-?[\d,]+\.?\d{0,2})")
RE_SUM = re.compile(r"合\s*计\s*[¥￥]\s*([\d,]+\.\d{2})\s*[¥￥]\s*([\d,]+\.\d{2})")
RE_NAME = re.compile(r"(?<!项)名\s*称\s*[:：]\s*([^\n]{2,60})")
RE_TAXID = re.compile(r"(?:纳税人识别号|统一社会信用代码)[^0-9A-Z]{0,8}([0-9A-Z]{15,20})")
RE_ITEM = re.compile(r"\*[^*\n]{1,24}\*[^\n]{0,60}")
RE_RATE = re.compile(r"(\d{1,2}(?:\.\d+)?)\s*%")

INVALID_CHARS = '<>:"/\\|?*'


# --------------------------------------------------------------------------- 工具
def norm(s: str) -> str:
    """全角转半角、压缩空白，便于关键词匹配。"""
    return re.sub(r"[ \t\u3000]+", "", unicodedata.normalize("NFKC", s or ""))


def to_float(s):
    if s is None:
        return None
    try:
        return round(float(str(s).replace(",", "").replace("¥", "").replace("￥", "").strip()), 2)
    except Exception:
        return None


def to_iso_date(y, m, d):
    try:
        return date(int(y), int(m), int(d)).isoformat()
    except Exception:
        return None


def is_red(text: str) -> bool:
    return ("红字" in text) or ("负数" in text) or ("冲红" in text)


# --------------------------------------------------------------------- 文本抽取
def pdf_pages(path: Path):
    """返回逐页文本列表。"""
    try:
        import pymupdf
    except ImportError:
        return []
    try:
        with pymupdf.open(path) as doc:
            return [pg.get_text("text") or "" for pg in doc]
    except Exception:
        return []


def ofd_variants(path: Path):
    """OFD 是 zip 容器，正文在 XML 的 <TextCode> 中。返回不同拼接方式的候选文本。"""
    chunks = []
    try:
        with zipfile.ZipFile(path) as zf:
            for name in zf.namelist():
                if not name.lower().endswith(".xml"):
                    continue
                try:
                    raw = zf.read(name).decode("utf-8", "ignore")
                except Exception:
                    continue
                for m in re.finditer(r"<TextCode[^>]*>(.*?)</TextCode>", raw, re.S):
                    chunks.append(unescape_xml(m.group(1)))
    except Exception:
        return []
    if not chunks:
        return []
    return ["".join(chunks), "\n".join(chunks)]


def unescape_xml(s: str) -> str:
    return (
        s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
        .replace("&apos;", "'").replace("&amp;", "&")
    )


TAG_ALIASES = {
    "invoice_no": ("invoicenumber", "invoiceno", "发票号码"),
    "invoice_code": ("invoicecode", "发票代码"),
    "date": ("issuedate", "issuetime", "invoicedate", "billingdate", "开票日期"),
    "total": ("totaltaxincludedamount", "totalamountwithtax", "amountwithtax", "totaltaxincluded", "价税合计"),
    "amount": ("totalamount", "taxexclusiveamount", "amountwithouttax", "amountexcludetax", "合计金额"),
    "tax": ("totaltax", "taxamount", "合计税额"),
    "buyer": ("buyername", "purchasername", "购买方名称"),
    "seller": ("sellername", "销售方名称"),
    "item": ("itemname", "goodsname", "项目名称"),
}


def xml_fields(raw: str) -> dict:
    """数值发票 XML：按标签别名抓字段。"""
    out = {}
    for m in re.finditer(r"<([A-Za-z0-9_:]+)>([^<>]{1,120})</\1>", raw):
        tag = re.sub(r"^[^:]+:", "", m.group(1)).lower().replace("_", "").replace("-", "")
        val = m.group(2).strip()
        for field, aliases in TAG_ALIASES.items():
            if field in out:
                continue
            if tag in aliases or any(tag.endswith(a) for a in aliases):
                out[field] = val
    return out


def read_candidates(path: Path):
    """返回候选文本列表（用于择优选解析结果）。"""
    ext = path.suffix.lower()
    if ext == ".pdf":
        pages = pdf_pages(path)
        return ["\n".join(pages)] if pages else []
    if ext == ".ofd":
        return ofd_variants(path)
    if ext == ".xml":
        try:
            raw = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            return []
        stripped = re.sub(r"<[^>]+>", " ", raw)
        return [raw, stripped]
    return []


# ----------------------------------------------------------------------- 解析
def parse_text(text: str) -> tuple[dict, int]:
    data = {
        "invoice_no": None, "invoice_code": None, "date": None,
        "total": None, "amount": None, "tax": None, "rate": None,
        "buyer": None, "seller": None, "buyer_tax_id": None,
        "item": None, "invoice_type": None,
    }
    if not text:
        return data, 0

    m = RE_NO.search(text)
    if m:
        data["invoice_no"] = m.group(1)
    m = RE_CODE.search(text)
    if m:
        data["invoice_code"] = m.group(1)

    for pattern in (RE_DATE, RE_DATE_ALT):
        m = pattern.search(text)
        if m:
            data["date"] = to_iso_date(*m.groups())
            if data["date"]:
                break

    m = RE_SMALL.search(text) or RE_TAX_TOTAL.search(text)
    if m:
        data["total"] = to_float(m.group(1))

    m = RE_SUM.search(text)
    if m:
        data["amount"] = to_float(m.group(1))
        data["tax"] = to_float(m.group(2))
        if data["total"] is None and data["amount"] is not None and data["tax"] is not None:
            data["total"] = round(data["amount"] + data["tax"], 2)

    rates = sorted({r for r in RE_RATE.findall(text)}, key=lambda x: float(x))
    if rates:
        data["rate"] = "、".join(f"{r}%" for r in rates[:3])

    # 金额为 0 且有一方为空：可能只是税率行，回退取票面最大金额
    if data["total"] in (0.0, None) and data["amount"] in (None, 0.0):
        nums = [to_float(x) for x in re.findall(r"[¥￥]\s*(-?[\d,]+\.\d{2})", text)]
        nums = [n for n in nums if n]
        if nums:
            data["total"] = max(nums)

    names = [n.strip() for n in RE_NAME.findall(text)]
    names = [n for n in names if n and n not in ("名称",)]
    if names:
        data["buyer"] = names[0]
    if len(names) > 1:
        data["seller"] = names[1]

    m = RE_TAXID.search(text)
    if m:
        data["buyer_tax_id"] = m.group(1)

    items = RE_ITEM.findall(text)
    if items:
        cleaned = []
        for raw in items:
            s = re.sub(r"(?:\s+\d[\d.,%]*)+\s*$", "", raw.strip())
            cleaned.append(s.strip() or raw.strip())
        data["item"] = "；".join(dict.fromkeys(cleaned))[:200]

    data["invoice_type"] = detect_type(text)

    score = 0
    score += 3 if data["invoice_no"] else 0
    score += 3 if data["date"] else 0
    score += 3 if data["total"] is not None else 0
    score += 1 if data["buyer"] else 0
    score += 1 if data["seller"] else 0
    score += 1 if data["item"] else 0
    return data, score


def detect_type(text: str) -> str:
    t = norm(text)
    if "铁路" in t and ("客票" in t or "车次" in t or "乘车" in t):
        return "铁路电子客票"
    if "航空运输电子客票行程单" in t or ("民航发展基金" in t and "行程单" in t):
        return "航空运输电子客票行程单"
    if "机动车销售统一发票" in t:
        return "机动车销售统一发票"
    if "二手车销售统一发票" in t:
        return "二手车销售统一发票"
    if "电子发票（增值税专用发票）" in t or "电子发票(增值税专用发票)" in t:
        return "数电票-专用发票"
    if "电子发票（普通发票）" in t or "电子发票(普通发票)" in t:
        return "数电票-普通发票"
    if "通行费" in t and "发票" in t:
        return "通行费电子发票"
    if "增值税电子专用发票" in t:
        return "增值税电子专用发票"
    if "增值税电子普通发票" in t:
        return "增值税电子普通发票"
    if "增值税专用发票" in t:
        return "增值税专用发票"
    if "增值税普通发票" in t:
        return "增值税普通发票"
    if "定额发票" in t:
        return "定额发票"
    return "未知"


def parse_file(path: Path) -> list[dict]:
    """解析单个文件，返回记录列表（合并扫描件可能返回多条）。"""
    ext = path.suffix.lower()
    recs = []

    if ext == ".pdf":
        pages = pdf_pages(path)
        if not pages:
            return []
        page_parses = [parse_text(p) for p in pages]
        distinct = {d["invoice_no"] for d, _ in page_parses if d["invoice_no"]}
        if len(distinct) > 1:
            for idx, (d, sc) in enumerate(page_parses, start=1):
                if sc < 4:
                    continue
                d["page"] = idx
                d["score"] = sc
                d["page_count"] = len(pages)
                recs.append(d)
            return recs
        whole = "\n".join(pages)
        d, sc = parse_text(whole)
        d["score"] = sc
        d["page_count"] = len(pages)
        return [d]

    if ext == ".xml":
        raw = path.read_text(encoding="utf-8", errors="ignore")
        fields = xml_fields(raw)
        d, sc = parse_text(raw)
        mapping = {"invoice_no": "invoice_no", "invoice_code": "invoice_code",
                   "amount": "amount", "tax": "tax", "buyer": "buyer", "seller": "seller"}
        for key, target in mapping.items():
            if fields.get(key) and not d.get(target):
                d[target] = fields[key]
        if fields.get("total"):
            d["total"] = to_float(fields["total"])
        if fields.get("date"):
            dm = re.search(r"(\d{4})[-/年]?(\d{1,2})[-/月]?(\d{1,2})", fields["date"])
            if dm:
                d["date"] = to_iso_date(*dm.groups())
        if fields.get("item"):
            d["item"] = fields["item"][:200]
        for k in ("total", "amount", "tax"):
            d[k] = to_float(d[k])
        if d["total"] is None and d["amount"] is not None and d["tax"] is not None:
            d["total"] = round(d["amount"] + d["tax"], 2)
        d["score"] = sc + 3
        d["page_count"] = 1
        return [d]

    candidates = read_candidates(path)
    if candidates:
        parsed = [parse_text(c) for c in candidates]
        d, sc = max(parsed, key=lambda x: x[1])
        d["score"] = sc
        d["page_count"] = 1
        if sc >= 4:
            return [d]
    return []


# ------------------------------------------------------------------ 分类与命名
def classify(rec: dict, rules: dict, extra: dict):
    item_txt = rec.get("item") or ""
    hay = norm(" ".join(x for x in [item_txt, rec.get("seller") or "", rec.get("invoice_type") or ""] if x))

    for kw, cat in (extra.get("category_overrides") or {}).items():
        if norm(kw) and norm(kw) in hay:
            return cat, f"人工规则：{kw}"

    best = None
    for rule in rules.get("rules", []):
        if not rule.get("keywords"):
            continue
        for kw in rule["keywords"]:
            if norm(kw) in hay:
                if best is None or rule["priority"] > best[0]["priority"]:
                    best = (rule, kw)
                break
    if best:
        return best[0]["category"], f"关键词：{best[1]}"
    return "其他", "未匹配规则"


def sanitize(name: str) -> str:
    s = "".join("_" if c in INVALID_CHARS else c for c in name)
    s = re.sub(r"\s+", "", s).strip(" .")
    return s[:110]


def make_name(rec: dict) -> str:
    ymd = (rec.get("date") or "").replace("-", "") or "日期待补"
    cat = rec.get("category") or "未分类"
    if rec.get("total") is None:
        amt = "金额待补"
    elif rec["total"] < 0:
        amt = f"红冲{abs(rec['total']):.2f}"
    else:
        amt = f"{rec['total']:.2f}"
    return sanitize(f"{ymd}-{cat}-{amt}")


def unique_name(dirpath: Path, base: str, ext: str, taken: set) -> str:
    candidate = f"{base}{ext}"
    i = 2
    while candidate.lower() in taken or (dirpath / candidate).exists():
        candidate = f"{base}_{i}{ext}"
        i += 1
    taken.add(candidate.lower())
    return candidate


# ------------------------------------------------------------------------ scan
def iter_files(src: Path, recursive: bool):
    for p in sorted(src.rglob("*") if recursive else src.glob("*")):
        if not p.is_file():
            continue
        if any(part in SKIP_DIRS or part.startswith(".") for part in p.relative_to(src).parts[:-1]):
            continue
        if p.suffix.lower() in INVOICE_EXTS:
            yield p


def load_overrides(path, src: Path) -> dict:
    if not path:
        return {}
    p = Path(path)
    if not p.exists():
        return {}
    data = json.loads(p.read_text(encoding="utf-8"))
    out = {}
    for k, v in data.items():
        out[str(Path(k)).lower()] = v
        out[Path(k).name.lower()] = v
    return out


def cmd_scan(args):
    src = Path(args.src).expanduser().resolve()
    if not src.is_dir():
        sys.exit(f"目录不存在：{src}")
    rules = json.loads(Path(args.rules).read_text(encoding="utf-8")) if args.rules else \
        json.loads(RULES_PATH.read_text(encoding="utf-8"))
    extra = {}
    if args.profile:
        extra = json.loads(Path(args.profile).read_text(encoding="utf-8"))
    if args.company:
        extra["company"] = args.company
    if args.tax_id:
        extra["tax_id"] = args.tax_id
    threshold = args.fixed_asset_threshold or extra.get("fixed_asset_threshold") or \
        rules.get("fixed_asset_threshold", 5000)
    overrides = load_overrides(args.overrides, src)

    invoices, skipped = [], []
    taken = set()

    for path in iter_files(src, not args.no_recursive):
        rel = path.relative_to(src).as_posix()
        ext = path.suffix.lower()
        try:
            parsed = parse_file(path)
        except Exception as exc:  # noqa: BLE001
            skipped.append({"file": rel, "reason": f"解析异常：{exc}"})
            continue

        if not parsed:
            if ext in IMAGE_EXTS:
                parsed = [{"invoice_no": None, "invoice_code": None, "date": None, "total": None,
                           "amount": None, "tax": None, "rate": None, "buyer": None,
                           "seller": None, "buyer_tax_id": None, "item": None,
                           "invoice_type": "图片待识别", "score": 0, "page_count": 1}]
                skipped.append({"file": rel, "reason": "图片发票已列入清单，要素需人工/视觉读取后写入 overrides 再重跑 scan"})
            else:
                skipped.append({"file": rel, "reason": "未识别到发票要素，非发票或版式不支持"})
                continue

        for item in parsed:
            ov = overrides.get(rel.lower()) or overrides.get(path.name.lower()) or {}
            for key in ("invoice_no", "invoice_code", "date", "total", "amount", "tax",
                        "buyer", "seller", "item", "invoice_type", "category", "note"):
                if ov.get(key) not in (None, ""):
                    item[key] = ov[key]
            if item.get("total") is not None:
                item["total"] = to_float(item["total"])

            rec = {
                "src": src.as_posix(),
                "rel": rel,
                "filename": path.name,
                "ext": ext,
                "page": item.get("page"),
                "page_count": item.get("page_count", 1),
                "invoice_no": item.get("invoice_no"),
                "invoice_code": item.get("invoice_code"),
                "date": item.get("date"),
                "total": item.get("total"),
                "amount": item.get("amount"),
                "tax": item.get("tax"),
                "rate": item.get("rate"),
                "buyer": item.get("buyer"),
                "seller": item.get("seller"),
                "buyer_tax_id": item.get("buyer_tax_id"),
                "item": item.get("item"),
                "invoice_type": item.get("invoice_type"),
                "note": ov.get("note", ""),
                "score": item.get("score", 0),
            }

            if ov.get("category"):
                rec["category"], rec["category_source"] = ov["category"], "人工指定"
            else:
                rec["category"], rec["category_source"] = classify(rec, rules, extra)

            rec["flags"] = build_flags(rec, extra, threshold, args.year)
            if rec.get("date") and rec.get("total") is not None:
                rec["new_name"] = make_name(rec) + ext
            else:
                rec["new_name"] = path.name
                rec["flags"].append("字段不全，暂保留原文件名")
            invoices.append(rec)

    mark_duplicates(invoices)
    assign_unique_names(invoices)

    manifest = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_dir": src.as_posix(),
        "recursive": not args.no_recursive,
        "year_filter": args.year,
        "company": extra.get("company"),
        "tax_id": extra.get("tax_id"),
        "fixed_asset_threshold": threshold,
        "count": len(invoices),
        "total_amount": round(sum(i["total"] or 0 for i in invoices), 2),
        "invoices": invoices,
        "skipped": skipped,
        "overrides_used": bool(overrides),
    }

    out = Path(args.manifest) if args.manifest else src / ARCHIVE_DIRNAME / "manifest.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"扫描目录：{src}")
    print(f"识别发票：{len(invoices)} 张，合计 {manifest['total_amount']:,.2f} 元")
    print(f"未识别/跳过：{len(skipped)} 个文件")
    print(f"清单已写入：{out}")
    warn = [i for i in invoices if i["flags"]]
    if warn:
        print(f"需人工复核：{len(warn)} 张（见清单 flags 字段）")
    return manifest["count"]


def build_flags(rec, extra, threshold, year) -> list[str]:
    flags = []
    if not rec.get("invoice_no") and rec.get("invoice_type") not in ("图片待识别",):
        flags.append("缺发票号码")
    if not rec.get("date"):
        flags.append("缺开票日期")
    if rec.get("total") is None:
        flags.append("缺金额")
    elif rec["total"] == 0:
        flags.append("金额为0")
    elif rec["total"] < 0:
        flags.append("负数红冲票")

    if rec.get("invoice_type") == "图片待识别":
        flags.append("图片待人工读取")

    company = extra.get("company")
    if company and rec.get("buyer"):
        if norm(company) not in norm(rec["buyer"]) and norm(rec["buyer"]) not in norm(company):
            flags.append("抬头与公司不符")
    tax_id = extra.get("tax_id")
    if tax_id and rec.get("buyer_tax_id") and rec["buyer_tax_id"].upper() != tax_id.upper():
        flags.append("购方税号不符")

    if rec.get("category") == "固定资产" and rec.get("total") and rec["total"] < threshold:
        flags.append(f"低于固定资产标准{threshold}元，建议费用化")
    if rec.get("category") in (extra.get("ambiguous") or []):
        flags.append("分类待确认")

    if year and rec.get("date") and rec["date"][:4] != str(year):
        flags.append("归属年度不符")
    if rec.get("invoice_type") == "未知":
        flags.append("票种未识别")
    return flags


def mark_duplicates(invoices):
    groups = {}
    for i in invoices:
        key = i.get("invoice_no") or f"{i.get('date')}|{i.get('total')}|{i.get('seller')}"
        groups.setdefault(key, []).append(i)
    for key, items in groups.items():
        if len(items) < 2:
            continue
        for dup in items:
            dup["flags"].append(f"疑似重复（{len(items)}张同号）")


def assign_unique_names(invoices):
    """按目标目录去重；同一文件拆页时用页码区分。"""
    taken = {}
    for i in invoices:
        base = Path(i["new_name"]).stem
        d = i["src"] if i.get("move_to") is None else i["move_to"]
        bucket = taken.setdefault(d, set())
        name = Path(i["new_name"]).name
        if name.lower() in bucket:
            i["new_name"] = unique_name(Path(d), base, i["ext"], bucket)
        else:
            bucket.add(name.lower())


# ----------------------------------------------------------------------- apply
def cmd_apply(args):
    man = json.loads(Path(args.manifest).read_text(encoding="utf-8"))
    src = Path(man["source_dir"])
    out_dir = Path(args.out).resolve() if args.out else src
    mode = args.mode
    invoices = man["invoices"]

    plan = []
    for rec in invoices:
        old = src / rec["rel"]
        if not old.exists():
            plan.append((rec, None, "源文件不存在，跳过"))
            continue
        if rec.get("page") and old.suffix.lower() == ".pdf" and rec["page_count"] > 1:
            target = out_dir / rec["new_name"]
            plan.append((rec, target, "拆页另存"))
            continue
        target = out_dir / rec["new_name"]
        plan.append((rec, target, mode))

    print(f"整理模式：{mode}　目标目录：{out_dir}")
    for rec, target, act in plan:
        print(f"  {rec['rel']}  ->  {target.name if target else '（跳过）'}   [{act}]")

    if not args.execute:
        print("\n以上为预览，未改动任何文件。确认后加 --execute 执行。")
        return

    out_dir.mkdir(parents=True, exist_ok=True)
    archive = Path(args.archive) if args.archive else out_dir / ARCHIVE_DIRNAME
    archive.mkdir(parents=True, exist_ok=True)
    rollback_path = archive / f"重命名记录_{datetime.now():%Y%m%d%H%M%S}.csv"
    taken = set()
    applied = 0
    split_src = []

    with rollback_path.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["原路径", "新路径", "发票号码", "类型", "金额", "模式"])
        for rec, target, act in plan:
            old = src / rec["rel"]
            if target is None or not old.exists():
                continue
            final = target
            if final.resolve() == old.resolve():
                rec["final_path"] = old.as_posix()
                continue
            if final.exists() and final.name.lower() in taken:
                final = final.with_name(unique_name(out_dir, final.stem, final.suffix, taken))
            elif final.exists() and args.mode != "copy":
                final = final.with_name(unique_name(out_dir, final.stem, final.suffix, taken))
            taken.add(final.name.lower())

            try:
                if rec.get("page") and old.suffix.lower() == ".pdf" and rec["page_count"] > 1:
                    import pymupdf
                    with pymupdf.open(old) as doc:
                        single = pymupdf.open()
                        single.insert_pdf(doc, from_page=rec["page"] - 1, to_page=rec["page"] - 1)
                        single.save(final)
                        single.close()
                    if old.as_posix() not in split_src:
                        split_src.append(old.as_posix())
                elif args.mode == "copy":
                    shutil.copy2(old, final)
                else:
                    shutil.move(str(old), str(final))
            except Exception as exc:  # noqa: BLE001
                print(f"  失败：{rec['rel']} -> {final.name}：{exc}")
                continue

            rec["final_path"] = final.as_posix()
            w.writerow([old.as_posix(), final.as_posix(), rec.get("invoice_no"),
                        rec.get("category"), rec.get("total"), args.mode])
            applied += 1

    man["organized_at"] = datetime.now().isoformat(timespec="seconds")
    man["mode"] = args.mode
    man["out_dir"] = out_dir.as_posix()
    Path(args.manifest).write_text(json.dumps(man, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n已处理 {applied} 张，回滚记录：{rollback_path}")
    if split_src:
        print("拆页已生成单票文件，以下原合并文件保留未动（核对无误后再手工归档）：")
        for s in split_src:
            print(f"  {s}")


# ---------------------------------------------------------------------- report
def deductibility(rec: dict) -> str:
    """专票才可能抵扣；用于交际应酬、集体福利的进项不得抵扣。"""
    t = rec.get("invoice_type") or ""
    if not t.endswith("专用发票"):
        return "不可抵扣"
    if rec.get("category") in ("业务招待费", "职工福利费"):
        return "不可抵扣(用途)"
    return "可抵扣"


COLUMNS = [
    ("序号", 6), ("开票日期", 12), ("类型", 14), ("金额(价税合计)", 14), ("不含税金额", 13),
    ("税额", 11), ("发票号码", 22), ("票种", 20), ("销售方", 30), ("购买方", 28),
    ("抬头校验", 14), ("重复校验", 14), ("可抵扣", 10), ("项目/服务名称", 34),
    ("文件名", 38), ("原始路径", 40), ("复核提示", 34),
]


def cmd_report(args):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    man = json.loads(Path(args.manifest).read_text(encoding="utf-8"))
    invoices = man["invoices"]
    rules = json.loads(RULES_PATH.read_text(encoding="utf-8"))
    notes = {r["category"]: r.get("note", "") for r in rules["rules"]}
    hints = rules.get("deduction_hint", {})

    wb = Workbook()
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_header(ws, row=1):
        for cell in ws[row]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        ws.freeze_panes = ws.cell(row=row + 1, column=1)
        ws.auto_filter.ref = ws.dimensions

    # ---- 明细
    ws = wb.active
    ws.title = "发票明细"
    ws.append([c[0] for c in COLUMNS])
    for i, (_, width) in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    for idx, rec in enumerate(invoices, start=1):
        dup = [f for f in rec["flags"] if f.startswith("疑似重复")]
        head = "通过" if not any(f in ("抬头与公司不符", "购方税号不符") for f in rec["flags"]) else "异常"
        d = None
        if rec.get("date"):
            try:
                d = date.fromisoformat(rec["date"])
            except Exception:
                d = None
        ws.append([
            idx, d or rec.get("date") or "待补", rec.get("category"),
            rec.get("total"), rec.get("amount"), rec.get("tax"),
            rec.get("invoice_no"), rec.get("invoice_type") or "未知",
            rec.get("seller"), rec.get("buyer"),
            head if rec.get("buyer") else "未识别",
            "重复" if dup else "唯一",
            deductibility(rec),
            rec.get("item"), rec.get("new_name"),
            rec.get("final_path") or rec.get("rel"),
            "；".join(rec["flags"]),
        ])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(COLUMNS)):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        row[1].number_format = "yyyy-mm-dd"
        for c in (row[3], row[4], row[5]):
            c.number_format = "#,##0.00"
    style_header(ws)

    # ---- 分类汇总
    ws2 = wb.create_sheet("分类汇总")
    ws2.append(["类型", "张数", "金额(价税合计)", "占比", "税前扣除口径", "科目说明/合规要点"])
    widths2 = [16, 8, 16, 10, 42, 56]
    for i, w in enumerate(widths2, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    agg = {}
    for rec in invoices:
        cat = rec.get("category") or "未分类"
        a = agg.setdefault(cat, {"n": 0, "total": 0.0})
        a["n"] += 1
        a["total"] += rec.get("total") or 0
    grand = sum(v["total"] for v in agg.values())

    for cat, v in sorted(agg.items(), key=lambda x: -x[1]["total"]):
        ws2.append([cat, v["n"], round(v["total"], 2),
                    (v["total"] / grand) if grand else 0,
                    hints.get(cat, next((r["deduction"] for r in rules["rules"] if r["category"] == cat), "据实判断")),
                    notes.get(cat, "")])
    ws2.append(["合计", sum(v["n"] for v in agg.values()), round(grand, 2), 1, "", ""])

    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, max_col=6):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        row[2].number_format = "#,##0.00"
        row[3].number_format = "0.0%"
    ws2.cell(row=ws2.max_row, column=1).font = Font(bold=True)
    style_header(ws2)

    # ---- 月度汇总
    ws3 = wb.create_sheet("月度汇总")
    ws3.append(["归属月份", "张数", "金额(价税合计)"])
    for i, w in enumerate([14, 10, 16], start=1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    monthly = {}
    for rec in invoices:
        month = (rec.get("date") or "")[:7] or "日期待补"
        m = monthly.setdefault(month, {"n": 0, "total": 0.0})
        m["n"] += 1
        m["total"] += rec.get("total") or 0
    for month, v in sorted(monthly.items()):
        ws3.append([month, v["n"], round(v["total"], 2)])
    ws3.append(["合计", sum(v["n"] for v in monthly.values()),
                round(sum(v["total"] for v in monthly.values()), 2)])
    for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row, max_col=3):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center")
        row[2].number_format = "#,##0.00"
    style_header(ws3)

    # ---- 异常待处理
    ws4 = wb.create_sheet("异常待处理")
    ws4.append(["序号", "文件名", "涉及发票", "问题", "建议处理"])
    for i, w in enumerate([6, 40, 22, 34, 52], start=1):
        ws4.column_dimensions[get_column_letter(i)].width = w
    advice = {
        "缺发票号码": "核对票面或到开票平台/税务数字账户重新下载",
        "缺开票日期": "人工补录日期后重新生成台账",
        "缺金额": "核对金额小写栏，或按不含税金额+税额手工计算",
        "图片待人工读取": "用图像识别读取票面，把字段写入 overrides 后重跑 scan",
        "抬头与公司不符": "个人抬头或他方抬头不得报销；联系商家重开",
        "购方税号不符": "税号错误须重开；核对是否为集团内其他主体",
        "疑似重复": "同一发票号只报销一次，其余剔除并注明原因",
        "分类待确认": "按业务实质确认科目，必要时与财务沟通",
        "负数红冲票": "与原蓝票配对，冲减当期费用",
        "低于固定资产标准": "改按低值易耗品/办公费一次性费用化",
        "归属年度不符": "跨年票据按权责发生制或汇算清缴口径处理",
        "票种未识别": "人工确认票种，避免误判可抵扣性",
        "金额为0": "核对是否为误下载或票面异常",
    }
    n = 0
    for rec in invoices:
        for f in rec["flags"]:
            n += 1
            ws4.append([n, rec.get("new_name"), rec.get("invoice_no") or "—", f,
                        advice.get(f.split("（")[0], "人工复核")])
    for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row, max_col=5):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    style_header(ws4)

    # ---- 使用说明
    ws5 = wb.create_sheet("说明")
    ws5.column_dimensions["A"].width = 26
    ws5.column_dimensions["B"].width = 96
    rows = [
        ("生成时间", man.get("generated_at", "")),
        ("扫描目录", man.get("source_dir", "")),
        ("整理模式", man.get("mode", "未执行重命名")),
        ("发票张数", man.get("count", 0)),
        ("金额合计", f"{man.get('total_amount', 0):,.2f}"),
        ("公司抬头", man.get("company") or "未设置（未做抬头校验）"),
        ("购方税号", man.get("tax_id") or "未设置"),
        ("命名规则", "开票年月日-类型-金额（价税合计，两位小数），重名自动加 _2"),
        ("金额口径", "金额=价税合计；不含税金额与税额用于账务入账与进项抵扣"),
        ("类型口径", "按企业会计准则费用性质归集，参照企业所得税税前扣除口径提示限额"),
        ("税前扣除提示", "限额类费用（招待/福利/教育经费/广告）需全年累计测算，单张发票不构成超限"),
        ("电子归档", "电子发票应保存原始电子文件（PDF/OFD/XML），保存期限不少于10年"),
        ("注意", "本表由发票票面信息自动抽取，入账前须与票面原件复核"),
    ]
    ws5.append(["项目", "内容"])
    for k, v in rows:
        ws5.append([k, v])
    for row in ws5.iter_rows(min_row=1, max_row=ws5.max_row, max_col=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    style_header(ws5)

    xlsx = Path(args.xlsx) if args.xlsx else \
        Path(man.get("out_dir") or man["source_dir"]) / ARCHIVE_DIRNAME / \
        f"发票汇总_{datetime.now():%Y%m%d%H%M}.xlsx"
    xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx)
    print(f"汇总表已生成：{xlsx}")
    print(f"明细 {len(invoices)} 行，类型 {len(agg)} 类，异常 {n} 条")
    return xlsx.as_posix()


# ------------------------------------------------------------------------ main
def main():
    parser = argparse.ArgumentParser(description="电脑硬盘发票归类整理")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("scan", help="扫描目录并生成清单")
    p.add_argument("--src", required=True)
    p.add_argument("--manifest")
    p.add_argument("--overrides", help="人工补录字段的 JSON（图片发票用）")
    p.add_argument("--profile", help="公司抬头/税号/规则覆盖的 JSON")
    p.add_argument("--rules")
    p.add_argument("--company")
    p.add_argument("--tax-id", dest="tax_id")
    p.add_argument("--fixed-asset-threshold", dest="fixed_asset_threshold", type=float)
    p.add_argument("--year", type=int, help="归属年度，用于标记跨年票据")
    p.add_argument("--no-recursive", action="store_true")
    p.set_defaults(func=cmd_scan)

    p = sub.add_parser("apply", help="执行重命名/移动/复制")
    p.add_argument("--manifest", required=True)
    p.add_argument("--mode", choices=["rename", "move", "copy"], default="rename")
    p.add_argument("--out", help="move/copy 的目标目录")
    p.add_argument("--archive", help="归档目录（清单/回滚记录/汇总表）")
    p.add_argument("--execute", action="store_true", help="真正执行，缺省仅预览")
    p.set_defaults(func=cmd_apply)

    p = sub.add_parser("report", help="生成 Excel 汇总")
    p.add_argument("--manifest", required=True)
    p.add_argument("--xlsx")
    p.set_defaults(func=cmd_report)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
