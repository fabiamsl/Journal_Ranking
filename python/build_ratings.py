"""Build one combined table of journal ratings from the ABDC list (Excel) and the VHB area lists (PDF/TXT).

Input (folder ../quellen/, see README):
    quellen/ABDC-JQL-*.xlsx                         ABDC Journal Quality List, sheet "2025 JQL"
    quellen/VHB/VHB_Rating_2024_Area_rating_*.txt   VHB area lists as text (pdftotext -layout)
    quellen/VHB/VHB_Rating_2024_Area_rating_*.pdf   used only when the matching .txt is missing

Output (in the tool folder, one level above this script):
    data/journal_ratings.csv    one row per journal, ABDC and every VHB area side by side
    data/journal_ratings.json   same content, used by journal_search.py
    data/vhb_entries.csv        every single VHB entry (area, title, ISSN, rating, votes), for checking
    Journal-Suche.html          search page, open by double-click (data embedded)

Usage:
    python python/build_ratings.py
    python python/build_ratings.py --quellen other/folder

Only the Python standard library is needed. pdftotext (Poppler) is needed only if a VHB .txt is missing.
"""

import argparse
import csv
import difflib
import json
import re
import shutil
import subprocess
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent      # python/
ROOT = HERE.parent                          # tool folder
PAGE = "Journal-Suche.html"

RANK_ABDC = {"A*": 4, "A": 3, "B": 2, "C": 1}
RANK_VHB = {"A+": 5, "A": 4, "B": 3, "C": 2, "D": 1}

# VHB file code -> short German label (the long English name is read from the PDF itself).
VHB_AREAS = {
    "BAFI": "Bankbetriebslehre/Finanzierung",
    "DLM": "Dienstleistungsmanagement",
    "HSM": "Hochschulmanagement",
    "INT": "Internationales Management",
    "LOG": "Logistik",
    "MARK": "Marketing",
    "NAMA": "Nachhaltigkeitsmanagement",
    "OEBWL": "Öffentliche BWL",
    "OR": "Operations Research",
    "ORG": "Organisation",
    "PERS": "Personal",
    "PROD": "Produktionswirtschaft",
    "RECH": "Rechnungswesen",
    "STEU": "Betriebswirtschaftliche Steuerlehre",
    "STRAT": "Strategisches Management",
    "TIE": "Technologie, Innovation, Entrepreneurship",
    "WEW": "Wissenschaftstheorie und Ethik",
    "WI": "Wirtschaftsinformatik",
}

# ABDC field-of-research codes (ANZSRC 2020) that occur in the list.
FOR_CODES = {
    "3501": "Accounting, auditing and accountability",
    "3502": "Banking, finance and investment",
    "3503": "Business systems in context",
    "3504": "Commercial services",
    "3505": "Human resources and industrial relations",
    "3506": "Marketing",
    "3507": "Strategy, management and organisational behaviour",
    "3508": "Tourism",
    "3509": "Transportation, logistics and supply chains",
    "3599": "Other commerce, management, tourism and services",
    "3801": "Applied economics",
    "3802": "Econometrics",
    "3803": "Economic theory",
    "3899": "Other economics",
    "4609": "Information systems",
    "4801": "Commercial law",
    "4905": "Statistics",
    "4901": "Applied mathematics",
    "4903": "Numerical and computational mathematics",
}

ISSN_RE = re.compile(r"^\d{4}-?\d{3}[\dXx]$|^\d{7}-[\dXx]$")


# ---------------------------------------------------------------- helpers

def norm_title(name):
    """Lower case, ASCII, no punctuation, no leading 'the', '&' -> 'and', no '(ABBR)'."""
    s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode().lower()
    s = re.sub(r"\(.*?\)", " ", s)
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    words = s.split()
    if words and words[0] == "the":
        words = words[1:]
    return " ".join(words)


def norm_issn(raw):
    s = re.sub(r"[^0-9Xx]", "", raw or "").upper()
    return f"{s[:4]}-{s[4:]}" if len(s) == 8 else ""


def clean(s):
    return re.sub(r"\s+", " ", (s or "").replace("\t", " ")).strip()


# ---------------------------------------------------------------- ABDC

def read_xlsx_sheet(path, sheet_name):
    m = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    z = zipfile.ZipFile(path)
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    target = None
    for sh in wb.iter(m + "sheet"):
        if sh.get("name") == sheet_name:
            rid = sh.get(r + "id")
            for rel in rels:
                if rel.get("Id") == rid:
                    target = "xl/" + rel.get("Target").lstrip("/").removeprefix("xl/")
    if target is None:
        raise SystemExit(f"Sheet '{sheet_name}' not found in {path.name}")
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        shared = ["".join(t.text or "" for t in si.iter(m + "t"))
                  for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(m + "si")]
    rows = []
    for row in ET.fromstring(z.read(target)).iter(m + "row"):
        vals = {}
        for c in row.findall(m + "c"):
            col = re.match(r"[A-Z]+", c.get("r")).group()
            v, t, inline = c.find(m + "v"), c.get("t"), c.find(m + "is")
            if t == "s" and v is not None:
                vals[col] = shared[int(v.text)]
            elif inline is not None:
                vals[col] = "".join(x.text or "" for x in inline.iter(m + "t"))
            else:
                vals[col] = v.text if v is not None else ""
        rows.append(vals)
    return rows


def load_abdc(path):
    rows = read_xlsx_sheet(path, "2025 JQL")
    header_idx = next(i for i, r in enumerate(rows) if "Journal Title" in map(clean, r.values()))
    header = {col: clean(v) for col, v in rows[header_idx].items()}
    col = {name: c for c, name in header.items()}
    out = []
    for r in rows[header_idx + 1:]:
        title = clean(r.get(col["Journal Title"]))
        rating = clean(r.get(col["2025 rating"])).upper()
        if not title or rating not in RANK_ABDC:
            continue
        for_code = clean(r.get(col["FoR"]))
        out.append({
            "title": title,
            "publisher": clean(r.get(col["Publisher"])),
            "issn": [i for i in (norm_issn(r.get(col["ISSN"])), norm_issn(r.get(col["ISSNOnline"]))) if i],
            "rating": rating,
            "for": for_code,
            "for_name": FOR_CODES.get(for_code, ""),
        })
    return out


# ---------------------------------------------------------------- VHB

def ensure_txt(pdf):
    txt = pdf.with_suffix(".txt")
    if txt.exists():
        return txt
    exe = shutil.which("pdftotext") or r"C:\Program Files\Git\mingw64\bin\pdftotext.exe"
    if not Path(exe).exists() and not shutil.which(exe):
        raise SystemExit(f"{txt.name} fehlt und pdftotext ist nicht installiert.")
    subprocess.run([exe, "-layout", "-enc", "UTF-8", str(pdf), str(txt)], check=True)
    return txt


def fields(line):
    """Split a layout line into (start column, text) pieces separated by 2+ spaces."""
    return [(m.start(), m.group()) for m in re.finditer(r"\S+(?: \S+)*", line)]


SKIP_LINE = re.compile(
    r"^(VHB Publication Media Rating|Criterion|©|letzte redaktionelle|VHB Rating \d{4} in der Version"
    r"|Scientific quality|(A\+|A|B|C|D) [a-z])"
)
HEADER_LINE = re.compile(r"\bTit(le|el)\b.*\bRating\b")


def parse_vhb_file(path, code):
    """Return the entries of one VHB area file (text made with pdftotext -layout).

    The table columns keep their character positions. Long titles are wrapped: one part stands
    on the line above and one on the line below the line that carries ISSN, rating and votes.
    That middle line is the anchor of an entry; title pieces around it are joined to it.
    """
    entries, problems = [], []
    section, pubtype = "", "Scientific journals"
    block, in_categories = [], False

    def flush():
        if block:
            entries.extend(parse_block(block, code, section, pubtype, problems, path.name))
        block.clear()

    for line in path.read_text(encoding="utf-8").splitlines():
        s = line.strip()
        if s.startswith("--- PDF-Seite"):
            flush(); in_categories = False
        elif s.startswith("Section "):
            flush(); section = clean(s[len("Section"):])
        elif s.startswith("Type of publication"):
            flush(); pubtype = clean(s[len("Type of publication"):])
        elif s.startswith("Categories"):
            flush(); in_categories = True
        elif in_categories or SKIP_LINE.match(s) or HEADER_LINE.search(s):
            continue
        else:
            block.append(line)
    flush()
    return entries, problems


# Group labels of the Wirtschaftsinformatik list (left column "Bezug zu WI"). pdftotext sometimes
# leaves only one space between label and title, so the labels are recognised by their text.
WI_LABELS = sorted([
    "Wirtschaftsinformatik or Information", "Systems Journals", "Systems conference proceedings",
    "Journals with or at the interface to", "Systems", "Information Systems",
    "Business and Management Journals", "(without explicit interface to WInf/IS)",
    "Computer Science Journals (without", "explicit interface to WInf/IS)",
    "Discipline-specific Journals (without", "explicit interface to WInf/IS, not",
    "Business/Management, not Computer", "Science)", "Discipline-independent Journals",
    "(without explicit interface to WInf/IS,", "transdisciplinary, interdisciplinary)",
    "Proceedings of conferences with or at the", "interface to Wirtschaftsinformatik or",
], key=len, reverse=True)

JOINING_END = re.compile(r"(\b(and|of|the|for|in|on|und|für|der|die|des|zu|&)|[-–/:,(])\s*$", re.I)


def strip_label(line, title_start):
    """Remove the group label column (WI lists) and keep the title where it is."""
    head = line.lstrip()
    lead = len(line) - len(head)
    if lead >= title_start - 3:
        return line
    for label in WI_LABELS:
        if head.startswith(label) and (len(head) == len(label) or head[len(label)] == " "):
            rest = head[len(label):]
            return " " * (lead + len(label)) + rest if rest.strip() else ""
    # unknown label: cut at the title column
    if line[title_start - 1:title_start] == " ":
        return " " * title_start + line[title_start:]
    return line


def classify(line, title_start, value_start):
    """Split one line into title pieces and values, using the column positions of the block."""
    if title_start > 5:
        line = strip_label(line, title_start)
    if not line.strip():
        return None
    info = {"title": [], "issn": "", "rating": "", "votes": "", "practice": "", "publisher": [], "raw": line}
    for start, text in fields(line):
        if text == "Link":
            continue
        if start < value_start - 2:
            m = re.match(r"^(.*\S)\s+(\d{4}-\d{3}[\dXx])$", text)   # title touching the ISSN
            if m:
                info["title"].append(m.group(1)); info["issn"] = m.group(2)
            else:
                info["title"].append(text)
        elif ISSN_RE.match(text) and not (info["issn"] or info["rating"] or info["votes"]):
            info["issn"] = text
        elif text in RANK_VHB and not info["rating"] and not info["votes"]:
            info["rating"] = text
        elif re.fullmatch(r"\*{1,3}", text) and not info["practice"]:
            info["practice"] = text                              # second criterion (RECH): practice
        elif re.fullmatch(r"\d{1,3}", text) and not info["publisher"]:
            if info["practice"]:
                continue                                         # votes of the practice criterion
            if not info["votes"]:
                info["votes"] = text
        else:
            info["publisher"].append(text)
    return info


def block_columns(block):
    """Title start and first value column, read from the lines that hold an ISSN or a rating."""
    title_starts, value_starts = [], []
    for line in block:
        f = fields(line)
        vals = [st for st, t in f if ISSN_RE.match(t) or t in RANK_VHB]
        if not vals:
            continue
        value_starts.append(min(vals))
        texts = [st for st, t in f if st < min(vals) - 2]
        if texts:
            title_starts.append(texts[-1] if len(texts) > 1 else texts[0])
    if not value_starts:
        return 0, 10**6
    value_start = min(sorted(value_starts)[len(value_starts) // 10:] or value_starts)
    title_start = max(set(title_starts), key=title_starts.count) if title_starts else 0
    return title_start, value_start


def unbalanced(text):
    return text.count("(") != text.count(")")


def parse_block(block, code, section, pubtype, problems, fname):
    """Join wrapped lines to their entries.

    Each entry has one anchor line (the line with the votes number). Free lines between two
    anchors are split: the first part continues the entry above, the rest starts the entry below.
    The split is chosen so that brackets close and no title ends on a joining word.
    """
    title_start, value_start = block_columns(block)
    parsed = [p for p in (classify(line, title_start, value_start) for line in block)]
    for p in parsed:
        if p and p["practice"] and not p["votes"]:
            p["votes"] = "0"            # practice rating only, no scientific rating
    rows = [p for p in parsed if p]
    anchors = [i for i, p in enumerate(rows) if p["votes"]]
    if not anchors:
        for p in rows:
            if p["title"]:
                problems.append(f"{fname}: Zeile nicht zugeordnet: {p['raw'].strip()[:120]}")
        return []

    pre = {i: [] for i in anchors}
    post = {i: [] for i in anchors}
    # lines before the first anchor / after the last one
    pre[anchors[0]] = list(range(0, anchors[0]))
    post[anchors[-1]] = list(range(anchors[-1] + 1, len(rows)))

    def text_of(idx_list):
        return " ".join(" ".join(rows[k]["title"]) for k in idx_list).strip()

    for a, b in zip(anchors, anchors[1:]):
        free = list(range(a + 1, b))
        if not free:
            continue
        m = len(free)
        b_has_post = b + 1 < len(rows) and not rows[b + 1]["votes"] and bool(rows[b + 1]["title"]) \
            and (b == anchors[-1] or b + 1 < anchors[anchors.index(b) + 1])
        best, best_score = None, None
        for s in range(m + 1):
            a_lines = pre[a] + [a] + free[:s]
            b_pre = free[s:]
            a_text = text_of(a_lines)
            score = 0
            if unbalanced(a_text):
                score += 3
            if JOINING_END.search(a_text):
                score += 2
            if b_pre and JOINING_END.search(text_of(b_pre)) is None and not unbalanced(text_of(b_pre)) \
                    and rows[b]["title"] and not b_has_post:
                score += 1                    # a complete line before a complete entry is a continuation
            if not rows[a]["title"] and s == 0:
                score += 2                    # an anchor without title needs a line below
            if not rows[b]["title"] and s == m:
                score += 2                    # ... or a line above
            if best_score is None or score < best_score:
                best, best_score = s, score
        post[a] = free[:best]
        pre[b] = free[best:]

    out = []
    for i in anchors:
        lines = pre[i] + [i] + post[i]
        title = " ".join(" ".join(rows[k]["title"]) for k in lines)
        publisher = " ".join(" ".join(rows[k]["publisher"]) for k in lines)
        rating = rows[i]["rating"] or next((rows[k]["rating"] for k in post[i] + pre[i][::-1] if rows[k]["rating"]), "")
        t = clean(title)
        p = rows[i]
        if not t or not (rating or p["practice"]):
            problems.append(f"{fname}: Eintrag ohne Titel oder Rating: {p['raw'].strip()[:120]}")
            continue
        if unbalanced(t):
            problems.append(f"{fname}: Klammern im Titel pruefen: {t}")
        out.append({
            "area": code,
            "section": section,
            "type": "Proceedings" if pubtype.lower().startswith("proceed") else "Journal",
            "title": t,
            "issn": norm_issn(p["issn"]),
            "issn_raw": p["issn"],
            "rating": rating,
            "votes": int(p["votes"]),
            "practice": p["practice"],
            "publisher": clean(publisher),
        })
    return out


def load_vhb(folder):
    entries, problems = [], []
    pdfs = sorted(folder.glob("VHB_Rating_*_Area_rating_*.pdf"))
    txts = {p.with_suffix(".txt") for p in pdfs} | set(folder.glob("VHB_Rating_*_Area_rating_*.txt"))
    for pdf in pdfs:
        ensure_txt(pdf)
    for txt in sorted(txts):
        raw_code = txt.stem.split("Area_rating_")[1]
        code = raw_code.split("_")[0].upper()
        e, p = parse_vhb_file(txt, code)
        entries += e
        problems += p
    return entries, problems


# ---------------------------------------------------------------- merge

def vhb_rank(r):
    return RANK_VHB.get(r, 0)


class Groups:
    """Union-find: items that share a key end up in one group."""
    def __init__(self):
        self.parent = {}

    def find(self, x):
        self.parent.setdefault(x, x)
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a, b):
        self.parent[self.find(a)] = self.find(b)


def merge(abdc, vhb, problems):
    """Group VHB entries of the same journal, then attach each group to at most one ABDC journal.

    ABDC journals are never merged with each other. A VHB group is attached by its title first and
    by its ISSN only when no title matches; conflicts between the two are reported.
    """
    g = Groups()
    for i, v in enumerate(vhb):
        g.find(("vhb", i))
        if v["issn"]:
            g.union(("vhb", i), ("issn", v["issn"]))
        prefix = "proc: " if v["type"] == "Proceedings" else ""
        g.union(("vhb", i), ("title", prefix + norm_title(v["title"])))
    vhb_groups = {}
    for i, v in enumerate(vhb):
        vhb_groups.setdefault(g.find(("vhb", i)), []).append(v)

    abdc_by_title, abdc_by_issn = {}, {}
    for k, a in enumerate(abdc):
        abdc_by_title.setdefault(norm_title(a["title"]), set()).add(k)
        for i in a["issn"]:
            abdc_by_issn.setdefault(i, set()).add(k)

    attached = {k: [] for k in range(len(abdc))}
    vhb_only = []
    for members in vhb_groups.values():
        if all(v["type"] == "Proceedings" for v in members):
            vhb_only.append(members); continue
        by_title = set().union(*(abdc_by_title.get(norm_title(v["title"]), set()) for v in members))
        by_issn = set().union(*(abdc_by_issn.get(v["issn"], set()) for v in members))
        target = None
        if len(by_title) == 1:
            target = next(iter(by_title))
            if by_issn and target not in by_issn:
                problems.append(f'Titel und ISSN widersprechen sich: VHB "{members[0]["title"]}" '
                                f'({members[0]["issn"]}) -> per Titel "{abdc[target]["title"]}", per ISSN '
                                + ", ".join(f'"{abdc[k]["title"]}"' for k in by_issn) + " (Titel gilt)")
        elif len(by_title) > 1:
            both = by_title & by_issn
            if len(both) == 1:
                target = next(iter(both))
            else:
                problems.append(f'Mehrdeutiger Titel, nicht mit ABDC verknuepft: VHB "{members[0]["title"]}"')
        elif len(by_issn) == 1:
            target = next(iter(by_issn))
        elif len(by_issn) > 1:
            problems.append(f'ISSN passt zu mehreren ABDC-Eintraegen, nicht verknuepft: VHB "{members[0]["title"]}"')
        else:
            # typo in one list, e.g. "Preferences" vs "Preference": very close titles only
            for v in members:
                close = difflib.get_close_matches(norm_title(v["title"]), list(abdc_by_title), n=1, cutoff=0.96)
                if close and len(abdc_by_title[close[0]]) == 1:
                    target = next(iter(abdc_by_title[close[0]]))
                    problems.append(f'Aehnlicher Titel verknuepft: VHB "{v["title"]}" -> ABDC "{abdc[target]["title"]}"')
                    break
        if target is None:
            vhb_only.append(members)
        else:
            attached[target] += members

    by_group = {("abdc", k): [("abdc", abdc[k])] + [("vhb", v) for v in attached[k]] for k in range(len(abdc))}
    for n, members in enumerate(vhb_only):
        by_group[("vhbonly", n)] = [("vhb", v) for v in members]

    journals = []
    for members in by_group.values():
        a_list = [it for s, it in members if s == "abdc"]
        v_list = [it for s, it in members if s == "vhb"]
        a = a_list[0] if a_list else None
        vhb_areas = {}
        for v in v_list:
            old = vhb_areas.get(v["area"])
            if not old or vhb_rank(v["rating"]) > vhb_rank(old["rating"]):
                vhb_areas[v["area"]] = {"rating": v["rating"], "votes": v["votes"]}
                if v.get("practice"):
                    vhb_areas[v["area"]]["practice"] = v["practice"]
        vhb_titles = [v["title"] for v in v_list]
        main_title = a["title"] if a else max(set(vhb_titles), key=vhb_titles.count)
        titles = []
        for t in [main_title] + vhb_titles:
            if t not in titles:
                titles.append(t)
        issns = []
        for i in (a["issn"] if a else []) + [v["issn"] for v in v_list]:
            if i and i not in issns:
                issns.append(i)
        rated = {k: v for k, v in vhb_areas.items() if v["rating"]}
        best = max(rated.items(), key=lambda kv: (vhb_rank(kv[1]["rating"]), kv[1]["votes"]), default=None)
        journals.append({
            "title": main_title,
            "alt_titles": titles[1:],
            "issn": issns,
            "publisher": (a["publisher"] if a else "") or next((v["publisher"] for v in v_list if v["publisher"]), ""),
            "type": "Proceedings" if not a and v_list and all(v["type"] == "Proceedings" for v in v_list) else "Journal",
            "abdc": a["rating"] if a else "",
            "abdc_for": a["for"] if a else "",
            "abdc_for_name": a["for_name"] if a else "",
            "vhb": vhb_areas,
            "vhb_best": best[1]["rating"] if best else "",
            "vhb_best_area": best[0] if best else "",
        })
    journals.sort(key=lambda j: norm_title(j["title"]))
    return journals


# ---------------------------------------------------------------- output

def cell(v):
    """CSV cell for one VHB area: rating, or the practice stars where only those exist (RECH)."""
    if not v:
        return ""
    return v["rating"] or f'Praxis {v.get("practice", "")}'


def write_outputs(journals, vhb, abdc_file, sections, problems):
    data_dir = ROOT / "data"
    data_dir.mkdir(exist_ok=True)
    areas = sorted(VHB_AREAS)

    with open(data_dir / "journal_ratings.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["titel", "weitere_titel", "issn", "verlag", "typ", "abdc_2025", "abdc_for",
                    "vhb_bestes", "vhb_bestes_bereich"] + [f"vhb_{a}" for a in areas])
        for j in journals:
            w.writerow([j["title"], " | ".join(j["alt_titles"]), " ".join(j["issn"]), j["publisher"],
                        j["type"], j["abdc"], j["abdc_for"], j["vhb_best"], j["vhb_best_area"]]
                       + [cell(j["vhb"].get(a)) for a in areas])

    with open(data_dir / "vhb_entries.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["bereich", "typ", "titel", "issn", "rating", "stimmen_prozent", "praxis_rech", "verlag"])
        for v in vhb:
            w.writerow([v["area"], v["type"], v["title"], v["issn_raw"], v["rating"],
                        v["votes"] if v["rating"] else "", v["practice"], v["publisher"]])

    meta = {
        "built": date.today().isoformat(),
        "abdc_source": abdc_file,
        "vhb_source": "VHB-Rating 2024 (Teilratings der Wissenschaftlichen Kommissionen)",
        "areas": {a: {"de": VHB_AREAS[a], "en": sections.get(a, "")} for a in areas},
        "for_codes": FOR_CODES,
        "counts": {
            "journals": len(journals),
            "abdc": sum(1 for j in journals if j["abdc"]),
            "vhb": sum(1 for j in journals if j["vhb"]),
            "both": sum(1 for j in journals if j["abdc"] and j["vhb"]),
        },
    }
    payload = {"meta": meta, "journals": journals}
    (data_dir / "journal_ratings.json").write_text(json.dumps(payload, ensure_ascii=False, indent=0),
                                                   encoding="utf-8")

    template = (HERE / "template.html").read_text(encoding="utf-8")
    compact = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = template.replace("/*__DATA__*/null", compact)
    (ROOT / PAGE).write_text(html, encoding="utf-8")

    (data_dir / "build_hinweise.txt").write_text(
        "Hinweise aus dem letzten Lauf von build_ratings.py (Zuordnung und Einlesen):\n\n"
        + ("\n".join(problems) if problems else "(keine)") + "\n", encoding="utf-8")
    return meta


def main():
    ap = argparse.ArgumentParser(description="Build the combined ABDC/VHB rating table and search page.")
    ap.add_argument("--quellen", default=str(ROOT / "quellen"), help="folder with ABDC .xlsx and VHB/ subfolder")
    args = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    src = Path(args.quellen)

    abdc_files = sorted(src.glob("ABDC*.xlsx"))
    if not abdc_files:
        raise SystemExit(f"Keine ABDC-Excel in {src} gefunden.")
    abdc = load_abdc(abdc_files[-1])
    vhb, problems = load_vhb(src / "VHB")
    sections = {}
    for v in vhb:
        sections.setdefault(v["area"], v["section"])
    journals = merge(abdc, vhb, problems)
    meta = write_outputs(journals, vhb, abdc_files[-1].name, sections, problems)

    c = meta["counts"]
    print(f"ABDC-Einträge: {len(abdc)}   VHB-Einträge: {len(vhb)} in {len(sections)} Bereichen")
    print(f"Zeitschriften gesamt: {c['journals']}  (ABDC: {c['abdc']}, VHB: {c['vhb']}, beide: {c['both']})")
    print(f"Hinweise: {len(problems)}  -> data/build_hinweise.txt")
    print("Geschrieben: data/journal_ratings.csv, data/journal_ratings.json, data/vhb_entries.csv, " + PAGE)


if __name__ == "__main__":
    sys.exit(main())
