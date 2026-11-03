"""Search the combined ABDC/VHB table (data/journal_ratings.json) from the command line or from Python.

Command line:
    python python/journal_search.py "journal of finance"
    python python/journal_search.py jfe "review of finance" 1540-6261        several queries at once
    python python/journal_search.py "strategic management" --bereich STRAT   show this VHB area first
    python python/journal_search.py "management science" --alle              list every VHB area
    python python/journal_search.py --datei journals.txt                     one journal per line, best hit each
    python python/journal_search.py "journal of finance" --json              machine-readable output

Python:
    from journal_search import search, lookup
    lookup("Journal of Finance")          -> best hit (dict) or None
    search("finance", limit=10)           -> list of (score, journal)

Only the Python standard library is used.
"""

import argparse
import difflib
import json
import re
import sys
import unicodedata
from functools import lru_cache
from pathlib import Path

HERE = Path(__file__).resolve().parent
DATA = HERE.parent / "data" / "journal_ratings.json"
STOP = {"of", "and", "the", "for", "in", "on", "a", "an", "to", "und", "fur", "der", "die", "das", "de", "la"}


def norm(text):
    s = unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode().lower()
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return " ".join(s.split())


def acronyms(title):
    """'Journal of Financial Economics' -> 'jfe'; '(JFQA)' in the title is used as well."""
    out = {m.lower() for m in re.findall(r"\(([A-Z][A-Za-z&]{1,9})\)", title)}
    out |= {m.lower() for m in re.findall(r"\b([A-Z]{2,8})\b\s*[:\-–]", title)}
    words = [w for w in norm(re.sub(r"\(.*?\)", " ", title)).split() if w not in STOP]
    if len(words) >= 2:
        out.add("".join(w[0] for w in words))
    return out


@lru_cache(maxsize=1)
def load(path=str(DATA)):
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    for j in payload["journals"]:
        names = [j["title"]] + j["alt_titles"]
        j["_keys"] = list(dict.fromkeys(norm(n) for n in names))
        j["_keys_nothe"] = [re.sub(r"^the ", "", k) for k in j["_keys"]]
        j["_acr"] = set().union(*(acronyms(n) for n in names))
        # main title without subtitle: "Theory and Decision: an international journal ..." -> "theory and decision"
        j["_main"] = {re.sub(r"^the ", "", norm(re.split(r"\s*(?::|\s[-–/]\s)\s*", n)[0])) for n in names}
    return payload


def score(q, j):
    """0..100: exact title > title start > all words > acronym > close spelling."""
    if not q:
        return 0
    if re.fullmatch(r"\d{4}-?\d{3}[\dx]", q.replace(" ", "")):
        issn = q.replace(" ", "").upper()
        issn = issn if "-" in issn else issn[:4] + "-" + issn[4:]
        return 100 if issn in j["issn"] else 0
    qn = re.sub(r"^the ", "", q)
    best = 0
    words = [w for w in qn.split() if w not in STOP] or qn.split()
    for k in j["_keys_nothe"]:
        if k == qn:
            return 100
        if qn in j["_main"]:
            best = max(best, 95)
        if k.startswith(qn + " ") or k.startswith(qn):
            best = max(best, 88 - min(len(k) - len(qn), 30) * 0.2)
        kw = k.split()
        if all(any(x.startswith(w) for x in kw) for w in words):
            cover = len(" ".join(words)) / max(len(k), 1)
            best = max(best, 60 + 20 * cover)
        elif qn in k:
            best = max(best, 55)
    if qn.replace(" ", "") in j["_acr"] and len(qn) >= 2:
        best = max(best, 82)
    if best < 50 and len(qn) >= 5:
        r = max(difflib.SequenceMatcher(None, qn, k).ratio() for k in j["_keys_nothe"])
        if r >= 0.75:
            best = max(best, 45 * r)
    return best


def prominence(j):
    return ({"A*": 4, "A": 3, "B": 2, "C": 1}.get(j["abdc"], 0)
            + {"A+": 5, "A": 4, "B": 3, "C": 2, "D": 1}.get(j["vhb_best"], 0))


def search(query, limit=10, data=None):
    data = data or load()
    q = norm(query)
    hits = [(score(q, j), j) for j in data["journals"]]
    hits = [h for h in hits if h[0] > 0]
    # equal match quality: better rated journals first (helps with acronyms such as "jfe")
    hits.sort(key=lambda h: (-round(h[0]), -prominence(h[1]), len(h[1]["title"])))
    return hits[:limit]


def lookup(name):
    """Best hit for a full journal name, or None if nothing fits well."""
    hits = search(name, limit=1)
    return hits[0][1] if hits and hits[0][0] >= 70 else None


def public(j):
    return {k: v for k, v in j.items() if not k.startswith("_")}


# ---------------------------------------------------------------- output

def fmt_vhb(j, area, show_all, areas):
    if not j["vhb"]:
        return ["VHB 2024   nicht gelistet"]
    lines = []
    order = sorted(j["vhb"].items(), key=lambda kv: (-{"A+": 5, "A": 4, "B": 3, "C": 2, "D": 1}.get(kv[1]["rating"], 0), kv[0]))
    if area:
        a = j["vhb"].get(area)
        lines.append(f"VHB {area:<6} {a['rating'] or 'nur Praxis-Rating ' + a.get('practice', '')}"
                     if a else f"VHB {area:<6} nicht in diesem Bereich gelistet")
    shown = order if show_all else order[:4]
    parts = []
    for code, v in shown:
        r = v["rating"] or f"Praxis {v.get('practice', '')}"
        parts.append(f"{code} {r}")
    more = "" if show_all or len(order) <= 4 else f"  (+{len(order) - 4} weitere, --alle)"
    lines.append("VHB 2024   " + ", ".join(parts) + more)
    return lines


def print_hit(s, j, area, show_all, areas):
    issn = ", ".join(j["issn"]) or "-"
    print(f"\n{j['title']}   [{issn}]   Treffer {s:.0f}")
    if j["alt_titles"]:
        print(f"  auch: {' | '.join(j['alt_titles'][:3])}")
    abdc = j["abdc"] or "nicht gelistet"
    field = f"  ({j['abdc_for']} {j['abdc_for_name']})" if j["abdc_for"] else ""
    print(f"  ABDC 2025  {abdc}{field}")
    for line in fmt_vhb(j, area, show_all, areas):
        print("  " + line)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="ABDC 2025 und VHB 2024 nebeneinander suchen.")
    ap.add_argument("queries", nargs="*", help="Zeitschriftenname, Abkuerzung (z. B. JFE) oder ISSN")
    ap.add_argument("--bereich", help="VHB-Bereich, der zuerst gezeigt wird, z. B. BAFI, MARK, STRAT")
    ap.add_argument("--alle", action="store_true", help="alle VHB-Bereiche zeigen")
    ap.add_argument("-n", type=int, default=3, help="Treffer je Suche (Standard 3)")
    ap.add_argument("--datei", help="Textdatei mit einem Zeitschriftennamen je Zeile")
    ap.add_argument("--json", action="store_true", help="Ausgabe als JSON")
    ap.add_argument("--bereiche", action="store_true", help="VHB-Bereichskuerzel auflisten")
    args = ap.parse_args()

    data = load()
    areas = data["meta"]["areas"]
    if args.bereiche:
        for code, names in areas.items():
            print(f"{code:<6} {names['de']}  /  {names['en']}")
        return 0
    queries = list(args.queries)
    if args.datei:
        queries += [l.strip() for l in Path(args.datei).read_text(encoding="utf-8").splitlines() if l.strip()]
        args.n = 1 if args.n == 3 else args.n
    if not queries:
        ap.print_help()
        return 1
    area = args.bereich.upper() if args.bereich else None
    if area and area not in areas:
        raise SystemExit(f"Unbekannter Bereich {area}. Moeglich: {', '.join(areas)}")

    result = {}
    for q in queries:
        hits = search(q, limit=args.n, data=data)
        if args.json:
            result[q] = [dict(score=round(s, 1), **public(j)) for s, j in hits]
            continue
        print(f"\n=== {q}")
        if not hits:
            print("  kein Treffer")
        for s, j in hits:
            print_hit(s, j, area, args.alle, areas)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
