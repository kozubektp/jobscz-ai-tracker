#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
jobscz_stats.py — z trackeru vygeneruje deterministicky statisticky souhrn
pro posledni zmerene obdobi: reports/stats_RRRR-MM.json (+ reports/latest.json).

Souhrn cte naplanovana uloha v Claude Cowork a pise z nej mesicni report.
Vsechna cisla se pocitaji zde (presne a opakovatelne), Claude je jen interpretuje.

Nazvy pozic a firem jsou data z verejneho webu (neduveryhodna) - pouze se
kopiruji jako text, nic se z nich neinterpretuje ani nespousti.

Bezi v GitHub Actions po kazdem dennim behu. Obsah nezavisi na case spusteni,
takze se soubor meni (a commituje) jen kdyz se zmeni data v trackeru.
"""
import os, re, json, html, datetime, unicodedata, collections
from openpyxl import load_workbook

TRACKER = "AI_jobs_tracker_jobscz.xlsx"
OUT_DIR = "reports"
SUMMARY = "Souhrn (trend)"
SNAP = "Inzer\u00e1ty (snapshot)"
FIRST_ROW = 5
PERIOD_COL = 11
REPO = "https://github.com/kozubektp/jobscz-ai-tracker"
MAX_NEW_A = 40
MAX_GONE_A = 20

def last_day(year, month):
    nxt = datetime.date(year + (month == 12), month % 12 + 1, 1)
    return nxt - datetime.timedelta(days=1)

def months_between(p1, p2):
    y1, m1 = map(int, p1.split("-")); y2, m2 = map(int, p2.split("-"))
    return (y2 - y1) * 12 + (m2 - m1)

def safe_text(s, limit=150):
    s = "".join(ch for ch in str(s or "") if unicodedata.category(ch)[0] != "C")
    s = re.sub(r"\s+", " ", s).strip()
    return s[:limit]

def num(v):
    return v if isinstance(v, (int, float)) else 0

def read_series(ws):
    rows, r, blanks = {}, FIRST_ROW, 0
    while r < 5000 and blanks < 5:
        a = ws.cell(row=r, column=1).value
        if a in (None, ""):
            blanks += 1; r += 1; continue
        blanks = 0
        date = str(a)[:10]
        period = str(ws.cell(row=r, column=PERIOD_COL).value or date[:7])
        total, matched, promoted, A = (num(ws.cell(row=r, column=c).value) for c in (2, 3, 4, 6))
        if total and matched:
            ai_rel = matched - promoted
            rows[period] = {
                "period": period, "date": date, "total": total, "matched": matched,
                "promoted": promoted, "ai_rel": ai_rel, "A": A, "B": ai_rel - A,
                "pct_ai_of_total": round(ai_rel / total * 100, 2),
                "pct_A_of_total": round(A / total * 100, 3),
                "pct_A_of_ai": round(A / ai_rel * 100, 1) if ai_rel else 0,
            }
        r += 1
    return [rows[k] for k in sorted(rows)]

def read_listings(ws2):
    by_date = collections.defaultdict(list)
    for row in ws2.iter_rows(min_row=2, values_only=True):
        if not row or not row[0] or not row[1]:
            continue
        by_date[str(row[0])[:10]].append({
            "id": str(row[1]), "cat": row[2], "title": safe_text(row[3]),
            "company": safe_text(row[4], 80), "promoted": row[5] == "ano"})
    return by_date

def norm_title(t):
    t = t.lower()
    t = re.sub(r"\((m/[fž](/[dx])?|ž/m|f/m)\)", "", t)
    return re.sub(r"\s+", " ", t).strip()

def top_companies(listings, cat=None, n=10):
    c = collections.Counter(x["company"] for x in listings
                            if x["company"] and not x["promoted"] and (cat is None or x["cat"] == cat))
    return [{"company": k, "count": v} for k, v in sorted(c.items(), key=lambda kv: (-kv[1], kv[0]))[:n]]

def delta(prev, last, key):
    a, b = prev[key], last[key]
    return {"prev": a, "last": b, "abs": b - a, "pct": round((b - a) / a * 100, 1) if a else None}

DEFINITIONS = {
    "A": "AI-zamerene role: AI/ML termin primo v NAZVU pozice (bez promovanych karet).",
    "B": "AI jako prvek: inzerat nalezeny fulltextem 'ai', AI neni v nazvu pozice.",
    "ai_rel": "AI-relevantni inzeraty = A + B.",
    "promoted": ("Pocet promovanych karet 'Prilezitost dne', ktere jobs.cz vklada do kazdeho vyhledavani. "
                 "Nejsou to AI shody a z A i B se VYRAZUJI. Nejde o presun mezi kategoriemi."),
    "final": "true = mereni z posledniho dne mesice (nebo dohnane po nem); false = predbezne/rucni mereni.",
    "churn": ("Porovnani ID inzeratu mezi dvema merenimi: survived = v obou, new = jen v poslednim, "
              "gone = jen v predchozim."),
    "A_dedup_latest": "A po slouceni stejne role inzerovane cesky i anglicky (m/z vs m/f).",
    "gap_months": "Pocet chybejicich mesicu mezi predchozim a poslednim merenim.",
}

CZ_MONTHS = ["leden", "\u00fanor", "b\u0159ezen", "duben", "kv\u011bten", "\u010derven", "\u010dervenec",
             "srpen", "z\u00e1\u0159\u00ed", "\u0159\u00edjen", "listopad", "prosinec"]
NAVY, GREEN, RED, GREY, LIGHT = "#19325A", "#15803d", "#b91c1c", "#6b7280", "#f3f4f6"
DASH = "\u2014"
PLACEHOLDERS = ["{{HLAVNI_ZJISTENI}}", "{{CO_SE_ZMENILO}}", "{{INTERPRETACE}}"]

def e(s):
    return html.escape(str(s), quote=True)

def fi(n):
    return f"{int(n):,}".replace(",", "\u00a0")

def fp(x, dec=1):
    return f"{x:.{dec}f}".replace(".", ",") + "\u00a0%"

def cz_date(iso):
    d = datetime.date.fromisoformat(iso)
    return f"{d.day}.\u00a0{d.month}.\u00a0{d.year}"

def change(v, unit, dec=1):
    if v is None:
        return f'<span style="color:{GREY};">\u2014</span>'
    col, arr = (GREEN, "\u25b2") if v > 0 else (RED, "\u25bc") if v < 0 else (GREY, "\u25a0")
    num_s = f"{v:+.{dec}f}".replace(".", ",")
    return f'<span style="color:{col};font-weight:bold;">{arr} {num_s}\u00a0{unit}</span>'

def kpi(label, value, sub):
    return (f'<td width="33%" valign="top" style="padding:6px;">'
            f'<table width="100%" cellpadding="0" cellspacing="0" style="background:{LIGHT};border-radius:6px;">'
            f'<tr><td style="padding:14px 14px 4px 14px;font-size:12px;color:{GREY};">{label}</td></tr>'
            f'<tr><td style="padding:0 14px;font-size:26px;font-weight:bold;color:{NAVY};">{value}</td></tr>'
            f'<tr><td style="padding:4px 14px 14px 14px;font-size:12px;">{sub}</td></tr></table></td>')

def section(title, inner):
    return (f'<tr><td style="padding:18px 28px 4px 28px;font-size:15px;font-weight:bold;color:{NAVY};'
            f'border-top:1px solid #e5e7eb;">{title}</td></tr>'
            f'<tr><td style="padding:4px 28px 8px 28px;font-size:14px;line-height:1.5;">{inner}</td></tr>')

def ul(placeholder):
    return f'<ul style="margin:6px 0;padding-left:20px;">{placeholder}</ul>'

def next_period(p):
    y, m = map(int, p.split("-"))
    return f"{y + (m == 12):04d}-{m % 12 + 1:02d}"

def trend_table(series):
    series = series[-12:]
    mx = max(s["pct_ai_of_total"] for s in series) or 1
    th = f'style="padding:6px 8px;font-size:11px;color:{GREY};text-align:right;border-bottom:1px solid #e5e7eb;"'
    rows = [f'<tr><td {th.replace("right", "left")}>Obdob\u00ed</td><td {th}>Celkem</td>'
            f'<td {th}>AI-rel.</td><td {th}>Pod\u00edl AI</td><td {th.replace("right", "left")}></td>'
            f'<td {th}>A</td><td {th}>A z celku</td></tr>']
    prev = None
    for s in series:
        if prev:
            p = next_period(prev)
            while p < s["period"]:
                rows.append(f'<tr><td colspan="7" style="padding:4px 8px;font-size:12px;color:{GREY};'
                            f'font-style:italic;">{p} \u2014 m\u011b\u0159en\u00ed chyb\u00ed</td></tr>')
                p = next_period(p)
        prev = s["period"]
        w = max(4, round(s["pct_ai_of_total"] / mx * 100))
        td = 'style="padding:6px 8px;font-size:13px;text-align:right;"'
        rows.append(
            f'<tr><td style="padding:6px 8px;font-size:13px;">{e(s["period"])}</td>'
            f'<td {td}>{fi(s["total"])}</td><td {td}>{fi(s["ai_rel"])}</td>'
            f'<td {td}><b>{fp(s["pct_ai_of_total"])}</b></td>'
            f'<td width="110" style="padding:6px 8px;"><table width="{w}%" cellpadding="0" cellspacing="0">'
            f'<tr><td height="10" bgcolor="{NAVY}" style="font-size:1px;line-height:1px;">&nbsp;</td></tr></table></td>'
            f'<td {td}>{fi(s["A"])}</td><td {td}>{fp(s["pct_A_of_total"], 2)}</td></tr>')
    return f'<table width="100%" cellpadding="0" cellspacing="0">{"".join(rows)}</table>'

def company_table(latest, previous):
    prev_map = {c["company"]: c["count"] for c in (previous or [])}
    th = f'style="padding:6px 8px;font-size:11px;color:{GREY};border-bottom:1px solid #e5e7eb;'
    rows = [f'<tr><td {th}text-align:left;">Firma</td><td {th}text-align:right;">A role nyn\u00ed</td>'
            f'<td {th}text-align:right;">minule</td></tr>']
    for c in latest[:6]:
        rows.append(f'<tr><td style="padding:5px 8px;font-size:13px;">{e(c["company"])}</td>'
                    f'<td style="padding:5px 8px;font-size:13px;text-align:right;"><b>{c["count"]}</b></td>'
                    f'<td style="padding:5px 8px;font-size:13px;text-align:right;color:{GREY};">'
                    f'{prev_map.get(c["company"], DASH)}</td></tr>')
    return f'<table width="100%" cellpadding="0" cellspacing="0">{"".join(rows)}</table>'

def render_email(out):
    last, prev = out["latest"], out.get("previous")
    y, m = map(int, last["period"].split("-"))
    month_label = f"{CZ_MONTHS[m - 1]} {y}"
    sub = f"M\u011b\u0159en\u00ed {cz_date(last['date'])}"
    if prev:
        sub += f" \u00b7 srovn\u00e1n\u00ed s {cz_date(prev['date'])}"
    badge = ""
    if not out["final"]:
        badge = (f'<tr><td style="padding:12px 28px 0 28px;"><span style="background:#fef3c7;color:#92400e;'
                 f'font-size:12px;padding:4px 10px;border-radius:10px;">P\u0159edb\u011b\u017en\u00e1 data '
                 f'\u2014 ne z konce m\u011bs\u00edce</span></td></tr>')
    d = out.get("deltas", {})
    sc = out.get("share_change_pp", {})
    A_sub = change(d["A"]["pct"], "%") if d else ""
    if out["A_dedup_latest"] < last["A"]:
        A_sub += f'<br><span style="color:{GREY};">{out["A_dedup_latest"]} bez duplicit</span>'
    kpis = (kpi("Pod\u00edl AI inzer\u00e1t\u016f", fp(last["pct_ai_of_total"]),
                change(sc.get("pct_ai_of_total"), "p.\u00a0b.") if sc else "")
            + kpi("AI-zam\u011b\u0159en\u00e9 role (A)", fi(last["A"]), A_sub)
            + kpi("Celkem inzer\u00e1t\u016f", fi(last["total"]),
                  change(d["total"]["pct"], "%") if d else ""))
    notes = "".join(f"<li>{e(n)}</li>" for n in out["notes"]) or "<li>Bez pozn\u00e1mek.</li>"
    footer = (f'<tr><td style="padding:16px 28px 8px 28px;border-top:1px solid #e5e7eb;font-size:12px;'
              f'color:{GREY};line-height:1.5;"><b>Metodika:</b> A = AI/ML term\u00edn v n\u00e1zvu pozice, '
              f'B = AI jen v popisu, AI-rel. = A+B; promovan\u00e9 karty jsou vy\u0159azen\u00e9. '
              f'Jeden port\u00e1l (jobs.cz), klasifikace podle n\u00e1zvu.'
              f'<ul style="margin:6px 0;padding-left:18px;">{notes}</ul></td></tr>'
              f'<tr><td style="padding:8px 28px 24px 28px;"><a href="{e(out["tracker_url"])}" '
              f'style="background:{NAVY};color:#ffffff;text-decoration:none;font-size:13px;'
              f'padding:9px 16px;border-radius:5px;display:inline-block;">Otev\u0159\u00edt data (Excel)</a>'
              f'<span style="font-size:12px;color:{GREY};">&nbsp;&nbsp;Jobs.cz AI tracker</span></td></tr>')
    body = (f'<tr><td style="background:{NAVY};padding:20px 28px;border-radius:8px 8px 0 0;">'
            f'<div style="font-size:12px;color:#c7d2fe;letter-spacing:1px;">JOBS.CZ AI REPORT</div>'
            f'<div style="font-size:22px;font-weight:bold;color:#ffffff;">{e(month_label)}</div>'
            f'<div style="font-size:12px;color:#c7d2fe;">{sub}</div></td></tr>'
            f'{badge}'
            f'<tr><td style="padding:14px 22px 6px 22px;"><table width="100%" cellpadding="0" cellspacing="0">'
            f'<tr>{kpis}</tr></table></td></tr>'
            + section("Hlavn\u00ed zji\u0161t\u011bn\u00ed", ul(PLACEHOLDERS[0]))
            + section("V\u00fdvoj v \u010dase", trend_table(out["series"]))
            + section("Co se zm\u011bnilo", ul(PLACEHOLDERS[1]))
            + section("Kdo nab\u00edr\u00e1 AI role", company_table(out["top_A_companies_latest"],
                                                                     out.get("top_A_companies_previous")))
            + section("Interpretace", ul(PLACEHOLDERS[2]))
            + footer)
    return ('<!DOCTYPE html><html><head><meta charset="utf-8"></head>'
            f'<body style="margin:0;padding:0;background:{LIGHT};">'
            f'<table width="100%" cellpadding="0" cellspacing="0" bgcolor="{LIGHT}"><tr>'
            f'<td align="center" style="padding:24px 10px;">'
            f'<table width="640" cellpadding="0" cellspacing="0" style="max-width:640px;width:100%;'
            f'background:#ffffff;border-radius:8px;font-family:Arial,Helvetica,sans-serif;color:#1f2937;">'
            f'{body}</table></td></tr></table></body></html>')

def main():
    wb = load_workbook(TRACKER, read_only=True)
    series = read_series(wb[SUMMARY])
    listings = read_listings(wb[SNAP])
    if not series:
        print("[stats] tracker nema data, nic negeneruji"); return 0
    last = series[-1]
    prev = series[-2] if len(series) > 1 else None
    y, m = map(int, last["period"].split("-"))
    L = [x for x in listings.get(last["date"], [])]
    LA = {x["id"]: x for x in L if x["cat"] == "A" and not x["promoted"]}

    out = {
        "schema": 2,
        "period": last["period"],
        "snapshot_date": last["date"],
        "final": last["date"] >= last_day(y, m).isoformat(),
        "source": REPO,
        "tracker_url": f"{REPO}/blob/main/{TRACKER}",
        "series": series,
        "latest": last,
        "previous": prev,
        "top_A_companies_latest": top_companies(L, "A"),
        "top_all_companies_latest": top_companies(L, None),
        "A_dedup_latest": len({(norm_title(x["title"]), x["company"]) for x in LA.values()}),
        "notes": [],
    }

    if prev:
        P = [x for x in listings.get(prev["date"], [])]
        PA = {x["id"]: x for x in P if x["cat"] == "A" and not x["promoted"]}
        Lid, Pid = {x["id"] for x in L}, {x["id"] for x in P}
        out["gap_months"] = months_between(prev["period"], last["period"]) - 1
        out["deltas"] = {k: delta(prev, last, k) for k in ("total", "ai_rel", "A", "B")}
        out["share_change_pp"] = {k: round(last[k] - prev[k], 3)
                                  for k in ("pct_ai_of_total", "pct_A_of_total", "pct_A_of_ai")}
        out["churn"] = {
            "all": {"survived": len(Lid & Pid), "new": len(Lid - Pid), "gone": len(Pid - Lid)},
            "A": {"survived": len(set(LA) & set(PA)), "new": len(set(LA) - set(PA)),
                  "gone": len(set(PA) - set(LA))},
            "listing_data_available": bool(L) and bool(P),
        }
        new_A = sorted((LA[i] for i in set(LA) - set(PA)), key=lambda x: (x["company"] == "", x["company"], x["title"]))
        gone_A = sorted((PA[i] for i in set(PA) - set(LA)), key=lambda x: (x["company"] == "", x["company"], x["title"]))
        out["new_A_roles"] = [f'{x["title"]} | {x["company"]}' for x in new_A[:MAX_NEW_A]]
        out["new_A_roles_truncated"] = max(0, len(new_A) - MAX_NEW_A)
        out["gone_A_roles_sample"] = [f'{x["title"]} | {x["company"]}' for x in gone_A[:MAX_GONE_A]]
        out["top_A_companies_previous"] = top_companies(P, "A")
        if out["gap_months"] > 0:
            out["notes"].append(f'Mezi obdobími {prev["period"]} a {last["period"]} chybí '
                                f'měření za {out["gap_months"]} měs.')
    if not out["final"]:
        out["notes"].append("Měření není z posledního dne měsíce (předběžné / ruční).")
    if out["A_dedup_latest"] < last["A"]:
        out["notes"].append(f'Po očištění o duplicitní inzeráty (stejná role CZ/EN) je A = '
                            f'{out["A_dedup_latest"]} místo {last["A"]}.')

    out["definitions"] = DEFINITIONS
    out["email_placeholders"] = PLACEHOLDERS
    out["email_template_html"] = render_email(out)

    os.makedirs(OUT_DIR, exist_ok=True)
    text = json.dumps(out, ensure_ascii=False, indent=1, sort_keys=True)
    for name in (f"stats_{last['period']}.json", "latest.json"):
        with open(os.path.join(OUT_DIR, name), "w", encoding="utf-8") as f:
            f.write(text + "\n")
    print(f"[stats] {OUT_DIR}/stats_{last['period']}.json (final={out['final']})")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
