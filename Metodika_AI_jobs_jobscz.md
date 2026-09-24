# Metodika: Sledování podílu AI inzerátů na jobs.cz

**Účel:** Měřit v čase, jaký počet a podíl pracovních inzerátů na jobs.cz se zaměřuje na AI role, resp. má AI jako prvek specializace. Měření se opakuje 1× měsíčně a zapisuje do souboru `AI_jobs_tracker_jobscz.xlsx` (list *Souhrn (trend)*).

První měření: **2026-06-04**

---

## 1. Zdroje dat (přesné URL)

| Údaj | URL | Jak se číslo získá |
|---|---|---|
| **Celkem všech inzerátů** | `https://www.jobs.cz/prace/?page=2` | text „Našli jsme **N** nabídek" (bez filtru; `page=2` vynutí vykreslení počtu) |
| **AI shody (fulltext)** | `https://www.jobs.cz/prace/?q[]=ai&page=1..N` | text „Našli jsme **N** nabídek" + stažení všech karet ze stránek |

Pozn.: počet u `q=ai` během dne mírně kolísá (rotace promovaných inzerátů); zaznamenává se hodnota z konkrétního snapshotu.

## 2. Postup sběru

1. Stáhni celkový počet inzerátů z `/prace/?page=2`.
2. Projdi všechny stránky `q[]=ai` (po 30 kartách) a z každé karty vytáhni:
   - `data-jobad-id` (ID inzerátu) → deduplikace,
   - `data-test-ad-title` (název pozice),
   - `alt` u loga firmy (název firmy, pokud je),
   - příznak textu „Příležitost dne" (= **promovaná** karta).
3. Deduplikuj podle ID.

Vše dělá skript `run_monthly_snapshot.py` (jeden příkaz).

## 3. Klasifikační pravidla (deterministická, podle NÁZVU pozice)

Klasifikace běží jen z názvu pozice — je tak 100% reprodukovatelná měsíc po měsíci bez subjektivního čtení popisů.

**Kategorie A — AI-zaměřená role** (AI je jádrem pozice):
název pozice obsahuje některý z těchto signálů:
- case-sensitive tokeny: `AI`, `A.I.`, `ML`, `LLM`, `GenAI`, `NLP`, `MLOps`, `GPT`
- fráze (bez diakritiky, malá písmena): `umela inteligence`, `strojove uceni`, `machine learning`, `generativni`, `gen ai`, `genai`, `deep learning`, `neuronov`, `computer vision`, `pocitacove videni`, `prompt engineer`, `data scientist`, `data science`

**Kategorie B — AI jako prvek / okrajová zmínka:**
inzerát byl nalezen přes `q=ai`, ale v názvu žádný AI signál není → AI je zmíněna až v popisu (jako dovednost, nástroj, kontext nebo benefit).

**Promované (vyřazené):**
karty „Příležitost dne" se zobrazují u každého hledání bez ohledu na dotaz; nejde o skutečné AI shody. Od **AI-relevantních** se odečítají.

## 4. Počítané metriky

- **AI-relevantní (A+B)** = AI shody − promované
- **% AI-rel. z celku** = AI-relevantní / celkem inzerátů
- **% A z celku** = A / celkem inzerátů
- **% A z AI-relevantních** = A / AI-relevantní

## 5. Časování a automatizace (od září 2026)

Měření běží automaticky přes GitHub Actions skriptem `jobscz_ai_monthly.py` (návod: `SETUP_GitHub.md`).

- Workflow se spouští **denně**, skript sám rozhodne, zda měřit.
- **Poslední den v měsíci** změří daný měsíc.
- **Dohnání:** pokud poslední den běh vypadne nebo selže, skript měsíc změří v prvních **10 dnech** následujícího měsíce. Řádek dostane období chybějícího měsíce (sloupec *Období*), datum snapshotu je skutečné datum měření.
- Období je **hotové**, když existuje řádek s tímto obdobím pořízený poslední den měsíce nebo později. Ruční měření v průběhu měsíce (`--force`) se automatickým měřením na konci měsíce přepíše. Původní snapshot zůstává ve složce `snapshots/`.
- Zmeškaný měsíc po 10. dni už dohnat nejde, protože jobs.cz ukazuje jen aktuální inzeráty. V řadě pak zůstane mezera.

Ruční měření kdykoli: v GitHubu *Actions → Run workflow → zaškrtnout force*, nebo lokálně `python jobscz_ai_monthly.py --force`.

## 6. Známá omezení (kvůli korektní interpretaci trendu)

- Klasifikace A/B je podle **názvu**, ne celého popisu. Část rolí, kde je AI reálně významný prvek, ale není v názvu, spadne do B; naopak ojediněle může název obsahovat „AI" jen marketingově.
- `q=ai` je fulltext — zachytí i čistě okrajové či marketingové zmínky AI. Proto je **% AI-rel. z celku horní hranicí** „čehokoli s AI".
- Inzeráty se v čase mění; měř ideálně vždy ve stejný den v měsíci a ve stejnou denní dobu.
- Pravidla v sekci 3 drž **neměnná** — jakákoli úprava klíčových slov ruší srovnatelnost s předchozími měřeními (případnou změnu zapiš s datem do tohoto souboru).

## 7. Výsledek prvního měření (2026-06-04)

| Metrika | Hodnota |
|---|---|
| Celkem inzerátů na jobs.cz | 18 357 |
| Shody `q=ai` (staženo unikátních) | 810 |
| Promované (vyřazeno) | 8 |
| **AI-relevantní (A+B)** | **802** |
| **A — AI-zaměřené role** | **136** |
| **B — AI jako prvek** | **666** |
| % AI-rel. z celku | 4,4 % |
| % A z celku | 0,7 % |
| % A z AI-relevantních | 17,0 % |

---

## 8. Změny metodiky a známé nepravidelnosti

- **2026-06-04, 2026-07-26, 2026-09-23:** první tři měření proběhla v nepravidelné dny (ručně). Srpen 2026 chybí. Od konce září 2026 se měří vždy poslední den v měsíci.
- **2026-09-23:** oprava. Promovaná karta („Příležitost dne“) s AI v názvu se dříve započítávala do kategorie A. Nově se promované karty odečítají z A i B. Starší měření to neovlivnilo (neobsahovala promovanou kartu v kategorii A).
- **2026-09-23:** pozorování. Některé firmy (např. Škoda Auto) inzerují tutéž roli dvakrát (česky „m/ž“ a anglicky „m/f“) a kategorie A je tím mírně nadhodnocená (září: 168 oficiálně, 161 po očištění). Metodiku kvůli srovnatelnosti neměníme, jen evidujeme.
- **2026-09-24:** přidán sloupec *Období* do listu *Souhrn (trend)*. Dosavadní řádky mají období podle data měření.
