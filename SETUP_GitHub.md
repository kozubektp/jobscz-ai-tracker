# Rozchození na GitHubu — krok za krokem

Celkem **cca 20 minut**, jednorázově. Pak to běží samo a nic dalšího nemusíš dělat.

**Co to dělá:** GitHub každý den ráno (7:00/8:00 českého času) spustí skript. Ten poslední den v měsíci změří jobs.cz. Kdyby ten den selhal, dožene měření do 10. dne dalšího měsíce. Výsledky uloží zpátky do repozitáře (tracker + snapshot). V ostatní dny jen zkontroluje, že není co dělat, a skončí.

Budeš potřebovat rozbalený ZIP `jobscz-ai-tracker.zip`. Obsahuje:

```
.github/workflows/monthly.yml   ← plán spouštění
jobscz_ai_monthly.py            ← skript
requirements.txt                ← knihovny pro Python
AI_jobs_tracker_jobscz.xlsx     ← tvůj tracker (červen, červenec, září)
snapshots/                      ← dosavadní snapshoty
Metodika_AI_jobs_jobscz.md, SETUP_GitHub.md, .gitignore
```

---

## 1. Účet (5 min, přeskoč, pokud už ho máš)

1. Jdi na **github.com/signup**, zadej e-mail, heslo a uživatelské jméno a potvrď kód z e-mailu.
2. Plán zvol **Free**. Na tohle bohatě stačí.
3. GitHub tě může požádat o **dvoufázové ověření (2FA)**. Zapni ho přes autentikační aplikaci v mobilu (Microsoft/Google Authenticator) a **ulož si záložní kódy**.

## 2. Repozitář (2 min)

1. Vpravo nahoře klikni na **+** → **New repository**.
2. **Repository name:** `jobscz-ai-tracker`
3. Zvol **Private**.
4. Nic dalšího nezaškrtávej (ani README) a klikni na **Create repository**.

## 3. Nahrání souborů (5 min)

1. Na stránce prázdného repozitáře klikni na odkaz **uploading an existing file**.
2. Otevři rozbalenou složku v Průzkumníku, **označ všechno uvnitř** (Ctrl+A, včetně složek `.github` a `snapshots`) a přetáhni to do okna prohlížeče.
3. Dole klikni na **Commit changes**.
4. **Kontrola:** v záložce **Code** musíš vidět složku `.github`. Klikni do ní, musí obsahovat `workflows/monthly.yml`.

   *Pokud tam `.github` chybí* (některé prohlížeče skryté složky nepřenesou): **Add file → Create new file**, jako název napiš přesně `.github/workflows/monthly.yml` (lomítka vytvoří složky), do obsahu vlož text souboru `monthly.yml` otevřeného v Poznámkovém bloku a dej **Commit changes**.

## 4. Povolení zápisu (1 min)

Bez tohohle kroku workflow proběhne, ale výsledky neuloží.

1. V repozitáři nahoře klikni na **Settings** (ozubené kolo).
2. Vlevo **Actions → General**.
3. Sjeď dolů k **Workflow permissions**, vyber **Read and write permissions** a klikni na **Save**.

## 5. Test (3 min)

1. Nahoře klikni na záložku **Actions**.
2. Vlevo vyber **Jobs.cz AI monthly snapshot**.
3. Vpravo klikni na **Run workflow**, **zaškrtni** „Zmerit hned“ a potvrď zeleným **Run workflow**.
4. Po pár vteřinách se objeví běh (žlutý kroužek). Za 1–3 minuty by měl být **zelený ✓**.
5. Přepni na **Code**. Nahoře uvidíš commit „AI snapshot RRRR-MM-DD“, aktualizovaný tracker a nový soubor ve `snapshots/`.

Test přepíše zářijový řádek dnešními čísly. To je v pořádku, protože 30. 9. ho automatický běh přepíše finálním měřením na konci měsíce.

## 6. Upozornění na chyby (1 min)

1. Vpravo nahoře klikni na svůj avatar → **Settings → Notifications**.
2. V sekci **System → Actions** nech zapnuté **Email** a volbu **Only notify for failed workflows**.

Když běh selže (typicky proto, že jobs.cz změní web), přijde ti e-mail.

**Hotovo.** Od teď to běží samo.

---

## Běžné používání

- **Stažení trackeru:** v záložce **Code** klikni na `AI_jobs_tracker_jobscz.xlsx` a vpravo na ikonu **Download raw file** (šipka dolů).
- **Kontrola běhů:** v záložce **Actions**. Denní běhy bez měření jsou taky zelené a trvají pár sekund. To je správně.
- **Ruční měření kdykoli:** stejně jako v kroku 5 (Run workflow se zaškrtnutím).
- **Tracker v repozitáři needituj ručně.** Když si ho chceš upravit, stáhni si kopii. Jinak hrozí konflikt se zápisem skriptu.

## Když něco selže

| Co vidíš | Co to znamená | Co udělat |
|---|---|---|
| Červený ✗, v logu `Nepodarilo se precist celkovy pocet` nebo `podezrele malo` | jobs.cz změnil web | Pošli mi text chyby z logu a upravím parser. Pokud to stihneš do 10. dne v měsíci, měsíc se ještě dožene. |
| Červený ✗ u kroku **Commit results**, chyba `403` / `Permission denied` | chybí povolení zápisu | Zopakuj krok 4. |
| V Actions není žádný workflow | chybí `.github/workflows/monthly.yml` | Zopakuj kontrolu v kroku 3. |

## Náklady

Zdarma. Soukromý repozitář má ve Free plánu 2 000 minut běhu měsíčně a tohle spotřebuje zhruba 35 (denní kontroly + jedno měření).
