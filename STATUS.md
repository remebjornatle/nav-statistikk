# NAV Diagnoseutvikling — Dashboard Status

## Hva er bygd

En statisk HTML/JS-dashboard (`dashboard.html`) med to faner:

### Fane 1: Uføretrygd (18–29 år)
- Kilde: NAV Tabell 1 og Tabell 3 (ICD-10-koder)
- Data: 2012–2025 (enkeltdiagnoser fra 2017)
- Viser: ME vs. andre enkeltdiagnoser, vekstranking, nye mottakere per år, diagnosegrupper indeksert til 2012
- ME (G933) er nest raskest voksende enkeltdiagnose 2017–2025: +253% (234 → 826)

### Fane 2: Sykefravær
- Kilde: NAV SYFRA 2301 (ICPC-koder)
- Data: 2015–2025
- Viser: A04 Slapphet/tretthet som nærmeste proxy for ME/CFS (ingen dedikert ME-kode i ICPC), vekstranking, psykiske lidelser detaljert
- A04 vokste +76% i rate (13.9 → 24.5 per 10 000 avtalte dagsverk)
- OBS: callout-boks i dashboardet forklarer ICPC-begrensningen

## Filer i mappen

| Fil | Beskrivelse |
|-----|-------------|
| `dashboard.html` | Selve dashboardet — åpne direkte i nettleser |
| `tabell1_beholdning.xlsx` | Uføretrygd: beholdning etter diagnosegruppe og kjønn, 2012–2025 |
| `tabell3_enkeltdiagnoser.xlsx` | Uføretrygd: topp 10 enkeltdiagnoser, 2017–2025 |
| `syfra2301_enkeltdiagnoser.xlsx` | Sykefravær: enkeltdiagnoser (tapte dagsverk, rate, tilfeller), 2015–2025 |
| `syfra2300_diagnosegrupper.xlsx` | Sykefravær: diagnosegrupper, 2010–2025 (ikke brukt ennå) |

## Hva gjenstår / neste steg

### Planlagt: GitHub + auto-oppdatering
Var enige om å sette opp:
1. Nytt offentlig GitHub-repo (f.eks. `nav-uforetrygd-dashboard`)
2. `dashboard.html` refaktoreres til å laste data fra `data.json` via `fetch()`
3. Python-skript (`scripts/fetch_data.py`) som:
   - Henter NAV-siden for å finne gjeldende Excel-URL (URL-en inneholder en hash som endres ved oppdatering)
   - Laster ned Excel-filene
   - Parser og skriver `data.json`
4. GitHub Actions workflow (`.github/workflows/update-data.yml`):
   - Kjøres manuelt (workflow_dispatch) + cron én gang i året (januar)
   - Committer og pusher ny `data.json` automatisk
5. GitHub Pages aktiveres for å serve dashboardet

### Andre potensielle utvidelser
- Kjønnssplit-visning for uføretrygd
- Tabell med råtall
- Data fra `syfra2300_diagnosegrupper.xlsx` er lastet ned men ikke brukt ennå (2010–2025, grupper)
- Sykefravær mangler data eldre enn 2015 for enkeltdiagnoser

## Datakilder (NAV)

- Uføretrygd: https://www.nav.no/no/nav-og-samfunn/statistikk/aap-nedsatt-arbeidsevne-og-uforetrygd-statistikk/uforetrygd/diagnoser-uforetrygd
- Sykefravær: https://www.nav.no/no/nav-og-samfunn/statistikk/sykefravar-statistikk/sykefravaersstatistikk-arsstatistikk
- Data publiseres vanligvis én gang i året (uføretrygd per desember, sykefravær ca. mars)

## GitHub-tilgang

`gh` CLI er autentisert som `remebjornatle` med `repo` og `workflow`-scopes — klar til å opprette repo og sette opp Actions.
