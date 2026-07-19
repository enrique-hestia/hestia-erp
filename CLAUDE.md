# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

**VestaOS** — the ERP of the Hestia Fertility clinic. It grew from a dashboard into a full ERP: income/expenses, banks and reconciliation, A/P and A/R, payroll, inventory, CFDI invoicing, reports and a printable manual. The frontend is one self-contained HTML file (`hestia-fertility-dashboard.html`, ~40k lines); the backend is a multi-file Google Apps Script project over Google Sheets.

There is no build step and no package manager. The frontend is served from GitHub Pages (`git push` = deploy).

## Architecture

### Two-layer system

**Frontend** (`hestia-fertility-dashboard.html`) — One HTML file with all CSS, HTML (every view panel) and JavaScript inlined. No bundler, no imports. Line 1 carries a BOM and a `<!-- app-build:YYYY.MM.DD.NN -->` marker that must be bumped on every change. Validate with `node validate-html.js` before pushing.

**Backend** — **38 `.gs` files** in one Apps Script project (NOT a single `api.gs`). They share ONE global scope, so a duplicated `var`/function across files silently wins by load order — see the warning in `deploy-gas.ps1`. Key entry points:
- **`core.gs` → `doGet`** — the ACTIVE GET router (add GET actions here).
- **`finance.gs` → `doPost`** — the POST router, plus most financial logic.
- `config.gs` — global constants (`SHEET_ID`, `BANKS_SS_ID`, `PRODUCTOS_SS_ID`, `PACIENTES_SS_ID`, …). Data lives across SEVERAL spreadsheets, not one.

The deployed Web App URL is hardcoded in the HTML as `APPS_SCRIPT_URL`.

### Data flow

```
Browser → fetch(APPS_SCRIPT_URL?action=…)   → core.gs doGet   → Sheets → JSON
Browser → fetch(APPS_SCRIPT_URL, {POST})    → finance.gs doPost → Sheets
```

Writes carry a session token; `_tokenHasPermission(token, perm)` (finance.gs) is the lock, fail-closed. A `window.fetch` wrapper injects the token into every POST — GET calls must pass `&token=` explicitly.

### View router

Navigation is purely DOM-based: all views (`#view-resumen`, `#view-ingresos`, etc.) are rendered in the HTML and hidden/shown via `display:none`. `navigateTo(view)` is the single routing function. `renderView(view, data)` lazy-renders charts for non-default views on first visit.

### Global state

- `sheetsData` — cached API response (null until first load)
- `currentSegment` — `'todos'` | `'local'` | `'internacional'`
- `currentPeriodo` — active period ID (e.g. `'2026-Q2'`)
- Chart instances stored as module-level vars (`ingresoChart`, `donutChart`, etc.) and destroyed before re-creating to prevent Chart.js memory leaks

### Google Sheets structure

Each periodic data sheet (`Mensual_Todos`, `Mensual_Local`, `Mensual_Internacional`, `CashFlow`) has a `Periodo` column as the first column (e.g. `2026-Q2`). Capture sheets (`Inventarios`, `Laboratorios`) also have `Periodo` as column 0. Static reference sheets (`Servicios`, `Funnel`, `Alertas`, `DonutServicios`, `DistribucionCostos`) are not period-filtered.

## Deploying

**Backend** — run `powershell -ExecutionPolicy Bypass -File .\deploy-gas.ps1` (one command). It copies only the `.gs` in its `$FILES` list into `gas-deploy\`, runs `clasp push` and re-deploys the Web App keeping the SAME URL. **When adding a new `.gs`, add it to `$FILES` or it will not deploy.** If clasp fails with `invalid_grant` / `invalid_rapt`, the token expired: the user must run `clasp login`.

**Frontend** — bump the `app-build` marker, run `node validate-html.js`, then commit and `git push` (GitHub Pages serves it). Always push after committing.

## Watch out

- **Shared global scope** across all `.gs`: never define the same function in two files.
- **Blanking bug** (has bitten 4+ times): building an update payload with `getElementById(x).value || ''` sends `''` for fields the form did not pre-fill, and the backend then WIPES them. If you add a field to a form, make sure the edit path pre-fills it AND that whatever serializes the record includes it.
- **Money paths** (banks, income, expenses) are the most delicate code. Prefer "config in front, existing code behind" so behavior stays byte-identical, and verify with a harness that loads the real functions.
- Every screen must be reachable from the menu — Panel de Control → Menú has a detector for orphan views.

## Dead files (do not use as reference)

`api.gs`, `api_config.gs`, `api_finance.gs`, `api_combined.gs`, `core_working.gs`, `analisis_working.gs`, `setup-sheets.gs`, `_diagnostico_catalogo.gs`, `revisar_catalogo.gs` are **not deployed** (absent from `$FILES`) and are superseded. Some header comments in live files still name them — those comments are stale.

## Design tokens

Colors, spacing, and typography are all CSS custom properties defined at `:root` in the `<style>` block. Dark mode is implemented via `[data-theme="dark"]` overrides. Brand color is dusty rose `#c46a7a`.

## CDN dependencies

- Chart.js 4.4.4 via jsDelivr
- Lucide icons via unpkg
- Satoshi font via fontshare
- SheetJS (xlsx) 0.18.5 via jsDelivr — lazy-loaded only when the Catálogo General "Importar Excel" button is used (see `_loadSheetJS()`)

All loaded from CDN — no local copies.
