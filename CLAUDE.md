# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

A single-page financial dashboard for Hestia Fertility clinic, built as one self-contained HTML file (`hestia-fertility-dashboard.html`). It fetches live data from Google Sheets via a Google Apps Script Web App and renders interactive charts using Chart.js.

There is no build step, no package manager, and no local server needed — open the HTML file in a browser or deploy it as a static file.

## Architecture

### Two-layer system

**Frontend** (`hestia-fertility-dashboard.html`) — A single HTML file containing all CSS (design tokens + component styles), HTML (all view panels), and JavaScript (data fetching, routing, chart rendering). Everything is inlined; no bundler, no imports.

**Backend** (`api.gs`) — A Google Apps Script deployed as a Web App. It reads data from a Google Sheets spreadsheet (`SHEET_ID = '1FMB2Qmv5z36sUDlVpwzjihNzrfS55k8MG32J04IBaR4'`) and returns a JSON payload via `doGet`. The deployed Web App URL is hardcoded in the HTML as `APPS_SCRIPT_URL`.

### Data flow

```
Browser → fetch(APPS_SCRIPT_URL?periodo=2026-Q2)
        → api.gs doGet()
        → reads Sheets tabs: Periodos, Mensual_Todos, Mensual_Local,
          Mensual_Internacional, CashFlow, Servicios, Funnel, Alertas,
          DonutServicios, DistribucionCostos, Inventarios, Laboratorios
        → returns JSON → sheetsData global → renderAll()
```

Auto-refresh runs every 30 seconds. If the API call fails, `getFallbackData()` provides hardcoded sample data.

### View router

Navigation is purely DOM-based: all views (`#view-resumen`, `#view-ingresos`, etc.) are rendered in the HTML and hidden/shown via `display:none`. `navigateTo(view)` is the single routing function. `renderView(view, data)` lazy-renders charts for non-default views on first visit.

### Global state

- `sheetsData` — cached API response (null until first load)
- `currentSegment` — `'todos'` | `'local'` | `'internacional'`
- `currentPeriodo` — active period ID (e.g. `'2026-Q2'`)
- Chart instances stored as module-level vars (`ingresoChart`, `donutChart`, etc.) and destroyed before re-creating to prevent Chart.js memory leaks

### Google Sheets structure

Each periodic data sheet (`Mensual_Todos`, `Mensual_Local`, `Mensual_Internacional`, `CashFlow`) has a `Periodo` column as the first column (e.g. `2026-Q2`). Capture sheets (`Inventarios`, `Laboratorios`) also have `Periodo` as column 0. Static reference sheets (`Servicios`, `Funnel`, `Alertas`, `DonutServicios`, `DistribucionCostos`) are not period-filtered.

## Deploying changes to the backend

After editing `api.gs`:
1. Open the Google Apps Script project linked to the spreadsheet
2. Deploy → Manage deployments → create a new version
3. Update `APPS_SCRIPT_URL` in `hestia-fertility-dashboard.html` if the deployment URL changed

## Initializing Sheets structure

Run `setupSheets()` in `api.gs` once from the Apps Script editor to create/reset all sheet tabs with headers and sample data. The older `setup-sheets.gs` is the original version without period-column support — `api.gs` contains the current authoritative version.

## Design tokens

Colors, spacing, and typography are all CSS custom properties defined at `:root` in the `<style>` block. Dark mode is implemented via `[data-theme="dark"]` overrides. Brand color is dusty rose `#c46a7a`.

## CDN dependencies

- Chart.js 4.4.4 via jsDelivr
- Lucide icons via unpkg
- Satoshi font via fontshare
- SheetJS (xlsx) 0.18.5 via jsDelivr — lazy-loaded only when the Catálogo General "Importar Excel" button is used (see `_loadSheetJS()`)

All loaded from CDN — no local copies.
