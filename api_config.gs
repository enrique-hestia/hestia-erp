/* ══════════════════════════════════════════════════════════════
   api_config.gs — Constantes globales del ERP Hestia
   ──────────────────────────────────────────────────────────────
   INSTRUCCIONES: Copiar este archivo como un .gs separado dentro
   del mismo proyecto de Google Apps Script. Todas las constantes
   aquí son accesibles globalmente por api_core.gs y api_finance.gs
   ══════════════════════════════════════════════════════════════ */

var API_VERSION  = 'v2026-06-15';
var AUTH_SECRET  = 'hestia2026erp-secret'; // ← Cambia este valor por algo único y secreto

/* ── IDs de Spreadsheets ──────────────────────────────────── */
var SHEET_ID    = '1FMB2Qmv5z36sUDlVpwzjihNzrfS55k8MG32J04IBaR4'; // ERP principal
var ER_SS_ID    = '17jlXzaIvohpN_UoE2kvLK1Bb6P6JxyKONsnrDX_St9U'; // Estado de Resultados
var BANKS_SS_ID = '1O1tmtuVMlDl6rsN0IVFH14KmYjQZhY6GOs68PU_u0cg'; // Cuentas bancarias
var LAB_SS_ID   = '1hYmIl4gSTVrvghP7KY0y0dC200o8w0zShXj63zP-TrQ'; // Laboratorio
var MED_SS_ID   = '1fiuUtw-sg2ELNxq9bCjaOtRz1n87wuVi8IOQYzEi8tM'; // Medicamentos
var PAC_SS_ID   = '1uoQU-vbefxWwaLxJyTFT25gj7Nr2223WISa3tqH-Rio'; // Pacientes
var PROD_SS_ID  = '1eXskEMPdwuwEuV7GmVDNfyO1ulxhsZ9F_2hDVRDdIAY'; // Productos
var QX_SS_ID    = LAB_SS_ID; // ← Reemplaza cuando exista el sheet de Quirófano

/* ── GIDs (pestañas individuales por ID numérico) ─────────── */
var ER_GID     = 1953492149; // Estado de Resultados
var BUDGET_GID = 2097864117; // Budget
var PL_GID     = 1415550816; // Operating P&L Statement (referencia visual)

var BANKS_GID = {
  santander:   0,
  amex:        13958125,
  mercadopago: 1036684249
};

/* ── Mapeo: nombre de pestaña → spreadsheet donde vive ─────── */
var CAPTURA_SHEETS = {
  // Medicamentos
  'Medicamentos':    MED_SS_ID,
  'Orden_Compra':    MED_SS_ID,
  'Ent. Med':        MED_SS_ID,
  'Lista Med':       MED_SS_ID,
  'Estimulacion':    MED_SS_ID,
  'Estimulación':    MED_SS_ID,
  'Salidas Med':     MED_SS_ID,
  // Laboratorio
  'Resumen':         LAB_SS_ID,
  'ART Lab':         LAB_SS_ID,
  'FET':             LAB_SS_ID,
  'Andrología':      LAB_SS_ID,
  'Andrologia':      LAB_SS_ID,
  'Inventario Crío': LAB_SS_ID,
  'Inventario Crio': LAB_SS_ID,
  'Insumos':         LAB_SS_ID,
  'Salidas Lab':     LAB_SS_ID,
  // Quirófano
  'Insumos Qx':      QX_SS_ID,
  'Salidas Qx':      QX_SS_ID,
  // Otras
  'Pacientes':       PAC_SS_ID,
  'Productos':       PROD_SS_ID
  // ── FASE 2: agregar aquí CxC, CxP, Nómina, etc. ───────────
};

/* ── Aliases: nombre alternativo → nombre exacto en Sheets ─── */
var SHEET_ALIASES = {
  'Orden_Compra': 'Ent. Med',
  'Estimulacion': 'Estimulación'
};

var CAPTURA_SHEET_ID_DEFAULT = SHEET_ID;
