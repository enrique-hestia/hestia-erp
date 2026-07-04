/* ==============================================================
   config.gs — Constantes Globales del ERP Hestia
   --------------------------------------------------------------
   IDs de Spreadsheets, GIDs, mapeo de hojas, auth secret
   Proyecto Google Apps Script — Hestia Fertility ERP
   Todas las constantes vienen de config.gs (mismo proyecto)
   ============================================================== */

/* ══════════════════════════════════════════════════════════════
   api_config.gs — Constantes globales del ERP Hestia
   ──────────────────────────────────────────────────────────────
   INSTRUCCIONES: Copiar este archivo como un .gs separado dentro
   del mismo proyecto de Google Apps Script. Todas las constantes
   aquí son accesibles globalmente por api_core.gs y api_finance.gs
   ══════════════════════════════════════════════════════════════ */

var API_VERSION  = 'v2026-06-15';
// Secreto de firma de tokens. Vive SOLO en Script Properties
// (Configuración del proyecto → Propiedades del script → clave: AUTH_SECRET).
// Ya no hay literal en el código: si falta la propiedad, falla con un error claro.
var AUTH_SECRET  = (function() {
  var p = PropertiesService.getScriptProperties().getProperty('AUTH_SECRET');
  if (!p) throw new Error('Falta la propiedad de script AUTH_SECRET. Configúrala en: Configuración del proyecto → Propiedades del script.');
  return p;
})();

/* ── IDs de Spreadsheets ──────────────────────────────────── */
var SHEET_ID    = '1FMB2Qmv5z36sUDlVpwzjihNzrfS55k8MG32J04IBaR4'; // ERP principal
var ER_SS_ID    = '17jlXzaIvohpN_UoE2kvLK1Bb6P6JxyKONsnrDX_St9U'; // Estado de Resultados
var BANKS_SS_ID = '1O1tmtuVMlDl6rsN0IVFH14KmYjQZhY6GOs68PU_u0cg'; // Cuentas bancarias
var LAB_SS_ID   = '1hYmIl4gSTVrvghP7KY0y0dC200o8w0zShXj63zP-TrQ'; // Laboratorio
var MED_SS_ID   = '1fiuUtw-sg2ELNxq9bCjaOtRz1n87wuVi8IOQYzEi8tM'; // Medicamentos
var PAC_SS_ID   = '1uoQU-vbefxWwaLxJyTFT25gj7Nr2223WISa3tqH-Rio'; // Pacientes
var PROD_SS_ID  = '1eXskEMPdwuwEuV7GmVDNfyO1ulxhsZ9F_2hDVRDdIAY'; // Productos
var QX_SS_ID    = LAB_SS_ID; // ← Reemplaza cuando exista el sheet de Quirófano
var CAJA_CHICA_SS_ID = '1uB9HnQLqHbotP0w21z6mVQcc4ABb428iKXEru8hvDQE'; // Caja Chica
var CXP_SS_ID        = '1iRjpYtkcqx-3NRwlVK-UYx09I0gVyiTDRtIA9X9RAQw'; // Cuentas por Pagar
var CXP_GID          = 1448371071;
var MED_INV_SS_ID    = ''; // Inventario de Medicamentos — llenar tras correr setupInventarioMedicamentos()

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
  'Productos':       PROD_SS_ID,
  'Caja Chica':      CAJA_CHICA_SS_ID
  // ── FASE 2: agregar aquí CxC, CxP, Nómina, etc. ───────────
};

/* ── Aliases: nombre alternativo → nombre exacto en Sheets ─── */
var SHEET_ALIASES = {
  'Orden_Compra': 'Ent. Med',
  'Estimulacion': 'Estimulación'
};

var CAPTURA_SHEET_ID_DEFAULT = SHEET_ID;

/* ══════════════════════════════════════════════════════════════
   api_finance.gs — Módulo Financiero del ERP Hestia
   ──────────────────────────────────────────────────────────────
   Funciones:
     · readCashFlow       — Flujo mensual desde hoja CashFlow
     · readServicios      — Servicios y márgenes
     · readFunnel         — Embudo de conversión
     · readAlertas        — Alertas del dashboard
     · readDonut          — Distribución donut
     · readPaisesOrigen   — Turismo médico
     · readCostos         — Distribución de costos
     · readEstadoResultados — P&L detallado desde hoja ER
     · readOperatingPL    — Operating P&L Statement (stateless)
     · _buildPLReport     — Constructor interno del P&L
     · readBanksData      — Saldos y movimientos de 3 bancos
     · saveBankRow        — Escribe movimiento en pestaña bancaria
     · doPost             — Handler HTTP POST
   ══════════════════════════════════════════════════════════════ */

/* Columnas CashFlow: A=Sucursal | B=Fecha | C=Mes | D=Flujo_MXN */