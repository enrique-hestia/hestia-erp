/**
 * HESTIA FERTILITY — Setup inicial de hojas
 * Ejecuta esta función UNA sola vez para crear la estructura en Sheets.
 * Después puedes editar los datos directamente en las hojas.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── HOJA: Mensual_Todos ──────────────────────────────────────────────
  setupSheet(ss, 'Mensual_Todos', [
    ['Mes', 'Ingresos', 'Gastos', 'Ciclos', 'CAC', 'Margen'],
    ['Nov', 1800, 1150, 38, 54, 36],
    ['Dic', 1950, 1200, 41, 53, 38],
    ['Ene', 2100, 1280, 43, 52, 39],
    ['Feb', 2250, 1350, 45, 51, 38],
    ['Mar', 2180, 1330, 44, 52, 39],
    ['Abr', 2400, 1490, 47, 51, 38],
  ]);

  // ── HOJA: Mensual_Local ──────────────────────────────────────────────
  setupSheet(ss, 'Mensual_Local', [
    ['Mes', 'Ingresos', 'Gastos', 'Ciclos', 'CAC', 'Margen'],
    ['Nov', 1100, 730, 26, 42, 34],
    ['Dic', 1200, 750, 28, 41, 37],
    ['Ene', 1250, 800, 29, 40, 36],
    ['Feb', 1380, 850, 31, 40, 38],
    ['Mar', 1320, 820, 30, 41, 38],
    ['Abr', 1470, 920, 32, 40, 37],
  ]);

  // ── HOJA: Mensual_Internacional ──────────────────────────────────────
  setupSheet(ss, 'Mensual_Internacional', [
    ['Mes', 'Ingresos', 'Gastos', 'Ciclos', 'CAC', 'Margen'],
    ['Nov', 700, 420, 12, 72, 40],
    ['Dic', 750, 450, 13, 70, 40],
    ['Ene', 850, 480, 14, 68, 43],
    ['Feb', 870, 500, 14, 68, 43],
    ['Mar', 860, 510, 14, 70, 41],
    ['Abr', 930, 570, 15, 69, 39],
  ]);

  // ── HOJA: Servicios ──────────────────────────────────────────────────
  setupSheet(ss, 'Servicios', [
    ['Nombre', 'Color', 'Ingresos_Label', 'Margen_Pct', 'Meta_Pct'],
    ['FIV / ICSI',       '#c46a7a', '$1,080K', 42, 80],
    ['Congelamiento',    '#3d8f8f', '$528K',  38, 75],
    ['Inseminación',     '#c47a1e', '$336K',  29, 60],
    ['Estudios PGTa',   '#7a52b0', '$264K',  55, 90],
    ['Consultas',        '#3a72a8', '$192K',  72, 100],
  ]);

  // ── HOJA: Funnel ─────────────────────────────────────────────────────
  setupSheet(ss, 'Funnel', [
    ['Etapa', 'Valor', 'Pct', 'Color'],
    ['Prospectos',  310, 100, '#c46a7a'],
    ['1a Consulta', 190,  61, '#3d8f8f'],
    ['Diagnóstico', 120,  39, '#c47a1e'],
    ['Cotización',   85,  27, '#7a52b0'],
    ['Ciclo activo', 47,  15, '#3a72a8'],
  ]);

  // ── HOJA: Alertas ────────────────────────────────────────────────────
  setupSheet(ss, 'Alertas', [
    ['Tipo', 'Icono', 'Titulo', 'Descripcion'],
    ['warn', 'alert-triangle', 'CPP alto en FIV',       '62 días promedio — meta: < 50 días'],
    ['ok',   'check-circle',   'Margen PGTa óptimo',    '55% este mes, supera meta del 50%'],
    ['info', 'info',           '12 ciclos internacionales', 'Record histórico Q2 — turismo médico'],
  ]);

  // ── HOJA: DonutServicios ─────────────────────────────────────────────
  setupSheet(ss, 'DonutServicios', [
    ['Servicio', 'Porcentaje', 'Color'],
    ['FIV / ICSI',    45, '#c46a7a'],
    ['Congelamiento', 22, '#3d8f8f'],
    ['Inseminación',  14, '#c47a1e'],
    ['Estudios PGTa', 11, '#7a52b0'],
    ['Consultas',      8, '#3a72a8'],
  ]);

  // ── HOJA: CashFlow ───────────────────────────────────────────────────
  setupSheet(ss, 'CashFlow', [
    ['Mes', 'Flujo_MXN_K'],
    ['Nov', 420],
    ['Dic', 580],
    ['Ene', 650],
    ['Feb', 720],
    ['Mar', 680],
    ['Abr', 740],
  ]);

  // ── HOJA: DistribucionCostos ─────────────────────────────────────────
  setupSheet(ss, 'DistribucionCostos', [
    ['Categoria', 'Porcentaje', 'Color'],
    ['Nómina',      38, '#c46a7a99'],
    ['Renta',       15, '#3d8f8f99'],
    ['Medicamentos',20, '#c47a1e99'],
    ['Equipos',     10, '#7a52b099'],
    ['Marketing',   11, '#3a72a899'],
    ['Otros',        6, '#b8b4ae99'],
  ]);

  SpreadsheetApp.getUi().alert('✅ ¡Listo! Se crearon todas las hojas de Hestia Fertility.\n\nAhora puedes editar los datos directamente en cada hoja.\n\nSiguiente paso: implementar la función doGet como Web App.');
}

/** Crea o limpia una hoja y llena con datos */
function setupSheet(ss, name, rows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clearContents();
  }
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  // Formato: cabecera en negrita con fondo rosa suave
  const header = sheet.getRange(1, 1, 1, rows[0].length);
  header.setFontWeight('bold');
  header.setBackground('#fce8f0');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
}
