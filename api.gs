/**
 * HESTIA FERTILITY — API Web App
 * Implementar como: Aplicación web
 * Ejecutar como: Yo
 * Acceso: Cualquiera
 *
 * Esta función responde peticiones GET del dashboard HTML
 * con todos los datos del spreadsheet en formato JSON.
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById('1FMB2Qmv5z36sUDlVpwzjihNzrfS55k8MG32J04IBaR4');
    const result = {
      todos:         readMensual(ss, 'Mensual_Todos'),
      local:         readMensual(ss, 'Mensual_Local'),
      internacional: readMensual(ss, 'Mensual_Internacional'),
      servicios:     readServicios(ss),
      funnel:        readFunnel(ss),
      alertas:       readAlertas(ss),
      donut:         readDonut(ss),
      cashflow:      readCashFlow(ss),
      costos:        readCostos(ss),
      updated:       new Date().toISOString()
    };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** Lee una hoja mensual (Todos / Local / Internacional) */
function readMensual(ss, sheetName) {
  const rows = ss.getSheetByName(sheetName).getDataRange().getValues();
  const data = rows.slice(1); // quitar cabecera
  return {
    meses:    data.map(r => String(r[0])),
    ingresos: data.map(r => Number(r[1])),
    gastos:   data.map(r => Number(r[2])),
    ciclos:   data.map(r => Number(r[3])),
    cac:      data.map(r => Number(r[4])),
    margen:   data.map(r => Number(r[5])),
  };
}

/** Lee la hoja Servicios */
function readServicios(ss) {
  const rows = ss.getSheetByName('Servicios').getDataRange().getValues();
  return rows.slice(1).map(r => ({
    name:     String(r[0]),
    color:    String(r[1]),
    ingresos: String(r[2]),
    margen:   Number(r[3]),
    meta:     Number(r[4]),
  }));
}

/** Lee la hoja Funnel */
function readFunnel(ss) {
  const rows = ss.getSheetByName('Funnel').getDataRange().getValues();
  return rows.slice(1).map(r => ({
    label: String(r[0]),
    val:   Number(r[1]),
    pct:   Number(r[2]),
    color: String(r[3]),
  }));
}

/** Lee la hoja Alertas */
function readAlertas(ss) {
  const rows = ss.getSheetByName('Alertas').getDataRange().getValues();
  return rows.slice(1).map(r => ({
    type:  String(r[0]),
    icon:  String(r[1]),
    title: String(r[2]),
    desc:  String(r[3]),
  }));
}

/** Lee la hoja DonutServicios */
function readDonut(ss) {
  const rows = ss.getSheetByName('DonutServicios').getDataRange().getValues();
  const data = rows.slice(1);
  return {
    labels: data.map(r => String(r[0])),
    data:   data.map(r => Number(r[1])),
    colors: data.map(r => String(r[2])),
  };
}

/** Lee la hoja CashFlow */
function readCashFlow(ss) {
  const rows = ss.getSheetByName('CashFlow').getDataRange().getValues();
  const data = rows.slice(1);
  return {
    meses: data.map(r => String(r[0])),
    flujo: data.map(r => Number(r[1])),
  };
}

/** Lee la hoja DistribucionCostos */
function readCostos(ss) {
  const rows = ss.getSheetByName('DistribucionCostos').getDataRange().getValues();
  const data = rows.slice(1);
  return {
    labels: data.map(r => String(r[0])),
    data:   data.map(r => Number(r[1])),
    colors: data.map(r => String(r[2])),
  };
}
