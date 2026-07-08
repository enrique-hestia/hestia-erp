/* ==============================================================
   lab.gs — Módulo Laboratorio
   --------------------------------------------------------------
   ART Lab, FET, Andrología, Inventario Crío, Insumos Lab
   Proyecto Google Apps Script — Hestia Fertility ERP
   Todas las constantes vienen de config.gs (mismo proyecto)
   ============================================================== */

function readLabResumen(fechaInicio, fechaFin) {

  // ── CONFIGURACIÓN DE COLUMNAS ──────────────────────────────────
  // ART Lab: encabezados en fila 1 + fila 2 fusionadas (A=0, B=1…)
  // Ajusta estos índices si agregas/mueves columnas en la hoja
  var ART_COL_MES       = 0;   // A: Mes-Año
  var ART_COL_FECHA     = 1;   // B: Date
  var ART_COL_OOCITOS   = 9;   // J: # oocytes (recuperados)
  // Cuando conozcas las columnas de 2PN y blastocistos, actualiza:
  var ART_COL_2PN       = -1;  // -1 = no configurado aún
  var ART_COL_BLASTO    = -1;  // -1 = no configurado aún
  var ART_COL_ICSI_DANO = -1;  // -1 = no configurado aún

  // FET: encabezado en fila 2 (fila 1 vacía)
  var FET_COL_FECHA    = 0;  // A: Fecha
  var FET_COL_SURVIVED = 5;  // F: Survived ("Si"/"No")
  var FET_COL_BETA     = 6;  // G: Beta
  var FET_COL_PREG     = 7;  // H: Clinical Preg.

  // Inventario Crío: encabezado en fila 1
  var CRIO_COL_NOMBRE  = 0;  // A: Nombre paciente
  var CRIO_COL_FECHA   = 1;  // B: Fecha Crío
  var CRIO_COL_OOV     = 2;  // C: Oov (ovocitos)
  var CRIO_COL_EMB     = 3;  // D: Emb (embriones)

  // Insumos: encabezado en fila 2 (fila 1 vacía); col A vacía → datos desde col B
  var INS_COL_INSUMO   = 3;  // D: Insumo
  var INS_COL_PROV     = 4;  // E: Proveedor
  var INS_COL_FECHA    = 2;  // C: Fecha
  var INS_COL_COSTO    = 8;  // I: Costo
  // ─────────────────────────────────────────────────────────────

  function pct(num, den) {
    if (!den || den === 0) return null;
    return Math.round((num / den) * 1000) / 10;
  }
  function fmtFecha(v) {
    if (!v) return '';
    if (v instanceof Date) return fmtDate(v);
    var d = new Date(v);
    return isNaN(d.getTime()) ? String(v) : fmtDate(d);
  }
  function mesLabel(v) {
    if (!v) return '';
    if (v instanceof Date) return (v.getFullYear() + '-' + String(v.getMonth()+1).padStart(2,'0'));
    var s = String(v);
    var m = s.match(/(\d{4})-(\d{2})/);
    return m ? m[0] : s.slice(0,7);
  }

  try {
    var ssLab = SpreadsheetApp.openById(LAB_SS_ID);

    // ── 1. Leer ART Lab (encabezados fusionados fila 1+2, datos desde fila 3) ──
    var shArt  = findSheet(ssLab, 'ART Lab');
    var artRaw = shArt ? shArt.getDataRange().getValues() : [];
    // Fusionar headers de fila 1 y fila 2
    var artData = artRaw.slice(2).filter(function(r) {
      return r.some(function(c) { return String(c).trim() !== ''; });
    });
    var totalCiclos = artData.length;
    var totalOocitos = artData.reduce(function(s,r) { return s + (Number(r[ART_COL_OOCITOS])||0); }, 0);

    // % Fecundación y Blastocistos: calcular si las columnas están configuradas
    var total2PN   = ART_COL_2PN   > -1 ? artData.reduce(function(s,r){ return s+(Number(r[ART_COL_2PN])||0); },0) : null;
    var totalBlasto= ART_COL_BLASTO > -1 ? artData.reduce(function(s,r){ return s+(Number(r[ART_COL_BLASTO])||0); },0) : null;

    // Tendencia por mes desde ART Lab
    var artPorMes = {};
    artData.forEach(function(r) {
      var mes = mesLabel(r[ART_COL_FECHA] || r[ART_COL_MES]);
      if (!mes) return;
      if (!artPorMes[mes]) artPorMes[mes] = { ciclos:0, oocitos:0, pn2:0, blasto:0 };
      artPorMes[mes].ciclos++;
      artPorMes[mes].oocitos += Number(r[ART_COL_OOCITOS]) || 0;
      if (ART_COL_2PN    > -1) artPorMes[mes].pn2    += Number(r[ART_COL_2PN])    || 0;
      if (ART_COL_BLASTO > -1) artPorMes[mes].blasto += Number(r[ART_COL_BLASTO]) || 0;
    });
    var mesesArt = Object.keys(artPorMes).sort().slice(-6);

    // ── 2. Leer FET (encabezado en fila 2, datos desde fila 3) ──
    var shFet  = findSheet(ssLab, 'FET');
    var fetRaw = shFet ? shFet.getDataRange().getValues() : [];
    var fetData = fetRaw.slice(2).filter(function(r) {
      return r.some(function(c) { return String(c).trim() !== ''; });
    });
    var totalFet      = fetData.length;
    var fetSurvividos = fetData.filter(function(r) {
      return String(r[FET_COL_SURVIVED]).trim().toLowerCase() === 'si';
    }).length;
    var fetPregnancy  = fetData.filter(function(r) {
      return String(r[FET_COL_PREG]).trim().toLowerCase() === 'si';
    }).length;

    // Tendencia FET por mes
    var fetPorMes = {};
    fetData.forEach(function(r) {
      var mes = mesLabel(r[FET_COL_FECHA]);
      if (!mes) return;
      if (!fetPorMes[mes]) fetPorMes[mes] = { total:0, survived:0 };
      fetPorMes[mes].total++;
      if (String(r[FET_COL_SURVIVED]).trim().toLowerCase() === 'si') fetPorMes[mes].survived++;
    });

    // ── 3. Inventario Crío (encabezado fila 1, datos desde fila 2) ──
    var shCrio  = findSheet(ssLab, 'Inventario Crío');
    var crioRaw = shCrio ? shCrio.getDataRange().getValues() : [];
    var crioData = crioRaw.slice(1).filter(function(r) {
      return String(r[CRIO_COL_NOMBRE]).trim() !== '';
    });
    var totalOvCrio  = crioData.reduce(function(s,r){ return s+(Number(r[CRIO_COL_OOV])||0); },0);
    var totalEmbCrio = crioData.reduce(function(s,r){ return s+(Number(r[CRIO_COL_EMB])||0); },0);

    // ── 4. Insumos (encabezado fila 2, datos desde fila 3) ──
    var shIns  = findSheet(ssLab, 'Insumos');
    var insRaw = shIns ? shIns.getDataRange().getValues() : [];
    var insData = insRaw.slice(2).filter(function(r) {
      return String(r[INS_COL_INSUMO]).trim() !== '';
    });

    // ── Construir respuesta ──
    var fetPct = pct(fetSurvividos, totalFet);
    var fecPct = (ART_COL_2PN > -1) ? pct(total2PN, totalOocitos) : null;
    var blaPct = (ART_COL_BLASTO > -1) ? pct(totalBlasto, total2PN) : null;

    return {
      view:   'lab-resumen',
      fuente: 'lab-resumen',
      kpis: {
        fecundacion:      fecPct,
        blastocistos:     blaPct,
        fetSupervivencia: fetPct,
        icsiDano:         null,           // configurar ART_COL_ICSI_DANO
        totalCiclos:      totalCiclos,
        totalFet:         totalFet,
        fetPregnancy:     fetPregnancy,
        totalOvCrio:      totalOvCrio,
        totalEmbCrio:     totalEmbCrio,
        mesPeriodo:       mesesArt.length ? mesesArt[mesesArt.length-1] : ''
      },
      tendencia: {
        meses:        mesesArt,
        ciclos:       mesesArt.map(function(m){ return (artPorMes[m]||{}).ciclos||0; }),
        fecundacion:  mesesArt.map(function(m){
          var d = artPorMes[m]||{}; return ART_COL_2PN>-1 ? pct(d.pn2||0, d.oocitos||0) : null;
        }),
        fetSupervivencia: mesesArt.map(function(m){
          var d = fetPorMes[m]||{}; return pct(d.survived||0, d.total||0);
        })
      },
      tanques: crioData.slice(0,10).map(function(r) {
        return {
          nombre: String(r[CRIO_COL_NOMBRE]).split(' ').slice(0,2).join(' '),
          oocitos: Number(r[CRIO_COL_OOV])||0,
          embriones: Number(r[CRIO_COL_EMB])||0
        };
      }),
      insumos: insData.map(function(r) {
        return {
          item:        String(r[INS_COL_INSUMO]),
          proveedor:   String(r[INS_COL_PROV] || ''),
          fecha:       fmtFecha(r[INS_COL_FECHA]),
          costo:       Number(r[INS_COL_COSTO]) || 0,
          estado:      'ok'
        };
      })
    };
  } catch(ex) {
    return { view: 'lab-resumen', fuente: 'lab-resumen', error: ex.message,
             kpis: {}, tendencia: { meses:[], ciclos:[], fecundacion:[], fetSupervivencia:[] },
             tanques: [], insumos: [] };
  }
}

/* ══ DASHBOARD QUIROFANO ════════════════════════════════════════
   Placeholder — ampliar cuando se cree el spreadsheet de Qx
   ══════════════════════════════════════════════════════════════ */