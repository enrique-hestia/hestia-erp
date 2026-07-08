/* ══════════════════════════════════════════════════════════════════════
   IMSS / INFONAVIT 2026  ·  Cuotas Obrero-Patronales  —  Hestia Clinic
   ----------------------------------------------------------------------
   Módulo autocontenido (sin dependencias). Pégalo en el <script> del
   dashboard. Misma lógica y parámetros que el papel de trabajo Excel.

   Uso rápido:
     var sbc = IMSS.integrarSBC({salarioDiario:500, despensaMes:1500}).sbcAplicable;
     var r   = IMSS.calcularCuotas(sbc, 30, 60);
     r.costoMensualPatron   // costo patronal mensualizado
   Base legal: LSS 25,28,30,71,106,107,147,168,211; LFT 76,80,87;
   reforma de pensiones DOF 16-dic-2020.
   ══════════════════════════════════════════════════════════════════════ */
var IMSS = (function () {

  /* ── Parámetros vigentes 2026 (editar cada año) ───────────────────── */
  var P = {
    UMA:        117.31,      // UMA diaria (vigente 01-feb-2026; ene-2026 = 113.14)
    SMG:        315.04,      // Salario Mínimo General diario, zona general 2026
    primaRT:    0.0054355,   // Prima de Riesgo de Trabajo Hestia (Y5486062101)

    // Mínimos de ley para el factor de integración (Art. 30 LSS)
    diasAguinaldo:  15,      // Art. 87 LFT
    diasVacaciones: 12,      // Art. 76 LFT (reforma 2023, 1er año)
    primaVacacional: 0.25,   // Art. 80 LFT

    // Exenciones para integrar el SBC (Art. 27 LSS)
    despensaExentaPctUMA: 0.40,  // exenta hasta 40% de la UMA diaria (frac. VI)
    premioExentoPctSDI:   0.10,  // c/premio exento hasta 10% del SDI  (frac. VII)

    // Cuotas MENSUALES (% sobre SBC, salvo cuota fija que va sobre UMA)
    mensual: {
      fija:      { p: 0.2040,  o: 0 },        // EyM cuota fija (×UMA×días) Art.106-I
      excedente: { p: 0.0110,  o: 0.0040 },   // EyM excedente 3 UMA       Art.106-II
      dinero:    { p: 0.0070,  o: 0.0025 },   // EyM prestaciones en dinero Art.107
      gmp:       { p: 0.0105,  o: 0.00375 },  // EyM gastos médicos pens.   Art.25
      iv:        { p: 0.0175,  o: 0.00625 },  // Invalidez y Vida           Art.147
      rt:        { p: null,    o: 0 },         // Riesgos de Trabajo = primaRT Art.71
      guarderia: { p: 0.0100,  o: 0 }          // Guarderías y Prest. Soc.   Art.211
    },

    // Cuotas BIMESTRALES (% sobre SBC)
    bimestral: {
      retiro:    { p: 0.0200, o: 0 },          // Retiro                Art.168-I
      ceav:      { p: null,   o: 0.01125 },    // Cesantía y Vejez (p = tabla) Art.168-II
      infonavit: { p: 0.0500, o: 0 }           // Vivienda              Art.29 LINFONAVIT
    },

    // Cesantía y Vejez — CUOTA PATRONAL 2026 por rango de SBC en veces UMA.
    // Reforma 2020 (transición 2023-2030). Salario mínimo => 3.150%.
    ceavPatron: [
      { umaMin: 0.00, pct: 0.03150 },  // hasta 1.00 (salario mínimo)
      { umaMin: 1.01, pct: 0.03676 },  // 1.01 a 1.50
      { umaMin: 1.51, pct: 0.04851 },  // 1.51 a 2.00
      { umaMin: 2.01, pct: 0.05556 },  // 2.01 a 2.50
      { umaMin: 2.51, pct: 0.06026 },  // 2.51 a 3.00
      { umaMin: 3.01, pct: 0.06361 },  // 3.01 a 3.50
      { umaMin: 3.51, pct: 0.06613 },  // 3.51 a 4.00
      { umaMin: 4.01, pct: 0.07513 }   // 4.01 en adelante
    ]
  };

  /* ── Helpers ──────────────────────────────────────────────────────── */
  function topeSBC()      { return P.UMA * 25; }          // Art. 28 LSS
  function limiteExc3()   { return P.UMA * 3;  }          // Art. 106-II
  function rtPct()        { return P.primaRT;  }
  var DIAS_MES = 30.4;                                    // factor mes→día (prestaciones)

  // Factor de integración (Art. 30 LSS frac. I)
  function factorIntegracion(diasAguinaldo, diasVacaciones, primaVac) {
    var a = (diasAguinaldo  != null ? diasAguinaldo  : P.diasAguinaldo);
    var v = (diasVacaciones != null ? diasVacaciones : P.diasVacaciones);
    var pv = (primaVac      != null ? primaVac       : P.primaVacacional);
    return 1 + a / 365 + (v * pv) / 365;
  }

  // % patronal de Cesantía y Vejez según el SBC
  function ceavPatronPct(sbc) {
    if (!sbc) return 0;
    if (sbc <= P.SMG) return P.ceavPatron[0].pct;   // salario mínimo => 3.150%
    var veces = sbc / P.UMA, pct = P.ceavPatron[0].pct;
    for (var i = 0; i < P.ceavPatron.length; i++) {
      if (veces >= P.ceavPatron[i].umaMin) pct = P.ceavPatron[i].pct;
    }
    return pct;
  }

  /* ── Integración del SBC desde sueldo + prestaciones ──────────────────
     Recibe montos MENSUALES de prestaciones. Aplica exenciones de Ley.
     args: { salarioDiario, diasAguinaldo, diasVacaciones, primaVac,
             despensaMes, premioPuntualidadMes, premioAsistenciaMes, otrasMes } */
  function integrarSBC(args) {
    args = args || {};
    var salarioDiario = Number(args.salarioDiario) || 0;
    var factor = factorIntegracion(args.diasAguinaldo, args.diasVacaciones, args.primaVac);
    var sdiFijo = salarioDiario * factor;

    var despDia  = (Number(args.despensaMes)           || 0) / DIAS_MES;
    var puntDia  = (Number(args.premioPuntualidadMes)  || 0) / DIAS_MES;
    var asisDia  = (Number(args.premioAsistenciaMes)   || 0) / DIAS_MES;
    var otrasDia = (Number(args.otrasMes)              || 0) / DIAS_MES;

    var despGrav = Math.max(0, despDia - P.despensaExentaPctUMA * P.UMA);   // frac. VI
    var puntGrav = Math.max(0, puntDia - P.premioExentoPctSDI   * sdiFijo); // frac. VII
    var asisGrav = Math.max(0, asisDia - P.premioExentoPctSDI   * sdiFijo); // frac. VII
    var otrasGrav = otrasDia;

    var sbc = sdiFijo + despGrav + puntGrav + asisGrav + otrasGrav;
    var sbcAplicable = sbc ? Math.min(sbc, topeSBC()) : 0;

    return {
      factor: factor, sdiFijo: sdiFijo,
      gravado: { despensa: despGrav, puntualidad: puntGrav, asistencia: asisGrav, otras: otrasGrav },
      sbc: sbc, sbcAplicable: sbcAplicable, topeSBC: topeSBC()
    };
  }

  /* ── Cálculo de cuotas a partir del SBC ───────────────────────────────
     sbc: salario base de cotización diario
     diasMes: días cotizados del mes (28-31)   · diasBim: días del bimestre (56-62) */
  function calcularCuotas(sbc, diasMes, diasBim) {
    sbc = Number(sbc) || 0;
    var dm = (diasMes != null ? diasMes : 30);
    var db = (diasBim != null ? diasBim : 60);
    var m = P.mensual, b = P.bimestral;
    var excBase = Math.max(0, sbc - limiteExc3());   // parte excedente de 3 UMA

    // --- MENSUALES ---
    var men = {
      fija:      { p: dm * P.UMA * m.fija.p,          o: 0 },
      excedente: { p: excBase * dm * m.excedente.p,   o: excBase * dm * m.excedente.o },
      dinero:    { p: sbc * dm * m.dinero.p,          o: sbc * dm * m.dinero.o },
      gmp:       { p: sbc * dm * m.gmp.p,             o: sbc * dm * m.gmp.o },
      iv:        { p: sbc * dm * m.iv.p,              o: sbc * dm * m.iv.o },
      rt:        { p: sbc * dm * rtPct(),             o: 0 },
      guarderia: { p: sbc * dm * m.guarderia.p,       o: 0 }
    };
    var mensualPatron = 0, mensualTrab = 0, k;
    for (k in men) { mensualPatron += men[k].p; mensualTrab += men[k].o; }

    // --- BIMESTRALES ---
    var ceavPct = ceavPatronPct(sbc);
    var bim = {
      retiro:    { p: sbc * db * b.retiro.p, o: 0 },
      ceav:      { p: sbc * db * ceavPct,    o: sbc * db * b.ceav.o, pct: ceavPct },
      infonavit: { p: sbc * db * b.infonavit.p, o: 0 }
    };
    var bimPatron = bim.retiro.p + bim.ceav.p + bim.infonavit.p;
    var bimTrab   = bim.ceav.o;

    return {
      sbc: sbc, diasMes: dm, diasBim: db,
      mensual:   { conceptos: men, patron: mensualPatron, trabajador: mensualTrab },
      bimestral: { conceptos: bim, patron: bimPatron,     trabajador: bimTrab },
      costoMensualPatron:        mensualPatron + bimPatron / 2,
      descuentoMensualTrabajador: mensualTrab + bimTrab / 2,
      get totalMensual() { return this.costoMensualPatron + this.descuentoMensualTrabajador; }
    };
  }

  /* ── Formato MXN (igual que el dashboard) ─────────────────────────── */
  function fmt(val) {
    return '$' + Number(val || 0).toLocaleString('es-MX',
      { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  return {
    params: P,
    topeSBC: topeSBC, limiteExc3: limiteExc3,
    factorIntegracion: factorIntegracion,
    ceavPatronPct: ceavPatronPct,
    integrarSBC: integrarSBC,
    calcularCuotas: calcularCuotas,
    fmt: fmt
  };
})();

/* ── Plantilla de empleados (datos reales del SUA, jun-2026) ──────────
   SBC diario tomado directo del SUA. Edita SBC para simular sueldos. */
var IMSS_EMPLEADOS = [
  { nombre: "ANA PAULA JASSO ARENAS",      nss: "03250121153", curp: "JAAA010917MDFSRNA0", sbc: 547.38 },
  { nombre: "NELIDA RODRIGUEZ OROZCO",     nss: "39927494482", curp: "ROON740502MDFDRL06", sbc: 725.77 },
  { nombre: "DANYA ITZEL ROJAS RODRIGUEZ", nss: "30119111240", curp: "RORD910924MDFJDN01", sbc: 609.77 },
  { nombre: "YAZMIN VELAZQUEZ VAZQUEZ",    nss: "11008208404", curp: "VEVY820827MDFLZZ06", sbc: 292.55 }
];

/* Ejemplo: total patronal mensual de toda la plantilla
   var totalMes = IMSS_EMPLEADOS.reduce(function (s, e) {
     return s + IMSS.calcularCuotas(e.sbc, 30, 60).costoMensualPatron;
   }, 0);
*/
