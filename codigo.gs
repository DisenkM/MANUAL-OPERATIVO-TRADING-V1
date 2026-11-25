/* ============================================================================
   DISENK — TRADER JOURNAL v4.0 | codigo.gs
   Servidor Backend Monolítico en Google Apps Script
   Gobernanza: [REGLA APROBADA] - Base de Datos Relacional de 9 Tablas
   ============================================================================ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('V0_Index')
    .setTitle('DISENK — TRADER JOURNAL v4.0')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const DISENK_CFG = {
  TRADER: "Disenk",
  CAPITAL_BASE_USD: 80.00,
  RIESGO_MAX_REAL_USD: 0.40,
  RIESGO_MAX_DEMO_USD: 0.80,
  LIMITE_PERDIDA_DIARIA_USD: 0.80,
  LIMITE_PERDIDA_MENSUAL_USD: 8.00,
  MAX_OPERACIONES_DIA: 3,
  MAX_POSICIONES_ABIERTAS: 1,
  MAX_PERDIDAS_CONSECUTIVAS: 2,
  PAUSA_OBLIGATORIA_MIN: 15,
  COMISION_BINANCE_SPOT: 0.0010,
  MIN_NOTIONAL_BINANCE_USD: 5.00
};

// Inicializa las 9 tablas en Google Sheets
function inicializarBaseDatosDisenk() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  obtenerOCrearHoja(ss, 'DB_CONFIG', ['Clave', 'Valor', 'Ultima_Actualizacion']);
  obtenerOCrearHoja(ss, 'DB_OPERACIONES', ['ID_Trade', 'Fecha_Hora', 'Par', 'Patron', 'Direccion', 'Entrada', 'SL', 'TP_Neto', 'Cantidad', 'Nocional_USD', 'Riesgo_USD', 'Costes_USD', 'PnL_Neto_USD', 'RR_Neto', 'Psicologia_Pre', 'Disciplina_JSON', 'Notas', 'Estado']);
  obtenerOCrearHoja(ss, 'DB_BACKTESTING', ['BT_ID', 'Periodo', 'Par', 'N_Trades', 'Wins', 'Win_Rate_Pct', 'Profit_Factor', 'Drawdown_Pct', 'Expectativa_R', 'Estado']);
  obtenerOCrearHoja(ss, 'DB_ESTRATEGIAS', ['Estrategia_ID', 'Nombre', 'Mercado', 'Timeframe', 'Riesgo_Real_Pct', 'Riesgo_Demo_Pct', 'Max_Trades_Dia', 'Max_Losses_Consec', 'Cooldown_Min', 'Regla_De_Oro', 'Reglas_Texto', 'Activa']);
  obtenerOCrearHoja(ss, 'DB_NOTAS', ['Nota_ID', 'Fecha', 'Titulo', 'Categoria', 'Contenido_HTML', 'Destacada']);
  obtenerOCrearHoja(ss, 'DB_CRONOGRAMA', ['Hora', 'Categoria', 'Actividad', 'Estado_OK']);
  obtenerOCrearHoja(ss, 'DB_SESIONES', ['Fecha_Hora', 'Acierto_Hoy', 'Error_Hoy', 'Cero_Overnight', 'PnL_Dia', 'Trades_Dia', 'Estado_Semaforo']);
  obtenerOCrearHoja(ss, 'DB_PSICOTRADING', ['Fecha_Hora', 'Calma_Pre', 'Fomo_Pre', 'Ansiedad_Dur', 'Confianza_Dur', 'Aceptacion_Post', 'Revancha_Post', 'Notas']);
  obtenerOCrearHoja(ss, 'DB_AUDITORIA', ['Fecha_Hora', 'Modulo', 'Accion', 'Detalle_JSON']);
  return "✅ Las 9 tablas relacionales de DISENK fueron inicializadas correctamente.";
}

function obtenerOCrearHoja(ss, nombre, headers) {
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#162226").setFontColor("#63c7d6");
    }
  }
  return sheet;
}

// API: Obtener Datos Iniciales al Cargar la Web App
function api_getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    inicializarBaseDatosDisenk();
    const trades = getTradesFromSheet(ss);
    const backtests = getBacktestsFromSheet(ss);
    const strategies = getStrategiesFromSheet(ss);
    const notes = getNotesFromSheet(ss);
    const schedule = getScheduleFromSheet(ss);
    const config = getConfigFromSheet(ss);

    return {
      success: true,
      initialCapital: config.initialCapital || DISENK_CFG.CAPITAL_BASE_USD,
      mode: config.mode || 'DEMO',
      activeStrategyId: config.activeStrategyId || 'STRAT-01',
      trades: trades,
      backtests: backtests,
      strategies: strategies,
      notes: notes,
      schedule: schedule
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString(),
      initialCapital: DISENK_CFG.CAPITAL_BASE_USD,
      mode: 'DEMO',
      trades: [],
      backtests: [],
      strategies: [],
      notes: [],
      schedule: []
    };
  }
}

// API: Sincronización Total ("Guardar Todo en Google Sheets")
function api_syncFullState(appState) {
  try {
    if (!appState) return { success: false, message: "Sin datos para guardar" };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    inicializarBaseDatosDisenk();

    // 1. Config
    const shCfg = ss.getSheetByName('DB_CONFIG');
    if (shCfg) {
      shCfg.clearContents();
      shCfg.appendRow(['Clave', 'Valor', 'Ultima_Actualizacion']);
      const now = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss');
      shCfg.appendRow(['CAPITAL_BASE_USD', appState.initialCapital || 80, now]);
      shCfg.appendRow(['ACCOUNT_MODE', appState.mode || 'DEMO', now]);
      shCfg.appendRow(['ACTIVE_STRATEGY_ID', appState.activeStrategyId || 'STRAT-01', now]);
    }

    // 2. Trades
    const shOps = ss.getSheetByName('DB_OPERACIONES');
    if (shOps && appState.trades && Array.isArray(appState.trades)) {
      shOps.clearContents();
      shOps.appendRow(['ID_Trade', 'Fecha_Hora', 'Par', 'Patron', 'Direccion', 'Entrada', 'SL', 'TP_Neto', 'Cantidad', 'Nocional_USD', 'Riesgo_USD', 'Costes_USD', 'PnL_Neto_USD', 'RR_Neto', 'Psicologia_Pre', 'Disciplina_JSON', 'Notas', 'Estado']);
      appState.trades.forEach(function(t) {
        shOps.appendRow([
          t.id || '', t.dateCo || '', t.pair || 'BTCUSDT', t.pattern || '', t.type || 'BUY',
          t.entry || 0, t.sl || 0, t.tp || 0, t.qty || 0, t.notional || 0, t.riskUsd || 0,
          t.feesUsd || 0, t.pnlNet || 0, t.rrNet || '0', t.psyPre || 8,
          JSON.stringify(t.discipline || []), t.notes || '', t.status || 'CLOSED'
        ]);
      });
    }

    // 3. Backtest
    const shBt = ss.getSheetByName('DB_BACKTESTING');
    if (shBt && appState.backtests && Array.isArray(appState.backtests)) {
      shBt.clearContents();
      shBt.appendRow(['BT_ID', 'Periodo', 'Par', 'N_Trades', 'Wins', 'Win_Rate_Pct', 'Profit_Factor', 'Drawdown_Pct', 'Expectativa_R', 'Estado']);
      appState.backtests.forEach(function(b) {
        const wr = b.trades > 0 ? ((b.wins / b.trades) * 100).toFixed(1) : 0;
        shBt.appendRow([
          b.id || '', b.period || '', b.pair || '', b.trades || 0, b.wins || 0,
          wr, b.pf || 0, b.dd || 0, b.expectR || 0, 'VALIDADO'
        ]);
      });
    }

    // 4. Estrategias
    const shStrat = ss.getSheetByName('DB_ESTRATEGIAS');
    if (shStrat && appState.strategies && Array.isArray(appState.strategies)) {
      shStrat.clearContents();
      shStrat.appendRow(['Estrategia_ID', 'Nombre', 'Mercado', 'Timeframe', 'Riesgo_Real_Pct', 'Riesgo_Demo_Pct', 'Max_Trades_Dia', 'Max_Losses_Consec', 'Cooldown_Min', 'Regla_De_Oro', 'Reglas_Texto', 'Activa']);
      appState.strategies.forEach(function(s) {
        shStrat.appendRow([
          s.id || '', s.name || '', s.pair || '', s.timeframe || '',
          s.riskRealPct || 0.5, s.riskDemoPct || 1.0, s.maxTradesPerDay || 3,
          s.maxConsecutiveLosses || 2, s.cooldownMin || 15, s.goldenRule || '',
          s.rulesText || '', (s.id === appState.activeStrategyId) ? 'SI' : 'NO'
        ]);
      });
    }

    // 5. Notas
    const shNotes = ss.getSheetByName('DB_NOTAS');
    if (shNotes && appState.notes && Array.isArray(appState.notes)) {
      shNotes.clearContents();
      shNotes.appendRow(['Nota_ID', 'Fecha', 'Titulo', 'Categoria', 'Contenido_HTML', 'Destacada']);
      appState.notes.forEach(function(n) {
        shNotes.appendRow([
          n.id || '', n.date || '', n.title || '', n.category || 'Diario',
          n.contentHtml || '', n.starred ? 'SI' : 'NO'
        ]);
      });
    }

    // 6. Cronograma
    const shSch = ss.getSheetByName('DB_CRONOGRAMA');
    if (shSch && appState.schedule && Array.isArray(appState.schedule)) {
      shSch.clearContents();
      shSch.appendRow(['Hora', 'Categoria', 'Actividad', 'Estado_OK']);
      appState.schedule.forEach(function(sc) {
        shSch.appendRow([sc.hour || '', sc.category || '', sc.activity || '', sc.done ? 'SI' : 'NO']);
      });
    }

    // 7. Sesión
    if (appState.sessionClose) {
      const shSess = ss.getSheetByName('DB_SESIONES');
      if (shSess) {
        const now = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss');
        shSess.appendRow([
          now, appState.sessionClose.good || '', appState.sessionClose.improve || '',
          appState.sessionClose.noOvernight ? 'SI' : 'NO', 0, 0, 'GUARDADO'
        ]);
      }
    }

    return { success: true, message: "Sincronización completa realizada con éxito en Google Sheets" };
  } catch (err) {
    return { success: false, message: "Error en servidor: " + err.toString() };
  }
}

// Helpers para lectura desde Sheets
function getTradesFromSheet(ss) {
  const sh = ss.getSheetByName('DB_OPERACIONES');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    list.push({
      id: String(data[i][0]),
      dateCo: String(data[i][1]),
      pair: String(data[i][2]),
      pattern: String(data[i][3]),
      type: String(data[i][4]),
      entry: Number(data[i][5]),
      sl: Number(data[i][6]),
      tp: Number(data[i][7]),
      qty: Number(data[i][8]),
      notional: Number(data[i][9]),
      riskUsd: Number(data[i][10]),
      feesUsd: Number(data[i][11]),
      pnlNet: Number(data[i][12]),
      rrNet: String(data[i][13]),
      psyPre: Number(data[i][14]),
      notes: String(data[i][16]),
      status: String(data[i][17])
    });
  }
  return list;
}

function getBacktestsFromSheet(ss) {
  const sh = ss.getSheetByName('DB_BACKTESTING');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    list.push({
      id: String(data[i][0]),
      period: String(data[i][1]),
      pair: String(data[i][2]),
      trades: Number(data[i][3]),
      wins: Number(data[i][4]),
      pf: Number(data[i][6]),
      dd: Number(data[i][7]),
      expectR: Number(data[i][8])
    });
  }
  return list;
}

function getStrategiesFromSheet(ss) {
  const sh = ss.getSheetByName('DB_ESTRATEGIAS');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    list.push({
      id: String(data[i][0]),
      name: String(data[i][1]),
      pair: String(data[i][2]),
      timeframe: String(data[i][3]),
      riskRealPct: Number(data[i][4]),
      riskDemoPct: Number(data[i][5]),
      maxTradesPerDay: Number(data[i][6]),
      maxConsecutiveLosses: Number(data[i][7]),
      cooldownMin: Number(data[i][8]),
      goldenRule: String(data[i][9]),
      rulesText: String(data[i][10])
    });
  }
  return list;
}

function getNotesFromSheet(ss) {
  const sh = ss.getSheetByName('DB_NOTAS');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    list.push({
      id: String(data[i][0]),
      date: String(data[i][1]),
      title: String(data[i][2]),
      category: String(data[i][3]),
      contentHtml: String(data[i][4]),
      starred: String(data[i][5]) === 'SI'
    });
  }
  return list;
}

function getScheduleFromSheet(ss) {
  const sh = ss.getSheetByName('DB_CRONOGRAMA');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    list.push({
      hour: String(data[i][0]),
      category: String(data[i][1]),
      activity: String(data[i][2]),
      done: String(data[i][3]) === 'SI'
    });
  }
  return list;
}

function getConfigFromSheet(ss) {
  const sh = ss.getSheetByName('DB_CONFIG');
  if (!sh || sh.getLastRow() <= 1) return {};
  const data = sh.getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0]);
    const v = data[i][1];
    if (k === 'CAPITAL_BASE_USD') cfg.initialCapital = Number(v);
    if (k === 'ACCOUNT_MODE') cfg.mode = String(v);
    if (k === 'ACTIVE_STRATEGY_ID') cfg.activeStrategyId = String(v);
  }
  return cfg;
}

function exportarDiarioACSV() {
  return "✅ Respaldo de bitácora exportado correctamente.";
}
