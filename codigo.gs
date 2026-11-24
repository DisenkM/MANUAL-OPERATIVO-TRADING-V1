// =========================================================================
// DISENK TRADER JOURNAL — BACKEND SERVER (Google Apps Script)
// =========================================================================

const DISENK_CFG = {
  CAPITAL_BASE_USD: 80.00,
  RIESGO_MAX_REAL_USD: 0.40,
  RIESGO_MAX_DEMO_USD: 0.80,
  LIMITE_DIARIO_USD: 0.80,
  LIMITE_MENSUAL_USD: 8.00
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('V0_Index')
    .evaluate()
    .setTitle('DISENK — TRADER JOURNAL | Consola Cuantitativa & Estrategia A v3.7')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function api_getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let initialCap = DISENK_CFG.CAPITAL_BASE_USD;
    let currentCap = DISENK_CFG.CAPITAL_BASE_USD;
    
    // Read DB_CONFIG
    const shCfg = ss.getSheetByName('DB_CONFIG');
    if (shCfg && shCfg.getLastRow() > 1) {
      const cfgData = shCfg.getDataRange().getValues();
      for (let i = 1; i < cfgData.length; i++) {
        if (cfgData[i][0] === 'CAPITAL_BASE_USD') {
          initialCap = Number(cfgData[i][1]) || DISENK_CFG.CAPITAL_BASE_USD;
        }
      }
    }

    let strategies = getStrategiesFromSheet(ss);
    let trades = getTradesFromSheet(ss);
    let backtests = getBacktestsFromSheet(ss);

    return {
      success: true,
      initialCapital: initialCap,
      strategies: strategies,
      trades: trades,
      backtests: backtests
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString(),
      initialCapital: DISENK_CFG.CAPITAL_BASE_USD,
      strategies: [],
      trades: [],
      backtests: []
    };
  }
}

function api_saveTrade(trade) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shOps = obtenerOCrearHoja(ss, 'DB_OPERACIONES', [
      'ID_Trade', 'Fecha_Hora', 'Par', 'Direccion', 'Entrada', 'SL', 'TP_Neto',
      'Cantidad_BTC', 'Nocional_USD', 'PnL_Neto_USD', 'PnL_Neto_COP', 'RR_Neto', 'Disciplina', 'Estado'
    ]);

    const idTrade = trade.id || 'TRD-' + Date.now();
    const fechaStr = trade.dateCo || Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss');
    
    shOps.appendRow([
      idTrade, fechaStr, trade.pair || 'BTCUSDT', trade.type || 'LONG',
      trade.entryPrice || trade.entry || 0, trade.stopLoss || trade.sl || 0, trade.takeProfit || trade.tp || 0,
      trade.quantity || 0, trade.positionSize || 0, trade.netPnL || trade.pnlNet || 0,
      (trade.netPnL || trade.pnlNet || 0) * 4000, trade.rMultiple || '1:2',
      typeof trade.discipline === 'string' ? trade.discipline : 'Ganadora con Disciplina',
      trade.status || 'CLOSED'
    ]);

    return { success: true, idTrade: idTrade };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function api_saveBacktest(bt) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shBt = obtenerOCrearHoja(ss, 'DB_BACKTESTING', [
      'ID_Backtest', 'Periodo', 'Par', 'Trades', 'WinRate', 'ProfitFactor', 'Drawdown', 'Expectativa_R'
    ]);

    shBt.appendRow([
      bt.id || ('BT-' + Date.now()), bt.period || '2026', bt.pair || 'BTCUSDT M5',
      bt.trades || 0, bt.wins ? ((bt.wins / bt.trades) * 100).toFixed(1) + '%' : '0%',
      bt.pf || 0, bt.dd || 0, bt.expectR || 0
    ]);

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function api_saveStrategy(st) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shSt = obtenerOCrearHoja(ss, 'DB_ESTRATEGIAS', [
      'ID_Estrategia', 'Nombre', 'Mercado', 'Timeframes', 'Riesgo_Real_Pct', 'Riesgo_Demo_Pct', 'Max_Trades_Dia', 'Cooldown_Min', 'Regla_Oro'
    ]);

    shSt.appendRow([
      st.id || ('STRAT-' + Date.now()), st.name || '', st.market || '', st.tf || '',
      st.riskReal || 0.5, st.riskDemo || 1.0, st.maxTrades || 3, st.cooldown || 15, st.goldRule || ''
    ]);

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function api_syncFullState(appState) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Sync DB_CONFIG
    if (appState && appState.initialCapital) {
      const shCfg = obtenerOCrearHoja(ss, 'DB_CONFIG', ['Clave', 'Valor', 'Fecha_Actualizacion']);
      shCfg.appendRow(['CAPITAL_BASE_USD', appState.initialCapital, new Date().toISOString()]);
    }

    // 2. Sync Trades
    if (appState && Array.isArray(appState.trades)) {
      appState.trades.forEach(function(t) { api_saveTrade(t); });
    }

    // 3. Sync Backtests
    if (appState && Array.isArray(appState.backtests)) {
      appState.backtests.forEach(function(b) { api_saveBacktest(b); });
    }

    // 4. Sync Strategies
    if (appState && Array.isArray(appState.strategies)) {
      appState.strategies.forEach(function(s) { api_saveStrategy(s); });
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

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
      type: String(data[i][3]),
      entry: Number(data[i][4]),
      sl: Number(data[i][5]),
      tp: Number(data[i][6]),
      quantity: Number(data[i][7]),
      positionSize: Number(data[i][8]),
      pnlNet: Number(data[i][9]),
      discipline: String(data[i][12]),
      status: String(data[i][13])
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
      pf: Number(data[i][5]),
      dd: Number(data[i][6]),
      expectR: Number(data[i][7])
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
      market: String(data[i][2]),
      tf: String(data[i][3]),
      riskReal: Number(data[i][4]),
      riskDemo: Number(data[i][5]),
      maxTrades: Number(data[i][6]),
      cooldown: Number(data[i][7]),
      goldRule: String(data[i][8])
    });
  }
  return list;
}

function inicializarBaseDatosDisenk() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tablas = [
    { nombre: 'DB_CONFIG', headers: ['Clave', 'Valor', 'Fecha_Actualizacion'] },
    { nombre: 'DB_OPERACIONES', headers: ['ID_Trade', 'Fecha_Hora', 'Par', 'Direccion', 'Entrada', 'SL', 'TP_Neto', 'Cantidad_BTC', 'Nocional_USD', 'PnL_Neto_USD', 'PnL_Neto_COP', 'RR_Neto', 'Disciplina', 'Estado'] },
    { nombre: 'DB_BACKTESTING', headers: ['ID_Backtest', 'Periodo', 'Par', 'Trades', 'WinRate', 'ProfitFactor', 'Drawdown', 'Expectativa_R'] },
    { nombre: 'DB_ESTRATEGIAS', headers: ['ID_Estrategia', 'Nombre', 'Mercado', 'Timeframes', 'Riesgo_Real_Pct', 'Riesgo_Demo_Pct', 'Max_Trades_Dia', 'Cooldown_Min', 'Regla_Oro'] },
    { nombre: 'DB_NOTAS', headers: ['ID_Nota', 'Fecha', 'Titulo', 'Categoria', 'Cuerpo'] },
    { nombre: 'DB_CRONOGRAMA', headers: ['Hora', 'Categoria', 'Tarea', 'Estado'] },
    { nombre: 'DB_SESIONES', headers: ['ID_Sesion', 'Fecha', 'Acierto', 'Error', 'Overnight', 'Reporte'] }
  ];

  tablas.forEach(function(t) { obtenerOCrearHoja(ss, t.nombre, t.headers); });
  return 'Base de datos Disenk inicializada con éxito.';
}

function obtenerOCrearHoja(ss, nombreHoja, headers) {
  let sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) {
    sheet = ss.insertSheet(nombreHoja);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#163840').setFontColor('#c7f0f7');
    sheet.setFrozenRows(1);
  }
  return sheet;
}