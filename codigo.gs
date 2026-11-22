/* ============================================================================
   DISENK — TRADER JOURNAL v4.0 (17 MÓDULOS UNIFICADOS)
   Servidor Backend Apps Script para V0_Index.html (Google Sheets DB)
   Gobernanza de Riesgo Inquebrantable | Zona Horaria: Colombia (UTC-5)
   ============================================================================ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('V0_Index')
    .setTitle('DISENK — TRADER JOURNAL | Consola Cuantitativa & 17 Módulos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Configuración Global Inquebrantable [REGLA APROBADA]
const DISENK_CFG = {
  TZ: "America/Bogota",
  CAPITAL_BASE_USD: 80.00,        // [REGLA APROBADA]
  RIESGO_REAL_PCT: 0.005,         // [REGLA APROBADA] 0.5%
  RIESGO_DEMO_PCT: 0.010,         // [REGLA APROBADA] 1.0%
  LIMITE_DIARIO_PCT: 0.010,       // [REGLA APROBADA] 1.0% del capital
  LIMITE_MENSUAL_PCT: 0.100,      // [REGLA APROBADA] 10.0% del capital
  MAX_TRADES_DIA: 3,               // [REGLA APROBADA]
  MAX_PERDIDAS_SEGUIDAS: 2,       // [REGLA APROBADA]
  PAUSA_OBLIGATORIA_MIN: 15,      // [REGLA APROBADA]
  COMISION_BINANCE_SPOT: 0.0010,   // [REGLA APROBADA] 0.10%
  MIN_NOTIONAL_BINANCE: 5.00,     // [REGLA APROBADA] $5.00 USDT
  SHEETS: {
    CONFIG: "DB_CONFIG",
    OPERACIONES: "DB_OPERACIONES",
    BACKTESTING: "DB_BACKTESTING",
    ESTRATEGIAS: "DB_ESTRATEGIAS",
    NOTAS: "DB_NOTAS",
    CRONOGRAMA: "DB_CRONOGRAMA",
    LEARN: "DB_LEARN",
    SESIONES: "DB_SESIONES",
    AUDITORIA: "DB_AUDITORIA"
  }
};

// 1. Inicialización de Base de Datos Relacional de 9 Tablas
function setupDisenkSheets() {
  return inicializarBaseDatosDisenk();
}

function inicializarBaseDatosDisenk() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemas = [
    { nombre: 'DB_CONFIG', headers: ['Clave', 'Valor', 'Etiqueta_Estado', 'Actualizado_UTC5'] },
    { nombre: 'DB_OPERACIONES', headers: ['ID_Trade', 'Fecha_Hora', 'Par', 'Direccion', 'Entrada', 'SL', 'TP_Neto', 'Cantidad_BTC', 'Nocional_USD', 'PnL_Neto_USD', 'PnL_Neto_COP', 'RR_Neto', 'Disciplina', 'Estado'] },
    { nombre: 'DB_BACKTESTING', headers: ['ID_Backtest', 'Periodo', 'Mercado', 'Trades', 'WinRate', 'ProfitFactor', 'Expectativa_R', 'MaxDD_Pct', 'Estado_Validacion'] },
    { nombre: 'DB_ESTRATEGIAS', headers: ['ID_Estrategia', 'Nombre', 'Mercado', 'TF', 'TargetRR', 'WinRate', 'Indicadores', 'Reglas', 'Gatillo'] },
    { nombre: 'DB_NOTAS', headers: ['ID_Nota', 'Titulo', 'Categoria', 'Fecha_Hora', 'Contenido_HTML'] },
    { nombre: 'DB_CRONOGRAMA', headers: ['Hora_Rango', 'Categoria', 'Descripcion', 'Cumplido'] },
    { nombre: 'DB_LEARN', headers: ['ID_Concepto', 'Titulo', 'Categoria', 'Explicacion', 'Tags'] },
    { nombre: 'DB_SESIONES', headers: ['ID_Sesion', 'Fecha_Hora', 'Modo', 'Capital_Base', 'PnL_Dia', 'Trades_Dia', 'Bloqueada'] },
    { nombre: 'DB_AUDITORIA', headers: ['ID_Audit', 'Timestamp_UTC5', 'Semaforo', 'Evento', 'Detalle_Regla'] }
  ];

  schemas.forEach(function(schema) {
    obtenerOCrearHoja(ss, schema.nombre, schema.headers);
  });

  return "🟢 Base de Datos Disenk v4.0 (9 Tablas DB_) inicializada correctamente en Google Sheets (UTC-5).";
}

function obtenerOCrearHoja(ss, nombre, headers) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1c2a30").setFontColor("#63c7d6");
  }
  return sh;
}

// 2. API: Carga Inicial de Datos desde Google Sheets
function api_getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Leer Capital Base de Config
    let initialCap = DISENK_CFG.CAPITAL_BASE_USD;
    const shCfg = ss.getSheetByName('DB_CONFIG');
    if (shCfg && shCfg.getLastRow() > 1) {
      const cfgData = shCfg.getDataRange().getValues();
      for (let i = 1; i < cfgData.length; i++) {
        if (String(cfgData[i][0]) === 'CAPITAL_BASE_USD') {
          initialCap = Number(cfgData[i][1]) || DISENK_CFG.CAPITAL_BASE_USD;
        }
      }
    }

    const trades = getTradesFromSheet(ss);
    const backtests = getBacktestsFromSheet(ss);
    const strategies = getStrategiesFromSheet(ss);
    const notes = getNotesFromSheet(ss);
    const schedule = getScheduleFromSheet(ss);
    const learnItems = getLearnFromSheet(ss);

    return {
      success: true,
      initialCapital: initialCap,
      trades: trades,
      backtests: backtests,
      strategies: strategies,
      notes: notes,
      schedule: schedule,
      learnItems: learnItems,
      config: DISENK_CFG
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString(),
      initialCapital: DISENK_CFG.CAPITAL_BASE_USD,
      trades: [], backtests: [], strategies: [], notes: [], schedule: [], learnItems: []
    };
  }
}

// Helper: Leer Trades
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
      qty: Number(data[i][7]),
      nocionalUsd: Number(data[i][8]),
      pnlNet: Number(data[i][9]),
      pnlNetCop: Number(data[i][10]),
      rMultiple: String(data[i][11]),
      discipline: String(data[i][12]),
      status: String(data[i][13])
    });
  }
  return list;
}

// Helper: Leer Backtests
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
      winRate: String(data[i][4]),
      pf: Number(data[i][5]),
      expectR: Number(data[i][6]),
      dd: Number(data[i][7])
    });
  }
  return list;
}

// Helper: Leer Estrategias
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
      targetRR: String(data[i][4]),
      winRate: String(data[i][5]),
      indicators: String(data[i][6]),
      rules: String(data[i][7]),
      trigger: String(data[i][8])
    });
  }
  return list;
}

// Helper: Leer Notas
function getNotesFromSheet(ss) {
  const sh = ss.getSheetByName('DB_NOTAS');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    list.push({
      id: String(data[i][0]),
      title: String(data[i][1]),
      category: String(data[i][2]),
      dateStr: String(data[i][3]),
      bodyHtml: String(data[i][4])
    });
  }
  return list;
}

// Helper: Leer Cronograma
function getScheduleFromSheet(ss) {
  const sh = ss.getSheetByName('DB_CRONOGRAMA');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    list.push({
      hour: String(data[i][0]),
      cat: String(data[i][1]),
      desc: String(data[i][2]),
      done: data[i][3] === true || String(data[i][3]).toLowerCase() === 'true'
    });
  }
  return list;
}

// Helper: Leer Learn
function getLearnFromSheet(ss) {
  const sh = ss.getSheetByName('DB_LEARN');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    list.push({
      id: String(data[i][0]),
      title: String(data[i][1]),
      cat: String(data[i][2]),
      text: String(data[i][3]),
      tags: String(data[i][4])
    });
  }
  return list;
}

// 3. API: Sincronización Total en Un Solo Proceso (Guardar Todo)
function api_syncFullState(appState) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Guardar Config Capital
    const shCfg = obtenerOCrearHoja(ss, 'DB_CONFIG', ['Clave', 'Valor', 'Etiqueta_Estado', 'Actualizado_UTC5']);
    const tsNow = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss');
    shCfg.getRange('A2:D2').setValues([['CAPITAL_BASE_USD', appState.initialCapital || 80.0, '[DATO DE USUARIO]', tsNow]]);

    // 2. Sincronizar Operaciones
    if (Array.isArray(appState.trades)) {
      const shOps = obtenerOCrearHoja(ss, 'DB_OPERACIONES', [
        'ID_Trade', 'Fecha_Hora', 'Par', 'Direccion', 'Entrada', 'SL', 'TP_Neto',
        'Cantidad_BTC', 'Nocional_USD', 'PnL_Neto_USD', 'PnL_Neto_COP', 'RR_Neto', 'Disciplina', 'Estado'
      ]);
      shOps.getRange('A2:N' + Math.max(2, shOps.getLastRow())).clearContent();
      const rowsOps = appState.trades.map(function(t) {
        return [
          t.id || ('TRD-' + Date.now()),
          t.dateCo || tsNow,
          t.pair || 'BTCUSDT',
          t.type || 'LONG',
          t.entry || 0,
          t.sl || 0,
          t.tp || 0,
          t.qty || 0,
          t.nocionalUsd || 0,
          t.pnlNet || 0,
          (t.pnlNet || 0) * 4000,
          t.rMultiple || '1:2',
          Array.isArray(t.discipline) ? t.discipline.join(',') : (t.discipline || '100% OK'),
          t.status || 'CLOSED'
        ];
      });
      if (rowsOps.length > 0) shOps.getRange(2, 1, rowsOps.length, 14).setValues(rowsOps);
    }

    // 3. Sincronizar Backtesting
    if (Array.isArray(appState.backtests)) {
      const shBt = obtenerOCrearHoja(ss, 'DB_BACKTESTING', [
        'ID_Backtest', 'Periodo', 'Mercado', 'Trades', 'WinRate', 'ProfitFactor', 'Expectativa_R', 'MaxDD_Pct', 'Estado_Validacion'
      ]);
      shBt.getRange('A2:I' + Math.max(2, shBt.getLastRow())).clearContent();
      const rowsBt = appState.backtests.map(function(b) {
        return [
          b.id || ('BT-' + Date.now()),
          b.period || '2026',
          b.pair || 'BTCUSDT M5',
          b.trades || 0,
          b.winRate || '60%',
          b.pf || 1.5,
          b.expectR || 0.4,
          b.dd || 2.0,
          (b.trades >= 100) ? 'VALIDADO 🟢' : 'INSUFICIENTE 🟡'
        ];
      });
      if (rowsBt.length > 0) shBt.getRange(2, 1, rowsBt.length, 9).setValues(rowsBt);
    }

    // 4. Sincronizar Estrategias
    if (Array.isArray(appState.strategies)) {
      const shStrat = obtenerOCrearHoja(ss, 'DB_ESTRATEGIAS', [
        'ID_Estrategia', 'Nombre', 'Mercado', 'TF', 'TargetRR', 'WinRate', 'Indicadores', 'Reglas', 'Gatillo'
      ]);
      shStrat.getRange('A2:I' + Math.max(2, shStrat.getLastRow())).clearContent();
      const rowsStrat = appState.strategies.map(function(s) {
        return [
          s.id || ('STRAT-' + Date.now()),
          s.name || 'Estrategia',
          s.market || 'Binance Spot',
          s.tf || '1H/5M/1M',
          s.targetRR || '1:2',
          s.winRate || '60%',
          s.indicators || 'EMA 50',
          s.rules || '',
          s.trigger || ''
        ];
      });
      if (rowsStrat.length > 0) shStrat.getRange(2, 1, rowsStrat.length, 9).setValues(rowsStrat);
    }

    // 5. Sincronizar Notas
    if (Array.isArray(appState.notes)) {
      const shNotes = obtenerOCrearHoja(ss, 'DB_NOTAS', ['ID_Nota', 'Titulo', 'Categoria', 'Fecha_Hora', 'Contenido_HTML']);
      shNotes.getRange('A2:E' + Math.max(2, shNotes.getLastRow())).clearContent();
      const rowsNotes = appState.notes.map(function(n) {
        return [
          n.id || ('NOTE-' + Date.now()),
          n.title || 'Sin Título',
          n.category || 'Reflexión',
          n.dateStr || tsNow,
          n.bodyHtml || ''
        ];
      });
      if (rowsNotes.length > 0) shNotes.getRange(2, 1, rowsNotes.length, 5).setValues(rowsNotes);
    }

    // 6. Sincronizar Cronograma
    if (Array.isArray(appState.schedule)) {
      const shSch = obtenerOCrearHoja(ss, 'DB_CRONOGRAMA', ['Hora_Rango', 'Categoria', 'Descripcion', 'Cumplido']);
      shSch.getRange('A2:D' + Math.max(2, shSch.getLastRow())).clearContent();
      const rowsSch = appState.schedule.map(function(sc) {
        return [
          sc.hour || '00:00',
          sc.cat || 'General',
          sc.desc || '',
          sc.done === true
        ];
      });
      if (rowsSch.length > 0) shSch.getRange(2, 1, rowsSch.length, 4).setValues(rowsSch);
    }

    // 7. Sincronizar Learn
    if (Array.isArray(appState.learnItems)) {
      const shLearn = obtenerOCrearHoja(ss, 'DB_LEARN', ['ID_Concepto', 'Titulo', 'Categoria', 'Explicacion', 'Tags']);
      shLearn.getRange('A2:E' + Math.max(2, shLearn.getLastRow())).clearContent();
      const rowsLearn = appState.learnItems.map(function(l) {
        return [
          l.id || ('LEARN-' + Date.now()),
          l.title || 'Concepto',
          l.cat || 'DICCIONARIO',
          l.text || '',
          l.tags || ''
        ];
      });
      if (rowsLearn.length > 0) shLearn.getRange(2, 1, rowsLearn.length, 5).setValues(rowsLearn);
    }

    return {
      success: true,
      timestamp: tsNow,
      message: "🟢 Todos los datos (17 Módulos) sincronizados con éxito en Google Sheets."
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// Exportar Respaldo CSV a Google Drive
function exportarDiarioACSV() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shOps = ss.getSheetByName('DB_OPERACIONES');
    if (!shOps) return 'Error: Hoja DB_OPERACIONES no existe';

    const data = shOps.getDataRange().getValues();
    const csvContent = data.map(function(r) {
      return r.map(function(c) {
        return '"' + String(c).replace(/"/g, '""') + '"';
      }).join(',');
    }).join('\n');

    const fileName = 'Disenk_Bitacora_Backup_' + Utilities.formatDate(new Date(), 'America/Bogota', 'yyyyMMdd_HHmmss') + '.csv';
    const file = DriveApp.createFile(fileName, csvContent, MimeType.CSV);

    return '✅ Respaldo exportado con éxito a Google Drive: ' + file.getName();
  } catch (err) {
    return '❌ Error al exportar a CSV: ' + err.toString();
  }
}
