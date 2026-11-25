/* ============================================================================
   DISENK — TRADER JOURNAL v4.1 | codigo.gs
   Backend en Google Apps Script (America/Bogota UTC-5)
   Sincronización BIDIRECCIONAL entre la vista (appState) y Google Sheets.

   Funciones expuestas a la vista (google.script.run):
     - api_syncFullState(appState)  -> guarda TODO en Sheets (vista -> DB)
     - api_loadFullState()          -> lee TODO desde Sheets (DB -> vista)

   Uso: pega este archivo como codigo.gs en un proyecto ligado a tu hoja de
   cálculo (Extensiones > Apps Script), sube la vista como V0_Index.html y
   despliega como Aplicación web. Las 9 hojas se crean solas.
   ============================================================================ */

var TZ = 'America/Bogota';
var MAX_CELL = 49000;      // Sheets admite 50.000 caracteres por celda
var AUDIT_MAX_ROWS = 1000; // tope de filas en DB_AUDITORIA

/* ---------------------------------------------------------------------------
   ESQUEMA DE LAS 9 TABLAS
   text: índices (base 0) de columnas que se guardan como TEXTO. Evita que
   Sheets convierta "2026-11-21 09:15" o "06:00" en fechas/horas, o que un
   texto que empiece por "=" se interprete como fórmula.
   --------------------------------------------------------------------------- */
var SCHEMA = {
  DB_CONFIG: {
    headers: ['Clave', 'Valor', 'Ultima_Actualizacion'],
    text: [0, 1, 2]
  },
  DB_OPERACIONES: {
    headers: ['ID_Trade', 'Fecha_Hora', 'Par', 'Patron', 'Direccion', 'Entrada', 'SL', 'TP_Neto', 'Cantidad', 'Nocional_USD', 'Riesgo_USD', 'Costes_USD', 'PnL_Neto_USD', 'RR_Neto', 'Psicologia_Pre', 'Disciplina_JSON', 'Notas', 'Estado', 'Origen', 'Estrategia_ID', 'Mercado', 'Apalancamiento'],
    text: [0, 1, 2, 3, 4, 13, 15, 16, 17, 18, 19, 20]
  },
  DB_BACKTESTING: {
    headers: ['BT_ID', 'Periodo', 'Par', 'N_Trades', 'Wins', 'Win_Rate_Pct', 'Profit_Factor', 'Drawdown_Pct', 'Expectativa_R', 'Estado'],
    text: [0, 1, 2, 9]
  },
  DB_ESTRATEGIAS: {
    headers: ['Estrategia_ID', 'Nombre', 'Mercado', 'Timeframe', 'Riesgo_Real_Pct', 'Riesgo_Demo_Pct', 'Max_Trades_Dia', 'Max_Losses_Consec', 'Cooldown_Min', 'Regla_De_Oro', 'Reglas_Texto', 'Activa',
              'Tipo_Mercado', 'Instrumento', 'Direccion_Permitida', 'RR_Objetivo', 'Apalancamiento', 'Tam_Contrato', 'Paso_Lote', 'Lote_Min', 'Min_Nocional', 'Fee_Pct', 'Comision_Lote', 'Spread', 'Slippage_Pct', 'Valor_Quote_USD', 'MMR_Pct', 'Pip_Size', 'SL_Metodo', 'SL_ATR_Mult', 'SL_Pct', 'Usa_Filtro_EMA'],
    text: [0, 1, 2, 3, 9, 10, 11, 12, 13, 14, 28, 31]
  },
  DB_NOTAS: {
    headers: ['Nota_ID', 'Fecha', 'Titulo', 'Categoria', 'Contenido_HTML', 'Destacada'],
    text: [0, 1, 2, 3, 4, 5]
  },
  DB_CRONOGRAMA: {
    headers: ['Hora', 'Categoria', 'Actividad', 'Estado_OK'],
    text: [0, 1, 2, 3]
  },
  DB_SESIONES: {
    headers: ['Fecha_Cierre', 'Acierto', 'Mejora', 'Cero_Overnight', 'PnL_Dia', 'Trades_Dia', 'Estado'],
    text: [0, 1, 2, 3, 6]
  },
  DB_PSICOTRADING: {
    headers: ['Psy_ID', 'Calma_Pre', 'FOMO_Pre', 'Ansiedad_Dur', 'Confianza_SL', 'Aceptacion_Post', 'Revancha_Post', 'Nota'],
    text: [0, 7]
  },
  DB_AUDITORIA: {
    headers: ['Audit_ID', 'Timestamp_UTC5', 'Semaforo', 'Evento', 'Detalle'],
    text: [0, 1, 2, 3, 4]
  }
};

/* Campos "sueltos" del appState que no tienen tabla propia: se guardan como
   un único JSON en DB_CONFIG (clave EXTRA_STATE_JSON). */
var EXTRA_KEYS = ['prepChecks', 'prepPair', 'prepNews', 'prepBias', 'prepLevels',
                  'checklist15', 'stressChecks', 'psycho', 'universalCalc', 'sessionClose', 'tradeDir', 'autoSync', '_uSig'];

/* ============================================================================
   ENTRADA WEB + MENÚ
   ============================================================================ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('V0_Index')
    .setTitle('DISENK — TRADER JOURNAL v4.1')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DISENK')
    .addItem('Crear / validar hojas DB_*', 'setupDatabase')
    .addToUi();
}

/** Crea (o repara cabeceras de) las 9 hojas. Idempotente. */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function (name) { getSheet_(ss, name); });
  SpreadsheetApp.flush();
  return 'Hojas listas: ' + Object.keys(SCHEMA).join(', ');
}

/* ============================================================================
   VISTA -> SHEETS
   ============================================================================ */
function api_syncFullState(appState) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    if (!appState || typeof appState !== 'object') {
      return { success: false, message: 'Sin datos para guardar' };
    }
    lock.waitLock(20000); // evita escrituras simultáneas (doble clic, 2 pestañas)
    locked = true;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    var counts = {};
    var ctx = { truncated: 0 };

    /* --- DB_CONFIG (+ campos sueltos como JSON) --- */
    var extra = {};
    EXTRA_KEYS.forEach(function (k) { if (appState[k] !== undefined) extra[k] = appState[k]; });
    writeTable_(ss, 'DB_CONFIG', [
      ['CAPITAL_BASE_USD', String(num_(appState.initialCapital, 140)), now],
      ['ACCOUNT_MODE', String(appState.mode || 'DEMO'), now],
      ['ACTIVE_STRATEGY_ID', String(appState.activeStrategyId || 'STRAT-01'), now],
      ['EXTRA_STATE_JSON', clip_(JSON.stringify(extra), ctx), now]
    ]);

    /* --- DB_OPERACIONES --- */
    if (Array.isArray(appState.trades)) {
      writeTable_(ss, 'DB_OPERACIONES', appState.trades.map(function (t) {
        return [
          t.id || '', t.dateCo || '', t.pair || 'BTCUSDT', t.pattern || '', t.type || 'BUY',
          num_(t.entry, 0), num_(t.sl, 0), num_(t.tp, 0), num_(t.qty, 0), num_(t.notional, 0),
          num_(t.riskUsd, 0), num_(t.feesUsd, 0), num_(t.pnlNet, 0), String(t.rrNet != null ? t.rrNet : '0'),
          num_(t.psyPre, 8), JSON.stringify(t.discipline || []), clip_(t.notes, ctx),
          t.status || 'CLOSED', t.origin || 'USER',
          t.strategyId || 'STRAT-01', t.market || '', num_(t.lev, 1)
        ];
      }));
      counts.trades = appState.trades.length;
    }

    /* --- DB_BACKTESTING --- */
    if (Array.isArray(appState.backtests)) {
      writeTable_(ss, 'DB_BACKTESTING', appState.backtests.map(function (b) {
        var n = num_(b.trades, 0), w = num_(b.wins, 0);
        var wr = n > 0 ? Math.round((w / n) * 1000) / 10 : 0;
        return [b.id || '', b.period || '', b.pair || '', n, w, wr,
                num_(b.pf, 0), num_(b.dd, 0), num_(b.expectR, 0), 'VALIDADO'];
      }));
      counts.backtests = appState.backtests.length;
    }

    /* --- DB_ESTRATEGIAS --- */
    if (Array.isArray(appState.strategies)) {
      writeTable_(ss, 'DB_ESTRATEGIAS', appState.strategies.map(function (s) {
        return [s.id || '', s.name || '', s.pair || '', s.timeframe || '',
                num_(s.riskRealPct, 0.5), num_(s.riskDemoPct, 1.0), num_(s.maxTradesPerDay, 3),
                num_(s.maxConsecutiveLosses, 2), num_(s.cooldownMin, 15),
                clip_(s.goldenRule, ctx), clip_(s.rulesText, ctx),
                (s.id === appState.activeStrategyId) ? 'SI' : 'NO',
                s.market || 'SPOT', s.instrument || '', s.direction || 'LONG', num_(s.rr, 2), num_(s.lev, 1),
                num_(s.contractSize, 1), num_(s.lotStep, 0), num_(s.minQty, 0), num_(s.minNotional, 0),
                num_(s.feePct, 0), num_(s.commPerLot, 0), num_(s.spread, 0), num_(s.slipPct, 0),
                num_(s.quoteToUsd, 1), num_(s.mmrPct, 0), num_(s.pipSize, 1),
                s.slMethod || 'SWING_ATR', num_(s.slAtrMult, 0.1), num_(s.slPct, 0.5), s.useEmaFilter ? 'SI' : 'NO'];
      }));
      counts.strategies = appState.strategies.length;
    }

    /* --- DB_NOTAS --- */
    if (Array.isArray(appState.notes)) {
      writeTable_(ss, 'DB_NOTAS', appState.notes.map(function (n) {
        return [n.id || '', n.date || '', n.title || '', n.category || 'Diario',
                clip_(n.contentHtml, ctx), n.starred ? 'SI' : 'NO'];
      }));
      counts.notes = appState.notes.length;
    }

    /* --- DB_CRONOGRAMA --- */
    if (Array.isArray(appState.schedule)) {
      writeTable_(ss, 'DB_CRONOGRAMA', appState.schedule.map(function (sc) {
        return [sc.hour || '', sc.category || '', sc.activity || '', sc.done ? 'SI' : 'NO'];
      }));
      counts.schedule = appState.schedule.length;
    }

    /* --- DB_SESIONES: una fila por día (se actualiza, no se duplica) --- */
    var sc = appState.sessionClose;
    if (sc && (sc.good || sc.improve)) {
      var dayTrades = (appState.trades || []).filter(function (t) {
        return String(t.dateCo || '').slice(0, 10) === today;
      });
      var pnlDia = dayTrades.reduce(function (a, t) { return a + num_(t.pnlNet, 0); }, 0);
      upsertRow_(ss, 'DB_SESIONES', today, [
        today, clip_(sc.good, ctx), clip_(sc.improve, ctx), sc.noOvernight ? 'SI' : 'NO',
        Math.round(pnlDia * 100) / 100, dayTrades.length, 'GUARDADO'
      ]);
    }

    /* --- DB_PSICOTRADING: una fila por día --- */
    var p = appState.psycho;
    if (p && typeof p === 'object') {
      upsertRow_(ss, 'DB_PSICOTRADING', today, [
        today, num_(p.preCalm, 0), num_(p.preFomo, 0), num_(p.durAnx, 0), num_(p.durTrust, 0),
        num_(p.postAccept, 0), num_(p.postRevenge, 0), clip_(p.note, ctx)
      ]);
    }

    /* --- DB_AUDITORIA --- */
    var detalle = Object.keys(counts).map(function (k) { return k + '=' + counts[k]; }).join(', ');
    if (ctx.truncated) detalle += ' | celdas_truncadas=' + ctx.truncated;
    appendAudit_(ss, now, ctx.truncated ? 'AMARILLO' : 'VERDE', 'SYNC_FULL', detalle);

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Sincronización completa en Google Sheets (' + now + ')' +
               (ctx.truncated ? ' — aviso: ' + ctx.truncated + ' celda(s) superaron 49.000 caracteres y se recortaron.' : ''),
      savedAt: now
    };
  } catch (err) {
    return { success: false, message: 'Error en servidor: ' + err.toString() };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/* ============================================================================
   SHEETS -> VISTA
   Devuelve un objeto con las mismas claves de appState. La vista lo mezcla
   con Object.assign, así que NO pisa estado local (activeModule, cooldown...).
   ============================================================================ */
function api_loadFullState() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cfgRows = readTable_(ss, 'DB_CONFIG');
    var opsRows = readTable_(ss, 'DB_OPERACIONES');

    // Base vacía: no hay nada que cargar
    if (!cfgRows.length && !opsRows.length) {
      return { success: true, empty: true, message: 'La base de datos está vacía; guarda primero desde la vista.' };
    }

    var state = {};

    /* CONFIG */
    var cfg = {};
    cfgRows.forEach(function (r) { cfg[str_(r[0])] = r[1]; });
    if (cfg.CAPITAL_BASE_USD !== undefined && cfg.CAPITAL_BASE_USD !== '') state.initialCapital = num_(cfg.CAPITAL_BASE_USD, 140);
    if (cfg.ACCOUNT_MODE) state.mode = str_(cfg.ACCOUNT_MODE) === 'REAL' ? 'REAL' : 'DEMO';
    if (cfg.ACTIVE_STRATEGY_ID) state.activeStrategyId = str_(cfg.ACTIVE_STRATEGY_ID);
    if (cfg.EXTRA_STATE_JSON) {
      try {
        var extra = JSON.parse(str_(cfg.EXTRA_STATE_JSON));
        Object.keys(extra).forEach(function (k) { if (EXTRA_KEYS.indexOf(k) >= 0) state[k] = extra[k]; });
      } catch (e) { /* JSON dañado: se ignora, no rompe la carga */ }
    }

    /* OPERACIONES */
    state.trades = opsRows.map(function (r) {
      return {
        id: str_(r[0]), dateCo: str_(r[1]), pair: str_(r[2]) || 'BTCUSDT', pattern: str_(r[3]),
        type: str_(r[4]) || 'BUY', entry: num_(r[5], 0), sl: num_(r[6], 0), tp: num_(r[7], 0),
        qty: num_(r[8], 0), notional: num_(r[9], 0), riskUsd: num_(r[10], 0), feesUsd: num_(r[11], 0),
        pnlNet: num_(r[12], 0), rrNet: num_(r[13], 0).toFixed(2), psyPre: num_(r[14], 8),
        discipline: parseJson_(r[15], [true, true, true, true, true]),
        notes: str_(r[16]), status: str_(r[17]) || 'CLOSED', origin: str_(r[18]) || 'USER',
        strategyId: str_(r[19]) || 'STRAT-01', market: str_(r[20]) || 'SPOT', lev: num_(r[21], 1)
      };
    });

    /* BACKTESTING */
    state.backtests = readTable_(ss, 'DB_BACKTESTING').map(function (r) {
      return { id: str_(r[0]), period: str_(r[1]), pair: str_(r[2]), trades: num_(r[3], 0),
               wins: num_(r[4], 0), pf: num_(r[6], 0), dd: num_(r[7], 0), expectR: num_(r[8], 0) };
    });

    /* ESTRATEGIAS */
    var strat = readTable_(ss, 'DB_ESTRATEGIAS').map(function (r) {
      return { id: str_(r[0]), name: str_(r[1]), pair: str_(r[2]), timeframe: str_(r[3]),
               riskRealPct: num_(r[4], 0.5), riskDemoPct: num_(r[5], 1.0), maxTradesPerDay: num_(r[6], 3),
               maxConsecutiveLosses: num_(r[7], 2), cooldownMin: num_(r[8], 15),
               goldenRule: str_(r[9]), rulesText: str_(r[10]), _active: str_(r[11]) === 'SI',
               // columnas nuevas: si la hoja es antigua quedan en null y la vista usa el perfil por defecto
               market: str_(r[12]) || null, instrument: str_(r[13]) || null, direction: str_(r[14]) || null,
               rr: numN_(r[15]), lev: numN_(r[16]), contractSize: numN_(r[17]), lotStep: numN_(r[18]), minQty: numN_(r[19]),
               minNotional: numN_(r[20]), feePct: numN_(r[21]), commPerLot: numN_(r[22]), spread: numN_(r[23]), slipPct: numN_(r[24]),
               quoteToUsd: numN_(r[25]), mmrPct: numN_(r[26]), pipSize: numN_(r[27]),
               slMethod: str_(r[28]) || null, slAtrMult: numN_(r[29]), slPct: numN_(r[30]), useEmaFilter: str_(r[31]) === 'SI' };
    });
    var activa = strat.filter(function (s) { return s._active; })[0];
    if (activa && !state.activeStrategyId) state.activeStrategyId = activa.id;
    strat.forEach(function (s) { delete s._active; });
    if (strat.length) state.strategies = strat;

    /* NOTAS */
    var notes = readTable_(ss, 'DB_NOTAS').map(function (r) {
      return { id: str_(r[0]), date: str_(r[1]), title: str_(r[2]), category: str_(r[3]) || 'Diario',
               contentHtml: str_(r[4]), starred: str_(r[5]) === 'SI' };
    });
    if (notes.length) {
      state.notes = notes;
      state.activeNoteId = notes[0].id;
    }

    /* CRONOGRAMA */
    var sched = readTable_(ss, 'DB_CRONOGRAMA').map(function (r) {
      return { hour: str_(r[0]), category: str_(r[1]), activity: str_(r[2]), done: str_(r[3]) === 'SI' };
    });
    if (sched.length) state.schedule = sched;

    return {
      success: true,
      empty: false,
      state: state,
      loadedAt: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
      message: 'Datos cargados desde Google Sheets: ' + state.trades.length + ' operaciones, ' +
               (state.notes ? state.notes.length : 0) + ' notas.'
    };
  } catch (err) {
    return { success: false, message: 'Error en servidor: ' + err.toString() };
  }
}

/* ============================================================================
   HELPERS
   ============================================================================ */
function getSheet_(ss, name) {
  var def = SCHEMA[name];
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var n = def.headers.length;
  var hdr = sh.getRange(1, 1, 1, n);
  hdr.setValues([def.headers]).setFontWeight('bold').setBackground('#0b1a1f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

/** Reemplaza el contenido de una tabla en UNA sola escritura (rápido). */
function writeTable_(ss, name, rows) {
  var def = SCHEMA[name];
  var sh = getSheet_(ss, name);
  var nCols = def.headers.length;
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, nCols).clearContent();
  if (!rows.length) return;
  var fmts = rows.map(function () {
    var f = [];
    for (var c = 0; c < nCols; c++) f.push(def.text.indexOf(c) >= 0 ? '@' : 'General');
    return f;
  });
  var rng = sh.getRange(2, 1, rows.length, nCols);
  rng.setNumberFormats(fmts); // el formato de texto va ANTES de escribir
  rng.setValues(rows);
}

/** Lee las filas de datos (sin cabecera) y descarta filas vacías. */
function readTable_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var nCols = SCHEMA[name].headers.length;
  return sh.getRange(2, 1, last - 1, nCols).getValues().filter(function (r) {
    return String(r[0]).trim() !== '';
  });
}

/** Inserta o actualiza la fila cuya columna A coincide con key. */
function upsertRow_(ss, name, key, row) {
  var def = SCHEMA[name];
  var sh = getSheet_(ss, name);
  var nCols = def.headers.length;
  var last = sh.getLastRow();
  var target = last + 1;
  if (last > 1) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (keyStr_(keys[i][0]) === key) { target = i + 2; break; }
    }
  }
  var fmt = [];
  for (var c = 0; c < nCols; c++) fmt.push(def.text.indexOf(c) >= 0 ? '@' : 'General');
  var rng = sh.getRange(target, 1, 1, nCols);
  rng.setNumberFormats([fmt]);
  rng.setValues([row]);
}

function appendAudit_(ss, ts, semaforo, evento, detalle) {
  var sh = getSheet_(ss, 'DB_AUDITORIA');
  var row = ['AUD-' + Utilities.getUuid().slice(0, 8), ts, semaforo, evento, detalle];
  var r = sh.getRange(sh.getLastRow() + 1, 1, 1, 5);
  r.setNumberFormats([['@', '@', '@', '@', '@']]);
  r.setValues([row]);
  var total = sh.getLastRow() - 1;
  if (total > AUDIT_MAX_ROWS) sh.deleteRows(2, total - AUDIT_MAX_ROWS);
}

function keyStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v).trim();
}

function num_(v, def) {
  if (v === '' || v === null || v === undefined) return def;
  var n = Number(v);
  return isNaN(n) ? def : n;
}

/** Número o null si la celda está vacía (la vista rellena el valor por defecto). */
function numN_(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(v);
  return isNaN(n) ? null : n;
}

function str_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm');
  return String(v);
}

function clip_(v, ctx) {
  var s = (v === null || v === undefined) ? '' : String(v);
  if (s.length > MAX_CELL) {
    if (ctx) ctx.truncated++;
    return s.slice(0, MAX_CELL);
  }
  return s;
}

function parseJson_(v, fallback) {
  try {
    var o = JSON.parse(str_(v));
    return Array.isArray(o) && o.length ? o : fallback;
  } catch (e) { return fallback; }
}

