
/**
 * CUADERNO DEL TRADER V5.3 — MVP
 * Backend Google Apps Script
 *
 * - Hojas aisladas con prefijo V53_
 * - No borra datos al inicializar
 * - Estrategias y reglas editables
 * - Versionado básico de estrategias
 * - Checklist PRE obligatorio para autorizar
 * - Checklist DURANTE y DESPUÉS registrables
 * - Cálculo de riesgo para Binance Spot LONG
 *
 * IMPORTANTE:
 * Este sistema NO envía órdenes a Binance.
 * No es una garantía de rentabilidad ni está validado
 * para operar con dinero real.
 */

const V53 = {
  VERSION: '5.3.0-MVP',
  PREFIX: 'V53_',
  SHEETS: {
    CONFIG: 'V53_CONFIG',
    STRATEGIES: 'V53_ESTRATEGIAS',
    RULES: 'V53_REGLAS',
    TRADES: 'V53_OPERACIONES',
    CHECKLISTS: 'V53_CHECKLISTS',
    LOG: 'V53_LOG'
  },
  DEFAULTS: {
    CAPITAL_USD: 80,
    RISK_REAL_PCT: 0.005,
    RISK_TEST_PCT: 0.01,
    FEE_RATE: 0.001,
    MIN_NOTIONAL_USD: 5,
    RR: 2,
    TIMEZONE: 'America/Bogota'
  }
};

/* =====================================================
   INICIO Y ESTRUCTURA
===================================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cuaderno V5.3')
    .addItem('Abrir Cuaderno', 'abrirCuaderno')
    .addSeparator()
    .addItem('Inicializar V5.3', 'setupV53')
    .addToUi();
}

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('V1_Index')
    .setTitle('Cuaderno del Trader V5.3')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function abrirCuaderno() {
  const html = HtmlService.createHtmlOutputFromFile('V1_Index')
    .setTitle('Cuaderno del Trader V5.3')
    .setWidth(1200)
    .setHeight(800);

  SpreadsheetApp.getUi().showModalDialog(html, 'Cuaderno del Trader V5.3');
}

function setupV53() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const schemas = {
    [V53.SHEETS.CONFIG]: [
      'CLAVE', 'VALOR', 'DESCRIPCION', 'ACTUALIZADO'
    ],
    [V53.SHEETS.STRATEGIES]: [
      'ID', 'NOMBRE', 'MERCADO', 'TIMEFRAME',
      'DIRECCION', 'VERSION', 'ESTADO', 'FECHA_CREACION'
    ],
    [V53.SHEETS.RULES]: [
      'ID_REGLA', 'ID_ESTRATEGIA', 'VERSION',
      'FASE', 'DESCRIPCION', 'TIPO',
      'OBLIGATORIA', 'ACTIVA', 'ORDEN'
    ],
    [V53.SHEETS.TRADES]: [
      'ID', 'FECHA_APERTURA', 'ESTRATEGIA_ID',
      'VERSION_ESTRATEGIA', 'SIMBOLO', 'DIRECCION',
      'CAPITAL_USD', 'RIESGO_PRESUPUESTADO',
      'ENTRADA', 'STOP_LOSS', 'TAKE_PROFIT',
      'CANTIDAD', 'ESTADO', 'FECHA_CIERRE',
      'SALIDA', 'COMISION_ENTRADA', 'COMISION_SALIDA',
      'PNL_BRUTO', 'PNL_NETO', 'RESULTADO_R', 'NOTAS'
    ],
    [V53.SHEETS.CHECKLISTS]: [
      'ID', 'TRADE_ID', 'FASE', 'ID_REGLA',
      'DESCRIPCION', 'CUMPLIDA', 'NOTA', 'FECHA'
    ],
    [V53.SHEETS.LOG]: [
      'FECHA', 'ACCION', 'DETALLE'
    ]
  };

  Object.keys(schemas).forEach(name => {
    let sh = ss.getSheetByName(name);

    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, schemas[name].length)
        .setValues([schemas[name]]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, schemas[name].length)
        .setFontWeight('bold');
    } else if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, schemas[name].length)
        .setValues([schemas[name]]);
      sh.setFrozenRows(1);
    }
  });

  insertarConfiguracionInicial_();
  insertarEstrategiaInicial_();
  registrarLog_('SETUP', 'Inicialización V5.3 ejecutada sin borrar datos.');

  return {
    ok: true,
    message: 'Estructura V5.3 inicializada. No se borraron hojas existentes.'
  };
}

function insertarConfiguracionInicial_() {
  const sh = hoja_(V53.SHEETS.CONFIG);
  const existentes = obtenerFilas_(sh);
  const claves = existentes.map(r => String(r[0]));

  const defaults = [
    ['CAPITAL_USD', 80, 'Capital de referencia en USD'],
    ['RIESGO_REAL_PCT', 0.005, 'Riesgo máximo por operación real'],
    ['RIESGO_TEST_PCT', 0.01, 'Riesgo máximo para pruebas'],
    ['FEE_RATE', 0.001, 'Comisión estimada por lado; verificar en Binance'],
    ['MIN_NOTIONAL_USD', 5, 'Mínimo de referencia editable; verificar filtros actuales'],
    ['RR', 2, 'Relación riesgo-beneficio objetivo'],
    ['TIMEZONE', 'America/Bogota', 'Zona horaria']
  ];

  defaults.forEach(row => {
    if (!claves.includes(row[0])) {
      sh.appendRow([...row, new Date()]);
    }
  });
}

function insertarEstrategiaInicial_() {
  const sh = hoja_(V53.SHEETS.STRATEGIES);

  if (sh.getLastRow() > 1) return;

  const id = 'EST-001';
  const fecha = new Date();

  sh.appendRow([
    id,
    'BTC Spot EMA50 + RSI14',
    'Binance Spot',
    'Por definir',
    'LONG',
    1,
    'BORRADOR',
    fecha
  ]);

  const reglas = [
    ['R001', id, 1, 'PRE', 'Confirmar contexto del mercado según las reglas de la estrategia.', 'MANUAL', true, true, 1],
    ['R002', id, 1, 'PRE', 'Confirmar RSI 14 según el umbral definido en la estrategia.', 'MANUAL', true, true, 2],
    ['R003', id, 1, 'PRE', 'Confirmar que el precio cumple la condición de EMA 50.', 'MANUAL', true, true, 3],
    ['R004', id, 1, 'PRE', 'Confirmar que la operación es LONG en Binance Spot.', 'MANUAL', true, true, 4],
    ['R005', id, 1, 'PRE', 'Verificar que el stop loss está por debajo de la entrada.', 'AUTO', true, true, 5],
    ['R006', id, 1, 'PRE', 'Verificar que el riesgo calculado no supera el presupuesto.', 'AUTO', true, true, 6],
    ['R007', id, 1, 'DURANTE', 'Respetar el plan y no mover el stop por impulso.', 'MANUAL', true, true, 1],
    ['R008', id, 1, 'DURANTE', 'No aumentar el tamaño de la posición durante la operación.', 'MANUAL', true, true, 2],
    ['R009', id, 1, 'DESPUES', 'Registrar el precio real de salida y las comisiones.', 'MANUAL', true, true, 1],
    ['R010', id, 1, 'DESPUES', 'Registrar si se respetó el plan de trading.', 'MANUAL', true, true, 2],
    ['R011', id, 1, 'DESPUES', 'Escribir una observación de aprendizaje.', 'MANUAL', true, true, 3]
  ];

  const rulesSheet = hoja_(V53.SHEETS.RULES);
  rulesSheet.getRange(
    rulesSheet.getLastRow() + 1,
    1,
    reglas.length,
    reglas[0].length
  ).setValues(reglas);
}

/* =====================================================
   API PARA LA INTERFAZ
===================================================== */

function api_getInitialData() {
  asegurarSetup_();

  return {
    version: V53.VERSION,
    config: obtenerConfig_(),
    strategies: obtenerEstrategias_(),
    rules: obtenerReglas_(),
    trades: obtenerOperaciones_()
  };
}

function api_saveConfig(data) {
  asegurarSetup_();

  const permitidas = [
    'CAPITAL_USD',
    'RIESGO_REAL_PCT',
    'RIESGO_TEST_PCT',
    'FEE_RATE',
    'MIN_NOTIONAL_USD',
    'RR'
  ];

  Object.keys(data || {}).forEach(key => {
    if (!permitidas.includes(key)) return;

    const value = Number(data[key]);
    if (!Number.isFinite(value) || value <= 0) {
      throw new Error('Valor inválido para ' + key);
    }

    guardarConfigValor_(key, value);
  });

  registrarLog_('CONFIG', 'Configuración actualizada.');
  return { ok: true, config: obtenerConfig_() };
}

function api_createStrategy(data) {
  asegurarSetup_();

  const nombre = String(data.nombre || '').trim();
  if (!nombre) throw new Error('Escribe el nombre de la estrategia.');

  const sh = hoja_(V53.SHEETS.STRATEGIES);
  const id = 'EST-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  sh.appendRow([
    id,
    nombre,
    data.mercado || 'Binance Spot',
    data.timeframe || 'Por definir',
    'LONG',
    1,
    'BORRADOR',
    new Date()
  ]);

  registrarLog_('ESTRATEGIA', 'Creada ' + id);

  return { ok: true, id: id };
}

function api_addRule(data) {
  asegurarSetup_();

  const strategyId = String(data.strategyId || '');
  const fase = String(data.fase || '').toUpperCase();
  const descripcion = String(data.descripcion || '').trim();

  if (!strategyId || !descripcion) {
    throw new Error('Falta estrategia o descripción.');
  }

  if (!['PRE', 'DURANTE', 'DESPUES'].includes(fase)) {
    throw new Error('Fase inválida.');
  }

  const strategy = obtenerEstrategias_()
    .find(s => s.ID === strategyId);

  if (!strategy) throw new Error('Estrategia no encontrada.');

  const sh = hoja_(V53.SHEETS.RULES);
  const id = 'R-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  sh.appendRow([
    id,
    strategyId,
    Number(strategy.VERSION),
    fase,
    descripcion,
    'MANUAL',
    data.obligatoria !== false,
    true,
    sh.getLastRow()
  ]);

  registrarLog_('REGLA', 'Agregada ' + id);
  return { ok: true, id: id };
}

function api_deactivateRule(ruleId) {
  asegurarSetup_();

  const sh = hoja_(V53.SHEETS.RULES);
  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(ruleId)) {
      sh.getRange(i + 1, 8).setValue(false);

      const strategyId = values[i][1];
      incrementarVersionEstrategia_(strategyId);

      registrarLog_('REGLA', 'Desactivada ' + ruleId);
      return { ok: true };
    }
  }

  throw new Error('Regla no encontrada.');
}

/* =====================================================
   CÁLCULO DE POSICIÓN
===================================================== */

function api_calculatePosition(data) {
  asegurarSetup_();

  const config = obtenerConfig_();

  const capital = Number(data.capital || config.CAPITAL_USD);
  const entry = Number(data.entry);
  const stop = Number(data.stop);
  const rr = Number(data.rr || config.RR);
  const fee = Number(config.FEE_RATE);
  const riskPct = Number(data.riskPct || config.RIESGO_REAL_PCT);
  const minNotional = Number(config.MIN_NOTIONAL_USD);

  if (![capital, entry, stop, rr, fee, riskPct].every(Number.isFinite)) {
    throw new Error('Hay valores numéricos inválidos.');
  }

  if (capital <= 0 || entry <= 0 || stop <= 0 || rr <= 0) {
    throw new Error('Los valores deben ser mayores que cero.');
  }

  if (stop >= entry) {
    throw new Error('En Spot LONG, el Stop Loss debe estar por debajo de la entrada.');
  }

  const maxRisk = data.mode === 'TEST'
    ? Number(config.RIESGO_TEST_PCT)
    : Number(config.RIESGO_REAL_PCT);

  if (riskPct <= 0 || riskPct > maxRisk) {
    throw new Error(
      'El riesgo supera el límite permitido para esta modalidad.'
    );
  }
  const riskBudget = capital * riskPct;
  const distance = entry - stop;

  // Incluye comisión estimada de entrada y salida.
  const costPerUnit = distance + (entry + stop) * fee;

  let quantity = riskBudget / costPerUnit;

  // Redondeo conservador a 6 decimales.
  quantity = Math.floor(quantity * 1e6) / 1e6;

  const notional = quantity * entry;
  const estimatedRisk = quantity * costPerUnit;

  if (notional < minNotional) {
    return {
      ok: false,
      blocked: true,
      reason: 'El valor nominal queda por debajo del mínimo de referencia configurado.',
      minNotional: minNotional,
      notional: notional,
      riskBudget: riskBudget
    };
  }

  if (estimatedRisk > riskBudget + 0.000001) {
    throw new Error('El riesgo final supera el presupuesto después del redondeo.');
  }

  const takeProfit = entry + (distance * rr);
  const grossPotential = (takeProfit - entry) * quantity;
  const estimatedFees = (entry + takeProfit) * quantity * fee;
  const netPotential = grossPotential - estimatedFees;

  return {
    ok: true,
    capital: capital,
    riskPct: riskPct,
    riskBudget: riskBudget,
    entry: entry,
    stop: stop,
    takeProfit: takeProfit,
    quantity: quantity,
    notional: notional,
    estimatedRisk: estimatedRisk,
    grossPotential: grossPotential,
    estimatedFees: estimatedFees,
    netPotential: netPotential,
    rr: rr
  };
}

/* =====================================================
   AUTORIZACIÓN PREVIA
===================================================== */

function api_authorizeTrade(data) {
  asegurarSetup_();

  const strategyId = String(data.strategyId || '');
  const strategy = obtenerEstrategias_()
    .find(s => s.ID === strategyId);

  if (!strategy) throw new Error('Estrategia no encontrada.');

  if (String(data.direction || 'LONG').toUpperCase() !== 'LONG') {
    throw new Error('Esta versión solo permite operaciones LONG en Binance Spot.');
  }

  const rules = obtenerReglas_().filter(r =>
    r.ID_ESTRATEGIA === strategyId &&
    Number(r.VERSION) === Number(strategy.VERSION) &&
    r.FASE === 'PRE' &&
    r.ACTIVA === true &&
    r.OBLIGATORIA === true
  );

  const answers = data.answers || {};
  const missing = [];

  rules.forEach(rule => {
    if (rule.TIPO === 'AUTO') {
      if (rule.ID_REGLA === 'R005' && !(Number(data.stop) < Number(data.entry))) {
        missing.push(rule.DESCRIPCION);
      }
      if (rule.ID_REGLA === 'R006' &&
          !(Number(data.calculatedRisk) <= Number(data.riskBudget))) {
        missing.push(rule.DESCRIPCION);
      }
    } else if (answers[rule.ID_REGLA] !== true) {
      missing.push(rule.DESCRIPCION);
    }
  });

  if (missing.length) {
    return {
      ok: false,
      authorized: false,
      missing: missing
    };
  }

  const calc = api_calculatePosition({
    capital: data.capital,
    entry: data.entry,
    stop: data.stop,
    rr: data.rr,
    riskPct: data.riskPct,
    mode: data.mode
  });

  if (!calc.ok) {
    return {
      ok: false,
      authorized: false,
      missing: [calc.reason]
    };
  }

  const tradeId = 'TRD-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  const sh = hoja_(V53.SHEETS.TRADES);

  sh.appendRow([
    tradeId,
    new Date(),
    strategyId,
    Number(strategy.VERSION),
    data.symbol || 'BTCUSDT',
    'LONG',
    calc.capital,
    calc.riskBudget,
    calc.entry,
    calc.stop,
    calc.takeProfit,
    calc.quantity,
    'ABIERTA',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    ''
  ]);

  guardarChecklist_(tradeId, strategyId, strategy.VERSION, 'PRE', answers);

  registrarLog_('TRADE_AUTORIZADO', tradeId);

  return {
    ok: true,
    authorized: true,
    tradeId: tradeId,
    calculation: calc
  };
}

/* =====================================================
   CHECKLIST DURANTE Y DESPUÉS
===================================================== */

function api_saveChecklist(data) {
  asegurarSetup_();

  const tradeId = String(data.tradeId || '');
  const fase = String(data.fase || '').toUpperCase();

  if (!['DURANTE', 'DESPUES'].includes(fase)) {
    throw new Error('Fase no válida.');
  }

  const trade = obtenerOperaciones_().find(t => t.ID === tradeId);
  if (!trade) throw new Error('Operación no encontrada.');

  guardarChecklist_(
    tradeId,
    trade.ESTRATEGIA_ID,
    trade.VERSION_ESTRATEGIA,
    fase,
    data.answers || {},
    data.notes || {}
  );

  registrarLog_('CHECKLIST_' + fase, tradeId);
  return { ok: true };
}

function guardarChecklist_(tradeId, strategyId, version, fase, answers, notes) {
  const rules = obtenerReglas_().filter(r =>
    r.ID_ESTRATEGIA === strategyId &&
    Number(r.VERSION) === Number(version) &&
    r.FASE === fase &&
    r.ACTIVA === true
  );

  const sh = hoja_(V53.SHEETS.CHECKLISTS);

  rules.forEach(rule => {
    const answer = answers[rule.ID_REGLA];

    // Se registra true o false; no se fuerza una respuesta positiva.
    if (answer === true || answer === false) {
      sh.appendRow([
        Utilities.getUuid(),
        tradeId,
        fase,
        rule.ID_REGLA,
        rule.DESCRIPCION,
        answer,
        notes[rule.ID_REGLA] || '',
        new Date()
      ]);
    }
  });
}

/* =====================================================
   CIERRE DE OPERACIÓN
===================================================== */

function api_closeTrade(data) {
  asegurarSetup_();

  const tradeId = String(data.tradeId || '');
  const exit = Number(data.exit);
  const exitFee = Number(data.exitFee || 0);

  if (!tradeId || !Number.isFinite(exit) || exit <= 0) {
    throw new Error('Datos de cierre inválidos.');
  }

  const sh = hoja_(V53.SHEETS.TRADES);
  const values = sh.getDataRange().getValues();
  let row = -1;
  let trade = null;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === tradeId) {
      row = i + 1;
      trade = values[i];
      break;
    }
  }

  if (!trade) throw new Error('Operación no encontrada.');
  if (trade[12] !== 'ABIERTA') throw new Error('La operación no está abierta.');

  const quantity = Number(trade[11]);
  const entry = Number(trade[8]);
  const entryFee = Number(data.entryFee || 0);
  const gross = (exit - entry) * quantity;
  const net = gross - entryFee - exitFee;
  const risk = Number(trade[7]);
  const resultR = risk > 0 ? net / risk : 0;

  sh.getRange(row, 13).setValue('CERRADA');
  sh.getRange(row, 14).setValue(new Date());
  sh.getRange(row, 15).setValue(exit);
  sh.getRange(row, 16).setValue(entryFee);
  sh.getRange(row, 17).setValue(exitFee);
  sh.getRange(row, 18).setValue(gross);
  sh.getRange(row, 19).setValue(net);
  sh.getRange(row, 20).setValue(resultR);
  sh.getRange(row, 21).setValue(data.notes || '');

  registrarLog_('TRADE_CERRADO', tradeId + ' | PnL neto: ' + net);

  return {
    ok: true,
    tradeId: tradeId,
    gross: gross,
    net: net,
    resultR: resultR
  };
}

/* =====================================================
   CONSULTAS
===================================================== */

function obtenerConfig_() {
  const sh = hoja_(V53.SHEETS.CONFIG);
  const rows = obtenerFilas_(sh);
  const config = {};

  rows.forEach(r => {
    config[String(r[0])] = r[1];
  });

  return config;
}

function obtenerEstrategias_() {
  return objetosDesdeHoja_(V53.SHEETS.STRATEGIES);
}

function obtenerReglas_() {
  return objetosDesdeHoja_(V53.SHEETS.RULES).map(r => ({
    ...r,
    OBLIGATORIA: r.OBLIGATORIA === true || String(r.OBLIGATORIA) === 'TRUE',
    ACTIVA: r.ACTIVA === true || String(r.ACTIVA) === 'TRUE'
  }));
}

function obtenerOperaciones_() {
  return objetosDesdeHoja_(V53.SHEETS.TRADES);
}

function objetosDesdeHoja_(name) {
  const sh = hoja_(name);
  const values = sh.getDataRange().getValues();

  if (values.length < 2) return [];

  const headers = values[0];

  return values.slice(1)
    .filter(row => row.some(v => v !== '' && v !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function obtenerFilas_(sh) {
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
    .getValues();
}

/* =====================================================
   UTILIDADES
===================================================== */

function hoja_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('No existe la hoja ' + name + '. Ejecuta setupV53().');
  return sh;
}

function asegurarSetup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(V53.SHEETS.CONFIG)) setupV53();
}

function guardarConfigValor_(key, value) {
  const sh = hoja_(V53.SHEETS.CONFIG);
  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === key) {
      sh.getRange(i + 1, 2).setValue(value);
      sh.getRange(i + 1, 4).setValue(new Date());
      return;
    }
  }

  sh.appendRow([key, value, '', new Date()]);
}


function incrementarVersionEstrategia_(strategyId) {
  const strategySheet = hoja_(V53.SHEETS.STRATEGIES);
  const rulesSheet = hoja_(V53.SHEETS.RULES);
  const strategies = strategySheet.getDataRange().getValues();

  let strategyRow = -1;
  let oldVersion = 0;

  for (let i = 1; i < strategies.length; i++) {
    if (String(strategies[i][0]) === String(strategyId)) {
      strategyRow = i + 1;
      oldVersion = Number(strategies[i][5]);
      break;
    }
  }

  if (strategyRow < 0) {
    throw new Error('Estrategia no encontrada.');
  }

  const newVersion = oldVersion + 1;
  const rules = rulesSheet.getDataRange().getValues();
  const copies = [];

  // Copia las reglas activas de la versión anterior.
  for (let i = 1; i < rules.length; i++) {
    const rule = rules[i];

    if (
      String(rule[1]) === String(strategyId) &&
      Number(rule[2]) === oldVersion &&
      (rule[7] === true || String(rule[7]) === 'TRUE')
    ) {
      copies.push([
        rule[0],
        rule[1],
        newVersion,
        rule[3],
        rule[4],
        rule[5],
        rule[6],
        true,
        rule[8]
      ]);
    }
  }

  if (copies.length) {
    rulesSheet.getRange(
      rulesSheet.getLastRow() + 1,
      1,
      copies.length,
      copies[0].length
    ).setValues(copies);
  }

  strategySheet.getRange(strategyRow, 6).setValue(newVersion);
  strategySheet.getRange(strategyRow, 7).setValue('BORRADOR');

  return newVersion;
}


function registrarLog_(action, detail) {
  const sh = hoja_(V53.SHEETS.LOG);
  sh.appendRow([new Date(), action, detail]);
}
