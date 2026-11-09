/**
 * ============================================================================
 * MANUAL OPERATIVO DE TRADING - SISTEMA COMPLETO EN GOOGLE APPS SCRIPT
 * Par: BTC/USDT | Binance Spot (Longs) | Capital: $80 USD | Riesgo: $0.40 USD
 * ============================================================================
 * 
 * INSTRUCCIONES DE INSTALACIÓN:
 * 1. Abre una hoja de cálculo limpia en Google Sheets.
 * 2. Ve a Extensiones -> Apps Script.
 * 3. Reemplaza todo el contenido del archivo Code.gs con este código.
 * 4. Haz clic en el botón de Guardar (💾) y recarga la hoja de Google Sheets.
 * 5. Utiliza el nuevo menú "🚀 Trading Manual BTC" -> "⚙️ 1. Crear / Resetear Estructura de Hojas".
 */

// 1. CREACIÓN DEL MENÚ PERSONALIZADO EN GOOGLE SHEETS
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Trading Manual BTC')
    .addItem('⚙️ 1. Crear / Resetear Estructura de Hojas', 'crearEstructuraCompleta')
    .addSeparator()
    .addItem('➕ 2. Registrar Trade Actual en el Diario', 'registrarTradeEnDiario')
    .addToUi();
}

// 2. CONSTRUCTOR PRINCIPAL DE LA ESTRUCTURA Y BASE DE DATOS
function crearEstructuraCompleta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const nombreCalc = "Calculadora Spot";
  const nombreCheck = "Checklist Operativo";
  const nombreDiario = "Diario de Trading";
  const nombreDash = "Dashboard & Estadísticas";
  
  const shCalc = obtenerOCrearHoja(ss, nombreCalc);
  const shCheck = obtenerOCrearHoja(ss, nombreCheck);
  const shDiario = obtenerOCrearHoja(ss, nombreDiario);
  const shDash = obtenerOCrearHoja(ss, nombreDash);
  
  configurarCalculadora(shCalc);
  configurarChecklist(shCheck);
  configurarDiario(shDiario);
  configurarDashboard(shDash);
  
  SpreadsheetApp.getUi().alert("✅ ¡Estructura creada exitosamente con la lógica de tu Manual Operativo!");
}

// ----------------------------------------------------------------------------
// A. CONFIGURACIÓN DE LA CALCULADORA DINÁMICA (PESTAÑA 1)
// ----------------------------------------------------------------------------
function configurarCalculadora(sheet) {
  sheet.clear();
  sheet.setHiddenGridlines(false);
  
  // Título principal
  sheet.getRange("B2:I2").merge()
       .setValue("CALCULADORA DE POSICIÓN - BINANCE SPOT (BTC/USDT)")
       .setFontWeight("bold").setFontSize(14)
       .setBackground("#1F4E79").setFontColor("#FFFFFF")
       .setHorizontalAlignment("center").setVerticalAlignment("middle");
       
  // Sección 1: Datos de Entrada (PASO 1)
  const inputs = [
    ["PARÁMETROS DE CUENTA Y TRADE", ""],
    ["Capital Inicial ($ USD)", 80.00],
    ["Riesgo Máximo por Trade ($ USD)", 0.40],
    ["Tasa de Cambio (COP / USD)", 4000],
    ["Comisión Binance Spot (0.10%)", 0.001],
    ["Precio de Entrada BTC ($ USD)", 65000.00],
    ["Precio de Stop Loss ($ USD)", 64500.00]
  ];
  
  sheet.getRange("B4:C10").setValues(inputs);
  sheet.getRange("B4:C4").merge().setFontWeight("bold").setBackground("#2F5496").setFontColor("#FFFFFF");
  sheet.getRange("B5:B10").setFontWeight("bold");
  sheet.getRange("C5:C10").setBackground("#DCE6F1"); // Color azul editable
  
  // Formatos numéricos de entrada
  sheet.getRange("C5:C7").setNumberFormat("$#,##0.00");
  sheet.getRange("C8").setNumberFormat("$#,##0");
  sheet.getRange("C9").setNumberFormat("0.000%");
  sheet.getRange("C10").setNumberFormat("$#,##0.00");
  
  // Sección 2: Resultados y Tamaño de Posición (PASO 2)
  sheet.getRange("E4:F4").merge().setValue("RESULTADOS Y TAMAÑO DE POSICIÓN")
       .setFontWeight("bold").setBackground("#2F5496").setFontColor("#FFFFFF");
       
  const formulasCalc = [
    ["Distancia al SL ($ USD)", "=IF(C10>0, C10-C11, 0)"],
    ["Variación de Stop Loss (% SL)", "=IF(C10>0, F5/C10, 0)"],
    ["Variación de Take Profit (% TP - 1:2)", "=F6*2"],
    ["Precio Objetivo Take Profit ($ USD)", "=IF(C10>0, C10+(F5*2), 0)"],
    ["Tamaño Posición (Contratos BTC)", "=IF(F5>0, C6/(F5+(C10*C9*2)), 0)"],
    ["Valor Nocional de la Orden ($ USD)", "=F9*C10"],
    ["Validación Mínimo Binance ($5 USDT)", '=IF(F10>=5, "✅ CUMPLE (>$5 USDT)", "❌ NO CUMPLE (<$5 USDT)")']
  ];
  
  sheet.getRange("E5:F11").setValues(formulasCalc);
  sheet.getRange("E5:E11").setFontWeight("bold");
  sheet.getRange("F5:F11").setBackground("#E2EFDA"); // Color verde resultados
  
  // Formatos numéricos de resultados
  sheet.getRange("F5").setNumberFormat("$#,##0.00");
  sheet.getRange("F6:F7").setNumberFormat("0.00%");
  sheet.getRange("F8").setNumberFormat("$#,##0.00");
  sheet.getRange("F9").setNumberFormat("0.00000000");
  sheet.getRange("F10").setNumberFormat("$#,##0.00");
  
  // Sección 3: Desglose de Comisiones y PnL Neto (PASO 3)
  sheet.getRange("H4:I4").merge().setValue("DESGLOSE DE COMISIONES Y PnL NETO")
       .setFontWeight("bold").setBackground("#2F5496").setFontColor("#FFFFFF");
       
  const formulasPnL = [
    ["Comisión Entrada ($ USD)", "=F10*C9"],
    ["Comisión Salida en TP ($ USD)", "=(F8*F9)*C9"],
    ["Comisión Salida en SL ($ USD)", "=(C11*F9)*C9"],
    ["Ganancia Neta Real en TP ($ USD)", "=((F8-C10)*F9)-(I5+I6)"],
    ["Ganancia Neta Real en TP ($ COP)", "=I8*C8"],
    ["Pérdida Neta Real en SL ($ USD)", "=((C10-C11)*F9)+(I5+I7)"],
    ["Pérdida Neta Real en SL ($ COP)", "=I10*C8"]
  ];
  
  sheet.getRange("H5:I11").setValues(formulasPnL);
  sheet.getRange("H5:H11").setFontWeight("bold");
  sheet.getRange("I5:I7").setBackground("#F2F2F2");
  sheet.getRange("I8:I9").setBackground("#C6EFCE").setFontColor("#006100").setFontWeight("bold"); // Verde TP
  sheet.getRange("I10:I11").setBackground("#FFC7CE").setFontColor("#9C0006").setFontWeight("bold"); // Rojo SL
  
  sheet.getRange("I5:I7").setNumberFormat("$#,##0.0000");
  sheet.getRange("I8").setNumberFormat("$#,##0.00");
  sheet.getRange("I9").setNumberFormat("$#,##0");
  sheet.getRange("I10").setNumberFormat("$#,##0.00");
  sheet.getRange("I11").setNumberFormat("$#,##0");

  sheet.autoResizeColumns(2, 8);
}

// ----------------------------------------------------------------------------
// B. CONFIGURACIÓN DEL CHECKLIST OPERATIVO (PESTAÑA 2)
// ----------------------------------------------------------------------------
function configurarChecklist(sheet) {
  sheet.clear();
  
  sheet.getRange("B2:E2").merge().setValue("CHECKLIST DE CONTROL OPERATIVO - PASO A PASO")
       .setFontWeight("bold").setFontSize(14)
       .setBackground("#1F4E79").setFontColor("#FFFFFF").setHorizontalAlignment("center");
       
  const headers = ["Fase", "Requisito / Regla del Manual", "Estado", "Verificación Técnica"];
  sheet.getRange("B4:E4").setValues([headers]).setFontWeight("bold").setBackground("#2F5496").setFontColor("#FFFFFF");
  
  const checklistData = [
    ["1. Pre-Sesión", "Disponibilidad de bloque de 2 horas continuas (UTC-5 Colombia)", "PENDIENTE", "Confirmar tiempo sin interrupciones"],
    ["1. Pre-Sesión", "Operativa exclusiva en Binance Spot BTC/USDT (Longs sin apalancamiento)", "PENDIENTE", "Verificar plataforma y par"],
    ["1. Pre-Sesión", "Capital inicial disponible ($80 USD) y pérdida del mes < $8 USD", "PENDIENTE", "Revisar saldo de la cuenta"],
    ["1. Pre-Sesión", "Sin posiciones abiertas y pérdidas del día < 2 consecutivas", "PENDIENTE", "Máximo 3 operaciones por día"],
    ["2. Filtro H1", "Precio de BTC por encima de la EMA 50 en H1", "PENDIENTE", "Tendencia macro alcista"],
    ["2. Filtro H1", "EMA 50 con pendiente claramente ascendente en H1", "PENDIENTE", "Inclinación positiva"],
    ["2. Filtro H1", "Indicador RSI (14) por encima de 50 en H1", "PENDIENTE", "Impulso comprador activo"],
    ["3. Setup M15", "Retroceso del precio a la EMA 50 y cierre de vela alcista por encima", "PENDIENTE", "Respeto del soporte dinámico"],
    ["3. Setup M15", "RSI (14) en M15 se mantiene superior a 50 al cierre de vela", "PENDIENTE", "Confirmación de fuerza"],
    ["3. Setup M15", "Vela actual cierra por encima del máximo de la vela M15 anterior", "PENDIENTE", "Gatillo de entrada validado"],
    ["4. Riesgo", "Stop Loss ubicado bajo el mínimo del retroceso en M15 (Riesgo <= $0.40 USD)", "PENDIENTE", "Gestión estricta de capital"],
    ["4. Riesgo", "Take Profit configurado para ofrecer un ratio mínimo de 1:2 R:R (+$0.80 USD)", "PENDIENTE", "Esperanza matemática positiva"],
    ["5. Ejecución", "Orden de compra a mercado ejecutada al inicio de la siguiente vela M15", "PENDIENTE", "Sin vacilar"],
    ["5. Ejecución", "Colocación inmediata de órdenes OCO / SL y TP fijas en Binance", "PENDIENTE", "No mover SL/TP en caliente"],
    ["6. Cierre", "Registro completo del trade en el Diario de Trading al finalizar la sesión", "PENDIENTE", "Evaluación disciplinaria"]
  ];
  
  sheet.getRange("B5:E19").setValues(checklistData);
  
  // Validación de datos desplegable para el estado
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(["✅ CUMPLIDO", "❌ NO CUMPLE", "PENDIENTE"]).build();
  sheet.getRange("C5:C19").setDataValidation(rule).setBackground("#FFF2CC");
  
  sheet.autoResizeColumns(2, 4);
}

// ----------------------------------------------------------------------------
// C. CONFIGURACIÓN DEL DIARIO DE TRADING / BASE DE DATOS (PESTAÑA 3)
// ----------------------------------------------------------------------------
function configurarDiario(sheet) {
  sheet.clear();
  
  sheet.getRange("A2:N2").merge().setValue("BITÁCORA Y BASE DE DATOS DE OPERACIONES")
       .setFontWeight("bold").setFontSize(14)
       .setBackground("#1F4E79").setFontColor("#FFFFFF").setHorizontalAlignment("center");
       
  const headers = [
    "ID Trade", "Fecha / Hora", "Precio Entrada ($)", "Stop Loss ($)", "Take Profit ($)",
    "Precio Salida Real ($)", "Contratos BTC", "Valor Nocional ($)", "Comisiones ($)",
    "PnL Neto ($ USD)", "PnL Neto ($ COP)", "Resultado R", "Plan Cumplido (100%)", "Emociones / Notas"
  ];
  
  sheet.getRange("A4:N4").setValues([headers]).setFontWeight("bold").setBackground("#2F5496").setFontColor("#FFFFFF");
  
  // Fila de ejemplo inicial
  const ejemplo = [
    1, new Date(), 65000.00, 64500.00, 66000.00, 66000.00, 0.00079365, 51.59, 0.1040,
    0.68, 2720, 1.70, "SÍ", "Trade ejecutado según checklist M15 sin vacilación."
  ];
  sheet.getRange("A5:N5").setValues([ejemplo]);
  
  // Formatos numéricos
  sheet.getRange("A5:A100").setHorizontalAlignment("center");
  sheet.getRange("B5:B100").setNumberFormat("YYYY-MM-DD HH:mm");
  sheet.getRange("C5:F100").setNumberFormat("$#,##0.00");
  sheet.getRange("G5:G100").setNumberFormat("0.00000000");
  sheet.getRange("H5:J100").setNumberFormat("$#,##0.00");
  sheet.getRange("K5:K100").setNumberFormat("$#,##0");
  sheet.getRange("L5:L100").setNumberFormat("0.00\"R\"");
  sheet.getRange("M5:M100").setHorizontalAlignment("center");
  
  sheet.autoResizeColumns(1, 14);
}

// ----------------------------------------------------------------------------
// D. CONFIGURACIÓN DEL DASHBOARD Y ESTADÍSTICAS (PESTAÑA 4)
// ----------------------------------------------------------------------------
function configurarDashboard(sheet) {
  sheet.clear();
  
  sheet.getRange("B2:C2").merge().setValue("DASHBOARD DE RENDIMIENTO Y RENDIMIENTO ACUMULADO")
       .setFontWeight("bold").setFontSize(14)
       .setBackground("#1F4E79").setFontColor("#FFFFFF").setHorizontalAlignment("center");
       
  const kpis = [
    ["MÉTRICA CLAVE", "VALOR AUTOMÁTICO"],
    ["Total Operaciones Registradas", "=COUNTA('Diario de Trading'!A5:A100)"],
    ["Operaciones Ganadoras (TP)", '=COUNTIF(\'Diario de Trading\'!J5:J100, ">0")'],
    ["Operaciones Perdedoras (SL)", '=COUNTIF(\'Diario de Trading\'!J5:J100, "<0")'],
    ["Tasa de Acierto (Win Rate %)", '=IF(C5>0, C6/C5, 0)'],
    ["PnL Total Acumulado ($ USD)", "=SUM('Diario de Trading'!J5:J100)"],
    ["PnL Total Acumulado ($ COP)", "=SUM('Diario de Trading'!K5:K100)"],
    ["Esperanza Matemática ($ EV / Trade)", '=IF(C5>0, C9/C5, 0)'],
    ["% Cumplimiento Estricto del Plan", '=IF(C5>0, COUNTIF(\'Diario de Trading\'!M5:M100, "SÍ")/C5, 0)']
  ];
  
  sheet.getRange("B4:C12").setValues(kpis);
  sheet.getRange("B4:C4").setFontWeight("bold").setBackground("#2F5496").setFontColor("#FFFFFF");
  sheet.getRange("B5:B12").setFontWeight("bold");
  sheet.getRange("C5:C12").setBackground("#E2EFDA").setFontWeight("bold");
  
  sheet.getRange("C8").setNumberFormat("0.00%");
  sheet.getRange("C9").setNumberFormat("$#,##0.00");
  sheet.getRange("C10").setNumberFormat("$#,##0");
  sheet.getRange("C11").setNumberFormat("$#,##0.00");
  sheet.getRange("C12").setNumberFormat("0.00%");
  
  sheet.autoResizeColumns(2, 2);
}

// ----------------------------------------------------------------------------
// 3. FUNCIÓN INTERACTIVA: REGISTRAR TRADE DESDE LA CALCULADORA AL DIARIO
// ----------------------------------------------------------------------------
function registrarTradeEnDiario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCalc = ss.getSheetByName("Calculadora Spot");
  const shDiario = ss.getSheetByName("Diario de Trading");
  
  if (!shCalc || !shDiario) {
    SpreadsheetApp.getUi().alert("❌ Primero debes ejecutar la opción 1 para crear las hojas.");
    return;
  }
  
  const entrada = shCalc.getRange("C10").getValue();
  const sl = shCalc.getRange("C11").getValue();
  const tp = shCalc.getRange("F8").getValue();
  const contratos = shCalc.getRange("F9").getValue();
  const nocional = shCalc.getRange("F10").getValue();
  const comisionesUSD = shCalc.getRange("I5").getValue() + shCalc.getRange("I6").getValue();
  const pnlUSD = shCalc.getRange("I8").getValue();
  const pnlCOP = shCalc.getRange("I9").getValue();
  
  if (entrada <= 0 || sl <= 0) {
    SpreadsheetApp.getUi().alert("❌ Ingresa un precio de entrada y Stop Loss válidos en la Calculadora.");
    return;
  }
  
  const ultFila = Math.max(shDiario.getLastRow(), 4);
  const nuevoId = ultFila === 4 ? 1 : Number(shDiario.getRange(ultFila, 1).getValue()) + 1;
  
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert("¿El trade alcanzó el Take Profit (TP)?", ui.ButtonSet.YES_NO_CANCEL);
  
  let precioSalida = tp;
  let pnlFinalUSD = pnlUSD;
  let pnlFinalCOP = pnlCOP;
  let rMultiplo = (tp - entrada) / (entrada - sl);
  
  if (resp === ui.Button.NO) {
    precioSalida = sl;
    pnlFinalUSD = -shCalc.getRange("I10").getValue();
    pnlFinalCOP = -shCalc.getRange("I11").getValue();
    rMultiplo = -1.0;
  } else if (resp === ui.Button.CANCEL) {
    return;
  }
  
  const nuevaFila = [
    nuevoId, new Date(), entrada, sl, tp, precioSalida, contratos, nocional,
    comisionesUSD, pnlFinalUSD, pnlFinalCOP, rMultiplo, "SÍ", "Registrado automáticamente desde Calculadora"
  ];
  
  shDiario.appendRow(nuevaFila);
  ui.alert("✅ Trade #" + nuevoId + " registrado exitosamente en la bitácora.");
}

// ----------------------------------------------------------------------------
// HELPER: OBTENER O CREAR HOJA
// ----------------------------------------------------------------------------
function obtenerOCrearHoja(ss, nombre) {
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
  }
  return sheet;
}
