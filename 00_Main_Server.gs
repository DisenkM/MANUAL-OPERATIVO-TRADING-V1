/* ============================================================================
   DISENK — TRADER JOURNAL | 00_Main_Server.gs
   Gobernanza: [REGLA APROBADA] - Servidor Principal y Enrutador de Plantillas
   ============================================================================ */

function doGet() {
  return HtmlService.createTemplateFromFile('V0_Index')
    .evaluate()
    .setTitle('DISENK — TRADER JOURNAL')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Función auxiliar para incluir vistas HTML hijas dentro de V0_Index.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Configuración Global Inquebrantable [REGLA APROBADA]
const DISENK_GLOBAL_CONFIG = {
  TRADER: "Disenk",
  CAPITAL_BASE_USD: 80.00,          // [REGLA APROBADA]
  RIESGO_MAX_REAL_USD: 0.40,        // [REGLA APROBADA] 0.5% max
  RIESGO_MAX_DEMO_USD: 0.80,        // [REGLA APROBADA] 1.0% max
  LIMITE_PERDIDA_DIARIA_USD: 0.80,  // [REGLA APROBADA]
  LIMITE_PERDIDA_MENSUAL_USD: 8.00, // [REGLA APROBADA]
  MAX_OPERACIONES_DIA: 3,           // [REGLA APROBADA]
  MAX_POSICIONES_ABIERTAS: 1,       // [REGLA APROBADA]
  MAX_PERDIDAS_CONSECUTIVAS: 2,     // [REGLA APROBADA]
  PAUSA_OBLIGATORIA_MIN: 15,        // [REGLA APROBADA]
  COMISION_BINANCE_SPOT: 0.0010,    // [REGLA APROBADA] 0.10% Maker/Taker
  MIN_NOTIONAL_BINANCE_USD: 5.00,   // [REGLA APROBADA]
  ESTRATEGIA_ACTIVA: "A v3.7 (LONG Only)" // [REGLA APROBADA]
};

// Inicialización de la Base de Datos Relacional de 7 Tablas
function inicializarBaseDatosDisenk() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tablas = [
    { nombre: "DB_CONFIG", headers: ["Clave", "Valor", "Etiqueta", "Fecha_Actualizacion"] },
    { nombre: "DB_SESIONES", headers: ["ID_Sesion", "Fecha_UTC5", "Mercado", "Noticias", "Checklist_Pre", "Estado"] },
    { nombre: "DB_OPERACIONES", headers: ["ID_Trade", "Fecha_Hora", "Par", "Direccion", "Entrada", "SL", "TP_Neto", "Cantidad_BTC", "Nocional_USD", "PnL_Neto_USD", "PnL_Neto_COP", "RR_Neto", "Disciplina", "Estado"] },
    { nombre: "DB_PSICOTRADING", headers: ["ID_Trade", "Calma_Pre", "FOMO_Pre", "Ansiedad_Durante", "Euforia_Post", "Frustracion_Post", "Patron_Conductual"] },
    { nombre: "DB_ESTADISTICAS", headers: ["Metrica", "Valor", "Objetivo", "Semaforo"] },
    { nombre: "DB_BACKTESTING", headers: ["ID_Backtest", "Mercado", "Trades", "WinRate", "ProfitFactor", "Expectativa", "MaxDD", "Estado_Validacion"] },
    { nombre: "DB_AUDITORIA_CONFIG", headers: ["ID_Cambio", "Fecha", "Parametro", "Valor_Ant", "Valor_Nuevo", "Motivo_Tecnico", "Usuario"] }
  ];

  tablas.forEach(t => {
    let sheet = ss.getSheetByName(t.nombre);
    if (!sheet) {
      sheet = ss.insertSheet(t.nombre);
      sheet.getRange(1, 1, 1, t.headers.length).setValues([t.headers])
           .setFontWeight("bold").setBackground("#161616").setFontColor("#3DDBD9");
    }
  });
  return "🟢 Base de datos DISENK estructurada relacionalmente.";
}
