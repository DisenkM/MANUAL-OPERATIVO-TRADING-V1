/* ============================================================================
   DISENK — TRADER JOURNAL | 11_Backtesting.gs
   Módulo 11: Motor de Simulación Histórica de Backtesting (5 Años)
   ============================================================================ */

function registrarResultadoBacktest(simulacion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shBack = ss.getSheetByName("DB_BACKTESTING");

  const idBacktest = "BKT-" + Date.now();
  // Simulación conservadora: Si vela toca SL y TP, contar SL [REGLA APROBADA]
  shBack.appendRow([
    idBacktest, simulacion.mercado, simulacion.tradesCount, simulacion.winRate,
    simulacion.profitFactor, simulacion.expectativa, simulacion.maxDrawdown,
    simulacion.tradesCount >= 100 ? "VALIDADO 🟢" : "INSUFICIENTE 🟡"
  ]);

  return { etiqueta: "[SUPUESTO DE SIMULACIÓN]", idBacktest: idBacktest };
}
