/* ============================================================================
   DISENK — TRADER JOURNAL | 08_DiarioOperaciones.gs
   Módulo 8: Diario de Operaciones & Clasificación Matricial
   ============================================================================ */

function registrarOperacionDiario(trade) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shOps = ss.getSheetByName("DB_OPERACIONES");
  
  const idTrade = "TRD-" + Date.now();
  const fecha = Utilities.formatDate(new Date(), "America/Bogota", "yyyy-MM-dd HH:mm:ss");

  // Matriz de Clasificación: (Ganadora/Perdedora) x (Disciplina/Incumpliendo)
  let matrizClasificacion = "";
  if (trade.pnlNetoUSD > 0 && trade.cumplioReglas) matrizClasificacion = "Ganadora con Disciplina 🟢";
  else if (trade.pnlNetoUSD > 0 && !trade.cumplioReglas) matrizClasificacion = "Ganadora Incumpliendo Reglas 🟡 (Peligro)";
  else if (trade.pnlNetoUSD <= 0 && trade.cumplioReglas) matrizClasificacion = "Perdedora con Disciplina 🟢 (Riesgo Controlado)";
  else matrizClasificacion = "Perdedora Incumpliendo Reglas 🔴 (Error Grave)";

  shOps.appendRow([
    idTrade, fecha, trade.par || "BTC/USDT", "LONG", trade.entrada, trade.sl, trade.tp,
    trade.cantidadBTC, trade.nocionalUSD, trade.pnlNetoUSD, trade.pnlNetoUSD * 4000,
    trade.rrNeto, matrizClasificacion, "CERRADA"
  ]);

  return { etiqueta: "[DATO DE USUARIO]", idTrade: idTrade, clasificacion: matrizClasificacion };
}
