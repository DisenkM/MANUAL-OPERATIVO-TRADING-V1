/* ============================================================================
   DISENK — TRADER JOURNAL | 01_Estado_Capital.gs
   Módulo 1: Dashboard de Estado & Capital [REGLA APROBADA]
   ============================================================================ */

function obtenerEstadoCapital() {
  const cfg = DISENK_GLOBAL_CONFIG;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shOps = ss.getSheetByName("DB_OPERACIONES");
  
  let pnlHoyUSD = 0;
  let pnlMesUSD = 0;
  let opsHoyCount = 0;
  let rachaPerdidas = 0;
  let posicionesAbiertas = 0;

  if (shOps && shOps.getLastRow() > 1) {
    const data = shOps.getDataRange().getValues();
    const hoyStr = Utilities.formatDate(new Date(), "America/Bogota", "yyyy-MM-dd");
    
    for (let i = 1; i < data.length; i++) {
      const fechaTrade = Utilities.formatDate(new Date(data[i][1]), "America/Bogota", "yyyy-MM-dd");
      const pnlNeto = Number(data[i][9]) || 0;
      const estado = String(data[i][13]);

      if (estado === "ABIERTA") posicionesAbiertas++;
      
      if (fechaTrade === hoyStr && estado === "CERRADA") {
        pnlHoyUSD += pnlNeto;
        opsHoyCount++;
        if (pnlNeto < 0) rachaPerdidas++;
        else rachaPerdidas = 0;
      }
      if (estado === "CERRADA") pnlMesUSD += pnlNeto;
    }
  }

  const capitalActualUSD = cfg.CAPITAL_BASE_USD + pnlMesUSD;
  const riesgoRestanteHoyUSD = Math.max(0, cfg.LIMITE_PERDIDA_DIARIA_USD - Math.abs(Math.min(0, pnlHoyUSD)));
  
  let estadoSemaforo = "🟢 HABILITADO";
  let motivoBloqueo = "Sesión operativa autorizada.";

  if (Math.abs(pnlHoyUSD) >= cfg.LIMITE_PERDIDA_DIARIA_USD) {
    estadoSemaforo = "🔴 BLOQUEADO";
    motivoBloqueo = "Límite de pérdida diaria alcanzado (USD " + cfg.LIMITE_PERDIDA_DIARIA_USD.toFixed(2) + ").";
  } else if (rachaPerdidas >= cfg.MAX_PERDIDAS_CONSECUTIVAS) {
    estadoSemaforo = "🔴 BLOQUEADO";
    motivoBloqueo = "Alcanzadas 2 pérdidas consecutivas en la sesión. Detener operativa.";
  } else if (opsHoyCount >= cfg.MAX_OPERACIONES_DIA) {
    estadoSemaforo = "🔴 BLOQUEADO";
    motivoBloqueo = "Límite máximo de 3 operaciones diarias alcanzado.";
  } else if (pnlMesUSD <= -cfg.LIMITE_PERDIDA_MENSUAL_USD) {
    estadoSemaforo = "🔴 BLOQUEADO";
    motivoBloqueo = "Límite mensual de pérdida alcanzado (USD " + cfg.LIMITE_PERDIDA_MENSUAL_USD.toFixed(2) + "). Operativa suspendida hasta el próximo mes.";
  }

  return {
    etiqueta: "[RESULTADO CALCULADO]",
    trader: cfg.TRADER,
    capitalInicialUSD: cfg.CAPITAL_BASE_USD,
    capitalActualUSD: capitalActualUSD,
    pnlHoyUSD: pnlHoyUSD,
    pnlMesUSD: pnlMesUSD,
    riesgoRestanteUSD: riesgoRestanteHoyUSD,
    opsHoy: opsHoyCount,
    rachaPerdidas: rachaPerdidas,
    posicionesAbiertas: posicionesAbiertas,
    estadoSemaforo: estadoSemaforo,
    motivoBloqueo: motivoBloqueo
  };
}
