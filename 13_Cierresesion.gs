/* ============================================================================
   DISENK — TRADER JOURNAL | 13_CierreSesion.gs
   Módulo 13: Informe de Cierre de Sesión Automático
   ============================================================================ */

function generarInformeCierreSesion(motivoCierre) {
  const estadoCap = obtenerEstadoCapital();
  
  return {
    etiqueta: "[RESULTADO CALCULADO]",
    trader: "Disenk",
    fecha: Utilities.formatDate(new Date(), "America/Bogota", "yyyy-MM-dd HH:mm:ss"),
    balanceFinalUSD: estadoCap.capitalActualUSD.toFixed(2),
    pnlSesionUSD: estadoCap.pnlHoyUSD.toFixed(2),
    operacionesHoy: estadoCap.opsHoy,
    motivoCierre: motivoCierre || "Límite de operaciones / Fin de tiempo disponible",
    estadoFinal: "🟢 SESIÓN CERRADA CON ÉXITO Y REGISTRADA EN DB"
  };
}
