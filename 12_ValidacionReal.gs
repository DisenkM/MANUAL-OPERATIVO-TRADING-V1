/* ============================================================================
   DISENK — TRADER JOURNAL | 12_ValidacionReal.gs
   Módulo 12: Protocolo de Validación 5/5 para Paso a Dinero Real
   ============================================================================ */

function evaluarPasoADineroReal(metricasDemo) {
  // metricasDemo: { pnlNetoPositivo, expectativaPositiva, drawdownEnLimite, outOfSamplePositivo, disciplinaDemostrada }
  const c1 = metricasDemo.pnlNetoPositivo;
  const c2 = metricasDemo.expectativaPositiva;
  const c3 = metricasDemo.drawdownEnLimite;
  const c4 = metricasDemo.outOfSamplePositivo;
  const c5 = metricasDemo.disciplinaDemostrada;

  const apto = c1 && c2 && c3 && c4 && c5;

  return {
    etiqueta: "[RESULTADO CALCULADO]",
    cumplimiento: (apto ? 5 : 4) + "/5",
    veredicto: apto ? "🟢 APROBADA PARA OPERAR DINERO REAL" : "🔴 NO APROBADA — PERMANECER EN DEMO"
  };
}
