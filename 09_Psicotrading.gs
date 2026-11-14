/* ============================================================================
   DISENK — TRADER JOURNAL | 09_Psicotrading.gs
   Módulo 9: Registro Emocional Tri-Fase (Antes, Durante, Después)
   ============================================================================ */

function registrarPsicotrading(emociones) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPsico = ss.getSheetByName("DB_PSICOTRADING");

  // Detección de patrones conductuales
  let patron = "Estado Psicológico Estable 🟢";
  if (emociones.fomoPre > 6 || emociones.frustracionPost > 6) {
    patron = "Sesgo Conductual Detectado: Impulsividad / Revancha 🔴";
  }

  shPsico.appendRow([
    emociones.idTrade, emociones.calmaPre, emociones.fomoPre,
    emociones.ansiedadDurante, emociones.euforiaPost, emociones.frustracionPost, patron
  ]);

  return { etiqueta: "[RESULTADO CALCULADO]", patronDetectado: patron };
}
