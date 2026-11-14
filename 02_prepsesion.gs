/* ============================================================================
   DISENK — TRADER JOURNAL | 02_PrepSesion.gs
   Módulo 2: Preparación de Sesión & Checklist Pre-Mercado
   ============================================================================ */

function registrarPreparacionSesion(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DB_SESIONES");
  
  const idSesion = "SES-" + Date.now();
  const fechaUTC5 = Utilities.formatDate(new Date(), "America/Bogota", "yyyy-MM-dd HH:mm:ss");
  
  sheet.appendRow([
    idSesion,
    fechaUTC5,
    payload.mercado || "BTC/USDT Spot",
    payload.noticias || "Sin eventos macro de alto impacto en los próximos 45m",
    payload.checklistValido ? "🟢 CUMPLE 100%" : "🔴 INCOMPLETO",
    payload.checklistValido ? "HABILITADO" : "BLOQUEADO"
  ]);

  return {
    etiqueta: "[DATO DE USUARIO]",
    idSesion: idSesion,
    estado: payload.checklistValido ? "🟢 SESIÓN REGISTRADA Y HABILITADA" : "🔴 SESIÓN BLOQUEADA POR CHECKLIST"
  };
}
