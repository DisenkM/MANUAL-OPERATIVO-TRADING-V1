/* ============================================================================
   DISENK — TRADER JOURNAL | 07_ChecklistAprobacion.gs
   Módulo 7: Checklist de Aprobación Institucional de 15 Puntos
   ============================================================================ */

function evaluarChecklist15Puntos(itemsChecklist) {
  // itemsChecklist: arreglo de 15 booleans
  const total = itemsChecklist.length;
  const aprobados = itemsChecklist.filter(Boolean).length;
  
  let veredicto = "🔴 RECHAZADA";
  if (aprobados === 15) veredicto = "🟢 VÁLIDA";
  else if (aprobados >= 13) veredicto = "🟡 PENDIENTE";

  return {
    etiqueta: "[RESULTADO CALCULADO]",
    aprobados: aprobados,
    total: total,
    porcentaje: ((aprobados / total) * 100).toFixed(1) + "%",
    veredicto: veredicto
  };
}
