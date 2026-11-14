/* ============================================================================
   DISENK — TRADER JOURNAL | 03_AnalisisMultitemporal.gs
   Módulo 3: Validación Multitemporal Estrategia A v3.7 (1D / 1H / M5)
   ============================================================================ */

function validarAnalisisMultitemporal(datosVelas) {
  // datosVelas: { c1D, ema1D_0, ema1D_5, c1H, ema1H_0, ema1H_5, cM5, emaM5, rsiM5, esVelaAlcista }
  let val1D = (datosVelas.c1D > datosVelas.ema1D_0) && (datosVelas.ema1D_0 > datosVelas.ema1D_5);
  let val1H = (datosVelas.c1H > datosVelas.ema1H_0) && (datosVelas.ema1H_0 > datosVelas.ema1H_5);
  
  let valM5 = (datosVelas.cM5 > datosVelas.emaM5) && (datosVelas.rsiM5 > 50) && datosVelas.esVelaAlcista;

  let esValido = val1D && val1H && valM5;

  return {
    etiqueta: "[RESULTADO CALCULADO]",
    filtro1D: val1D ? "🟢 CUMPLE" : "🔴 NO CUMPLE",
    filtro1H: val1H ? "🟢 CUMPLE" : "🔴 NO CUMPLE",
    filtroM5: valM5 ? "🟢 CUMPLE" : "🔴 NO CUMPLE",
    veredicto: esValido ? "🟢 CONTEXTO ALCISTA VÁLIDO" : "🔴 CONTEXTO NO VÁLIDO / BLOQUEADO"
  };
}
