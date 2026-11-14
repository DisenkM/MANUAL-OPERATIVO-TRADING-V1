/* ============================================================================
   DISENK — TRADER JOURNAL | 06_CalculadoraTP_RR.gs
   Módulo 6: Calculadora Take-Profit 1:2 Neto & Ratio Riesgo:Beneficio
   ============================================================================ */

function calcularTakeProfitNeto(precioEntrada, cantidadBTC, riesgoUSD) {
  const comision = DISENK_GLOBAL_CONFIG.COMISION_BINANCE_SPOT;
  
  // Beneficio neto requerido = 2 veces el riesgo neto (1:2 Neto)
  const beneficioNetoObjetivoUSD = riesgoUSD * 2.0; 
  
  // Fórmula exacta despejando costes de compra y venta Binance Spot
  const costeEntradaUSD = cantidadBTC * precioEntrada * comision;
  const precioTPNeto = precioEntrada + (beneficioNetoObjetivoUSD + costeEntradaUSD) / (cantidadBTC * (1 - comision));

  return {
    etiqueta: "[RESULTADO CALCULADO]",
    etiquetaEstado: "[CONFIGURACIÓN PENDIENTE DE APROBACIÓN FINAL]",
    precioTPNeto: precioTPNeto,
    rrNeto: "1:2.0 Neto de Costes",
    costeEstimadoEntradaUSD: costeEntradaUSD,
    costeEstimadoSalidaUSD: cantidadBTC * precioTPNeto * comision
  };
}
