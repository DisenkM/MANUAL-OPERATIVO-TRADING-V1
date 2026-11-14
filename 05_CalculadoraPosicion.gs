/* ============================================================================
   DISENK — TRADER JOURNAL | 05_CalculadoraPosicion.gs
   Módulo 5: Calculadora de Tamaño de Posición Binance Spot 1x
   ============================================================================ */

function calcularPosicionBinanceSpot(precioEntrada, precioSL, esCuentaReal) {
  const cfg = DISENK_GLOBAL_CONFIG;
  const riesgoDisponibleUSD = esCuentaReal ? cfg.RIESGO_MAX_REAL_USD : cfg.RIESGO_MAX_DEMO_USD;

  if (precioSL >= precioEntrada) {
    return { estado: "🔴 BLOQUEADO", motivo: "SL incoherente con posición LONG." };
  }

  const distanciaSL = precioEntrada - precioSL;
  const perdidaPorUnidadBTC = distanciaSL + (precioEntrada + precioSL) * cfg.COMISION_BINANCE_SPOT;

  let cantidadBTC = riesgoDisponibleUSD / perdidaPorUnidadBTC;
  cantidadBTC = Math.floor(cantidadBTC * 100000) / 100000; // Redondeo conservador a 5 decimales Binance

  const nocionalUSD = cantidadBTC * precioEntrada;

  if (nocionalUSD < cfg.MIN_NOTIONAL_BINANCE_USD) {
    return {
      estado: "🔴 NO OPERAR — EL TAMAÑO MÍNIMO SUPERA EL RIESGO AUTORIZADO",
      motivo: "El valor nocional ($" + nocionalUSD.toFixed(2) + ") es inferior al mínimo de Binance ($5.00 USDT). Ajustar el mínimo violaría el límite de riesgo de $" + riesgoDisponibleUSD.toFixed(2) + "."
    };
  }

  return {
    etiqueta: "[RESULTADO CALCULADO]",
    estado: "🟢 HABILITADO",
    riesgoUSD: riesgoDisponibleUSD,
    cantidadBTC: cantidadBTC,
    nocionalUSD: nocionalUSD,
    cuentaTipo: esCuentaReal ? "REAL (0.5%)" : "DEMO (1.0%)"
  };
}
