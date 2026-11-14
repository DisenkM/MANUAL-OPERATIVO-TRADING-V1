/* ============================================================================
   DISENK — TRADER JOURNAL | 04_CalculadoraSL.gs
   Módulo 4: Calculadora de Stop-Loss M5 (Mínimo 10 Velas + Margen ATR 14)
   ============================================================================ */

function calcularStopLossM5(minimo10VelasM5, atr14M5, precioEntrada) {
  if (!minimo10VelasM5 || !atr14M5 || !precioEntrada) {
    return { estado: "🔴 INCOMPLETO", motivo: "Faltan datos de entrada [CONFIGURACIÓN PENDIENTE]" };
  }

  const factorATR = 0.10; // [REGLA APROBADA]
  const margen = atr14M5 * factorATR;
  const precioSL = minimo10VelasM5 - margen;
  const distanciaAbsoluta = precioEntrada - precioSL;
  const distanciaPorcentual = (distanciaAbsoluta / precioEntrada) * 100;

  if (precioSL >= precioEntrada) {
    return { estado: "🔴 BLOQUEADO", motivo: "El Stop-Loss debe ser menor al precio de entrada (LONG)." };
  }

  return {
    etiqueta: "[RESULTADO CALCULADO]",
    margenATR: margen,
    precioSL: precioSL,
    distanciaUSD: distanciaAbsoluta,
    distanciaPct: distanciaPorcentual,
    estado: "🟢 SL VÁLIDO"
  };
}
