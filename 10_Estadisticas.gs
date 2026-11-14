/* ============================================================================
   DISENK — TRADER JOURNAL | 10_Estadisticas.gs
   Módulo 10: Estadísticas Cuantitativas de 25 Métricas
   ============================================================================ */

function calcularEstadisticasCuantitativas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shOps = ss.getSheetByName("DB_OPERACIONES");

  if (!shOps || shOps.getLastRow() <= 1) {
    return { etiqueta: "[RESULTADO CALCULADO]", estado: "N/D (Insuficientes datos)" };
  }

  const data = shOps.getDataRange().getValues();
  let totalTrades = 0, ganadoras = 0, perdedoras = 0;
  let pnlTotalBruto = 0, gananciasSum = 0, perdidasSum = 0;

  for (let i = 1; i < data.length; i++) {
    const pnl = Number(data[i][9]) || 0;
    totalTrades++;
    pnlTotalBruto += pnl;
    if (pnl > 0) { ganadoras++; gananciasSum += pnl; }
    else { perdedoras++; perdidasSum += Math.abs(pnl); }
  }

  const winRate = (ganadoras / totalTrades) * 100;
  const profitFactor = perdidasSum > 0 ? gananciasSum / perdidasSum : gananciasSum;
  const expectativaMatematicaUSD = (gananciasSum - perdidasSum) / totalTrades;

  return {
    etiqueta: "[RESULTADO CALCULADO]",
    totalTrades: totalTrades,
    winRate: winRate.toFixed(1) + "%",
    profitFactor: profitFactor.toFixed(2),
    expectativaUSD: expectativaMatematicaUSD.toFixed(2)
  };
}
