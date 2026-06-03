#!/usr/bin/env python3
"""
Monitor de Portafolio eToro
Ejecuta en tu computador local para obtener datos en tiempo real.
Modo offline disponible con los datos de tu última captura.
"""

import sys
from datetime import datetime

# ──────────────────────────────────────────────────────────────
# TU PORTAFOLIO — actualiza precio_apertura e invertido cuando
# hagas nuevas compras en eToro
# ──────────────────────────────────────────────────────────────
PORTFOLIO = {
    "QQQ":  {"invertido": 1206.46, "precio_apertura": 598.47164, "nombre": "Invesco QQQ (NASDAQ-100)"},
    "VOO":  {"invertido": 1099.15, "precio_apertura": 619.14594, "nombre": "Vanguard S&P 500 ETF"},
    "SCHD": {"invertido":  820.11, "precio_apertura":  29.97711, "nombre": "Schwab US Dividend Equity ETF"},
    "VXUS": {"invertido":  630.97, "precio_apertura":  78.34672, "nombre": "Vanguard Total International ETF"},
    "GLD":  {"invertido":  451.47, "precio_apertura": 412.88090, "nombre": "SPDR Gold"},
    "IEMG": {"invertido":  400.41, "precio_apertura":  72.11247, "nombre": "iShares Core MSCI Emerging Markets"},
    "BIL":  {"invertido":  300.00, "precio_apertura":  91.58000, "nombre": "SPDR Bloomberg 1-3M T-Bill"},
    "IWN":  {"invertido":  150.00, "precio_apertura": 207.42323, "nombre": "iShares Russell 2000 Value ETF"},
    "VNQ":  {"invertido":  100.00, "precio_apertura":  95.13000, "nombre": "Vanguard Real Estate ETF"},
}

# Datos de captura del 2026-06-03 (modo offline)
SNAPSHOT_2026_06_03 = {
    "QQQ":  {"precio": 743.93, "cambio_dia": -0.30, "gp_pct": 24.30},
    "VOO":  {"precio": 693.16, "cambio_dia": -0.73, "gp_pct": 11.95},
    "SCHD": {"precio":  32.34, "cambio_dia": -0.09, "gp_pct":  7.88},
    "VXUS": {"precio":  86.07, "cambio_dia": -1.02, "gp_pct":  9.86},
    "GLD":  {"precio": 407.28, "cambio_dia": -1.13, "gp_pct": -1.36},
    "IEMG": {"precio":  84.82, "cambio_dia": -1.36, "gp_pct": 17.62},
    "BIL":  {"precio":  91.40, "cambio_dia":  0.01, "gp_pct": -0.20},
    "IWN":  {"precio": 212.10, "cambio_dia": -1.35, "gp_pct":  2.25},
    "VNQ":  {"precio":  94.39, "cambio_dia": -0.14, "gp_pct": -0.78},
}

UMBRAL_RSI_SOBREVENTA  = 35
UMBRAL_RSI_SOBRECOMPRA = 70


def calcular_rsi(precios, periodos=14):
    import pandas as pd
    delta = precios.diff()
    ganancias = delta.clip(lower=0)
    perdidas  = -delta.clip(upper=0)
    avg_gan = ganancias.rolling(periodos).mean()
    avg_per = perdidas.rolling(periodos).mean()
    rs = avg_gan / avg_per
    return 100 - (100 / (1 + rs))


def analizar_online():
    import yfinance as yf
    import pandas as pd

    resultados = []
    print("Descargando datos en tiempo real...\n")

    for ticker_str, datos in PORTFOLIO.items():
        try:
            ticker = yf.Ticker(ticker_str)
            hist   = ticker.history(period="6mo")
            info   = ticker.info

            if hist.empty:
                print(f"  {ticker_str}: sin datos")
                continue

            precio_actual = hist["Close"].iloc[-1]
            maximo_52s    = info.get("fiftyTwoWeekHigh", hist["Close"].max())
            ma50          = hist["Close"].rolling(50).mean().iloc[-1]
            rsi           = calcular_rsi(hist["Close"]).iloc[-1]

            precio_apertura = datos["precio_apertura"]
            invertido       = datos["invertido"]
            unidades        = invertido / precio_apertura
            valor_actual    = unidades * precio_actual
            gp              = valor_actual - invertido
            gp_pct          = (gp / invertido) * 100
            dist_max        = ((precio_actual - maximo_52s) / maximo_52s) * 100

            señal = señal_simple(precio_actual, ma50, rsi)

            resultados.append({
                "Ticker": ticker_str,
                "Nombre": datos["nombre"],
                "Precio $": round(precio_actual, 2),
                "Invertido $": round(invertido, 2),
                "Valor $": round(valor_actual, 2),
                "G/P $": round(gp, 2),
                "G/P %": round(gp_pct, 2),
                "RSI": round(rsi, 1),
                "vs MA50 %": round(((precio_actual - ma50) / ma50) * 100, 1),
                "vs Máx52s %": round(dist_max, 1),
                "Señal": señal,
            })
            print(f"  {ticker_str:6} ${precio_actual:.2f}  G/P {gp_pct:+.1f}%  RSI {rsi:.0f}  → {señal}")

        except Exception as e:
            print(f"  {ticker_str}: error — {e}")

    return resultados


def analizar_offline():
    """Usa el snapshot manual de tu última captura de eToro."""
    import pandas as pd

    resultados = []
    snapshot = SNAPSHOT_2026_06_03

    print("Modo OFFLINE — usando datos de captura 2026-06-03\n")

    for ticker_str, datos in PORTFOLIO.items():
        snap = snapshot.get(ticker_str, {})
        precio_actual   = snap.get("precio", datos["precio_apertura"])
        gp_pct          = snap.get("gp_pct", 0.0)
        cambio_dia      = snap.get("cambio_dia", 0.0)

        invertido    = datos["invertido"]
        gp           = invertido * gp_pct / 100
        valor_actual = invertido + gp

        # Sin datos históricos no calculamos RSI real; usamos heurística simple
        rsi_estimado = "N/A"
        señal = señal_por_cambio(gp_pct, cambio_dia)

        resultados.append({
            "Ticker": ticker_str,
            "Nombre": datos["nombre"],
            "Precio $": precio_actual,
            "Invertido $": round(invertido, 2),
            "Valor $": round(valor_actual, 2),
            "G/P $": round(gp, 2),
            "G/P %": round(gp_pct, 2),
            "Cambio día %": cambio_dia,
            "Señal": señal,
        })

    return resultados


def señal_simple(precio, ma50, rsi):
    puntos = 0
    if rsi < UMBRAL_RSI_SOBREVENTA:
        puntos += 2
    elif rsi > UMBRAL_RSI_SOBRECOMPRA:
        puntos -= 2
    if precio > ma50:
        puntos += 1
    else:
        puntos -= 1
    if puntos >= 2:
        return "✅ COMPRAR"
    elif puntos <= -2:
        return "⛔ ESPERAR"
    return "⏳ MANTENER"


def señal_por_cambio(gp_pct, cambio_dia):
    """Señal simple cuando no hay RSI disponible."""
    if cambio_dia <= -1.5 and gp_pct > 0:
        return "✅ COMPRAR"   # caída del día en activo que ya es rentable
    elif gp_pct > 20:
        return "⏳ MANTENER"  # muy ganancioso, sin urgencia de comprar más
    elif gp_pct < -3:
        return "⛔ ESPERAR"
    return "⏳ MANTENER"


def imprimir_tabla(resultados, modo):
    try:
        from tabulate import tabulate
        usar_tabulate = True
    except ImportError:
        usar_tabulate = False

    cols_online  = ["Ticker", "Precio $", "Invertido $", "Valor $", "G/P $", "G/P %", "RSI", "vs MA50 %", "vs Máx52s %", "Señal"]
    cols_offline = ["Ticker", "Precio $", "Invertido $", "Valor $", "G/P $", "G/P %", "Cambio día %", "Señal"]
    cols = cols_online if modo == "online" else cols_offline

    filas = [[r.get(c, "—") for c in cols] for r in resultados]

    if usar_tabulate:
        print(tabulate(filas, headers=cols, tablefmt="rounded_outline", floatfmt=".2f"))
    else:
        print("  " + "  ".join(f"{c:>14}" for c in cols))
        for fila in filas:
            print("  " + "  ".join(f"{str(v):>14}" for v in fila))


def main():
    offline = "--offline" in sys.argv or "-o" in sys.argv

    print(f"\n{'='*70}")
    print(f"  MONITOR PORTAFOLIO eToro  —  {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    modo = "OFFLINE (snapshot)" if offline else "ONLINE (tiempo real)"
    print(f"  Modo: {modo}")
    print(f"{'='*70}\n")

    resultados = analizar_offline() if offline else analizar_online()

    if not resultados:
        print("Sin datos para mostrar.")
        return

    # ── RESUMEN ──
    total_invertido = sum(r["Invertido $"] for r in resultados)
    total_valor     = sum(r["Valor $"] for r in resultados)
    total_gp        = total_valor - total_invertido
    total_gp_pct    = (total_gp / total_invertido) * 100

    print(f"\n{'─'*70}")
    print(f"  RESUMEN TOTAL DEL PORTAFOLIO")
    print(f"{'─'*70}")
    print(f"  Invertido:    ${total_invertido:>10,.2f}")
    print(f"  Valor actual: ${total_valor:>10,.2f}")
    signo = "+" if total_gp >= 0 else ""
    print(f"  G/P total:    {signo}${total_gp:>9,.2f}  ({signo}{total_gp_pct:.2f}%)")

    # ── TABLA ──
    print(f"\n{'─'*70}")
    print(f"  DETALLE POR ACTIVO")
    print(f"{'─'*70}\n")
    modo_str = "offline" if offline else "online"
    imprimir_tabla(resultados, modo_str)

    # ── SEÑALES ──
    comprar  = [r for r in resultados if r["Señal"] == "✅ COMPRAR"]
    esperar  = [r for r in resultados if r["Señal"] == "⛔ ESPERAR"]
    mantener = [r for r in resultados if r["Señal"] == "⏳ MANTENER"]

    print(f"\n{'─'*70}")
    print(f"  SEÑALES DE INVERSIÓN")
    print(f"{'─'*70}")

    if comprar:
        print(f"\n  ✅ OPORTUNIDADES DE COMPRA:")
        for r in comprar:
            gp = r['G/P %']
            signo = "+" if gp >= 0 else ""
            print(f"     • {r['Ticker']:6}  G/P={signo}{gp:.1f}%  —  {r['Nombre']}")

    if mantener:
        print(f"\n  ⏳ MANTENER (sin acción urgente):")
        print(f"     {', '.join(r['Ticker'] for r in mantener)}")

    if esperar:
        print(f"\n  ⛔ ESPERAR ANTES DE INVERTIR MÁS:")
        for r in esperar:
            print(f"     • {r['Ticker']:6}  G/P={r['G/P %']:.1f}%  —  {r['Nombre']}")

    # ── TOP / BOTTOM ──
    ordenado = sorted(resultados, key=lambda x: x["G/P %"], reverse=True)
    print(f"\n{'─'*70}")
    print(f"  RANKING DE RENDIMIENTO")
    print(f"{'─'*70}")
    for i, r in enumerate(ordenado):
        emoji = "🏆" if i < 3 else ("📉" if i >= len(ordenado) - 2 else "  ")
        signo = "+" if r["G/P $"] >= 0 else ""
        print(f"  {emoji} {r['Ticker']:6}  {signo}${r['G/P $']:>7.2f}  ({signo}{r['G/P %']:.2f}%)")

    print(f"\n{'='*70}")
    print(f"  Uso: python3 etoro_monitor.py          → datos en tiempo real")
    print(f"       python3 etoro_monitor.py --offline → usa última captura")
    print(f"\n  AVISO: señales orientativas. No son asesoría financiera.")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()
