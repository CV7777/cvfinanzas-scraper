#!/usr/bin/env python3
"""
Tipo de Cambio BCCR - Ventanilla
Retorna un JSON con todas las entidades y sus tipos de cambio del día.
https://gee.bccr.fi.cr/IndicadoresEconomicos/Cuadros/frmConsultaTCVentanilla.aspx
"""

import json
import sys
from pathlib import Path

import requests
from bs4 import BeautifulSoup


URL = "https://gee.bccr.fi.cr/IndicadoresEconomicos/Cuadros/frmConsultaTCVentanilla.aspx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}


def parsear_numero(texto: str) -> float | None:
    texto = texto.strip().replace("\xa0", "").replace(" ", "")
    if not texto:
        return None
    try:
        return float(texto.replace(",", "."))
    except ValueError:
        return None


def obtener_tipos_de_cambio() -> list[dict]:
    resp = requests.get(URL, headers=HEADERS, timeout=15)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")

    tabla_datos = None
    for tabla in soup.find_all("table"):
        textos = [td.get_text(strip=True) for td in tabla.find_all("td")]
        if "Entidad Autorizada" in textos and "Compra" in textos:
            tabla_datos = tabla
            break

    if not tabla_datos:
        raise RuntimeError("No se encontró la tabla de tipos de cambio en la página.")

    registros = []
    tipo_entidad_actual = ""

    for fila in tabla_datos.find_all("tr"):
        celdas = fila.find_all("td")
        if len(celdas) < 5:
            continue

        textos = [c.get_text(strip=True) for c in celdas]

        if textos[1] == "Entidad Autorizada":
            continue

        if textos[0]:
            tipo_entidad_actual = textos[0]

        entidad = textos[1]
        compra = parsear_numero(textos[2])
        venta = parsear_numero(textos[3])

        if entidad and compra is not None and venta is not None:
            registros.append({
                "entidad": entidad,
                "compra": compra,
                "venta": venta,
            })

    return registros


def main():
    try:
        registros = obtener_tipos_de_cambio()
    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)

    if not registros:
        print("⚠️ No se encontraron registros.")
        sys.exit(1)

    # Mejor entidad: compra más alta y venta más baja (menor spread)
    # Se prioriza compra alta y venta baja usando score: compra - venta
    # (mayor score = mejor para el cliente)
    mejor = max(registros, key=lambda r: r["compra"] - r["venta"])

    resultado = {
        "entidad": mejor["entidad"],
        "compra": mejor["compra"],
        "venta": mejor["venta"],
    }

    # Guardar en archivo JSON
    ruta_salida = Path(__file__).parent / "datos-json/tipo_cambio_BCCR.json"
    with open(ruta_salida, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    print(f"✅ Mejor tipo de cambio guardado en {ruta_salida.name}")
    print(f"   🏦 {mejor['entidad']}")
    print(f"   💰 Compra: ₡{mejor['compra']} (vendés dólares)")
    print(f"   💵 Venta:  ₡{mejor['venta']} (comprás dólares)")
    print(f"   📊 Spread: ₡{round(mejor['venta'] - mejor['compra'], 2)}")


if __name__ == "__main__":
    main()