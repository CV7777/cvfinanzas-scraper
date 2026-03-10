"""
CV Finanzas - Scraper MONEX (BCCR)
Extrae: Promedio Simple y Monto Total del día
Guarda en: Excel Online via Microsoft Graph API
Ejecutar: 2 veces al día (13:05 y 17:00 hora Costa Rica)
"""

import os
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import pytz
import json
import sys

# ── CONFIGURACIÓN ──────────────────────────────────────────
SHAREPOINT_SITE   = "cvfinanzas-my.sharepoint.com"
SHAREPOINT_USER   = "carlos@cvfinanzas.com"
EXCEL_FILE_NAME   = "CV Finanzas - Tipo de Cambio.xlsx"
TABLE_NAME        = "TipoCambio"

# Estos valores los obtenés en Azure (instrucciones abajo)
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
# ───────────────────────────────────────────────────────────

CR_TZ = pytz.timezone("America/Costa_Rica")

def get_token():
    """Obtiene token de acceso a Microsoft Graph API"""
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()["access_token"]

def get_excel_session(token, drive_id, item_id):
    """Abre sesión persistente en el Excel"""
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/createSession"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, json={"persistChanges": True})
    r.raise_for_status()
    return r.json()["id"]

def find_excel_item(token):
    """Busca el archivo Excel en OneDrive del usuario"""
    headers = {"Authorization": f"Bearer {token}"}

    # Intentar directamente con el email del usuario
    url = f"https://graph.microsoft.com/v1.0/users/{SHAREPOINT_USER}/drive/root/search(q='{EXCEL_FILE_NAME}')"
    r = requests.get(url, headers=headers)

    if r.status_code != 200:
        print(f"  ⚠ Intento 1 falló ({r.status_code}), probando alternativa...")
        # Fallback: buscar en el site de SharePoint directamente
        url2 = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE}/drive/root/search(q='{EXCEL_FILE_NAME}')"
        r = requests.get(url2, headers=headers)

    if r.status_code != 200:
        print(f"  ⚠ Intento 2 falló ({r.status_code}), probando alternativa...")
        # Fallback 2: listar todos los drives del site
        url3 = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE}/drives"
        r3 = requests.get(url3, headers=headers)
        if r3.status_code == 200:
            drives = r3.json().get("value", [])
            print(f"  Drives encontrados: {[d['name'] for d in drives]}")
            for drive in drives:
                url4 = f"https://graph.microsoft.com/v1.0/drives/{drive['id']}/root/search(q='{EXCEL_FILE_NAME}')"
                r4 = requests.get(url4, headers=headers)
                if r4.status_code == 200:
                    items = r4.json().get("value", [])
                    if items:
                        item = items[0]
                        return item["parentReference"]["driveId"], item["id"]
        raise Exception(f"No se pudo encontrar el archivo en ningún drive.")

    r.raise_for_status()
    items = r.json().get("value", [])
    if not items:
        raise Exception(f"No se encontró el archivo: {EXCEL_FILE_NAME}")
    item = items[0]
    return item["parentReference"]["driveId"], item["id"]

def scrape_bccr():
    """Extrae datos de MONEX del BCCR para el día de hoy"""
    now_cr = datetime.now(CR_TZ)
    fecha_str = now_cr.strftime("%Y/%m/%d")
    fecha_label = now_cr.strftime("%d/%m/%Y")
    sesion = "13:05" if now_cr.hour < 15 else "17:00"

    url = (
        f"https://gee.bccr.fi.cr/indicadoreseconomicos/Cuadros/frmVerCatCuadro.aspx"
        f"?CodCuadro=770&Idioma=1"
        f"&FecInicial={fecha_str}&FecFinal={fecha_str}&Filtro=0"
    )

    headers = {"User-Agent": "Mozilla/5.0 (compatible; CVFinanzas/1.0)"}
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()

    soup = BeautifulSoup(r.text, "html.parser")
    rows = soup.find_all("tr")

    data = {
        "promedio_ponderado": None,
        "monto_total": None,
        "minimo": None,
        "maximo": None,
    }

    # Track which section we are in (tipo de cambio vs monto negociado)
    in_tipo_cambio = False
    in_monto = False

    for row in rows:
        cells = [td.get_text(strip=True).replace("\xa0", " ").strip() for td in row.find_all("td")]
        text = " ".join(cells).lower()

        # Detect section headers
        if "tipo de cambio negociado" in text:
            in_tipo_cambio = True
            in_monto = False
            continue
        if "monto negociado" in text:
            in_tipo_cambio = False
            in_monto = True
            continue
        if "mejores ofertas" in text:
            in_tipo_cambio = False
            in_monto = False
            continue

        nums = [_parse_num(cell) for cell in cells if _is_number(cell)]

        if in_tipo_cambio:
            if "promedio ponderado" in text and "sesión anterior" not in text and "anterior" not in text:
                if nums:
                    data["promedio_ponderado"] = nums[-1]
            elif "mínimo" in text or "minimo" in text:
                if nums:
                    data["minimo"] = nums[-1]
            elif "máximo" in text or "maximo" in text:
                if nums:
                    data["maximo"] = nums[-1]

        if in_monto:
            if "monto total" in text or ("total" in text and "calces" not in text and "calce" not in text):
                if nums:
                    data["monto_total"] = nums[-1]

    if not data["promedio_ponderado"]:
        print("  Sin datos disponibles aun para hoy (el BCCR publica a las 13:05 y 17:00).")
        return None

    return {
        "fecha": fecha_label,
        "promedio_ponderado": data["promedio_ponderado"],
        "monto_total": data["monto_total"] or 0,
        "minimo": data["minimo"] or 0,
        "maximo": data["maximo"] or 0,
        "sesion": sesion,
        "timestamp": now_cr.strftime("%Y-%m-%d %H:%M:%S")
    }

def _is_number(s):
    try:
        float(s.replace(".", "").replace(",", "").replace("-", ""))
        return len(s) > 0
    except:
        return False

def _parse_num(s):
    try:
        s = s.strip()
        # BCCR usa punto como separador de miles y coma como decimal
        # Ej: 478.260,00 o 63.376.000,00
        if "," in s and "." in s:
            # tiene ambos: puntos = miles, coma = decimal
            cleaned = s.replace(".", "").replace(",", ".")
        elif "," in s:
            # solo coma: puede ser decimal
            cleaned = s.replace(",", ".")
        else:
            cleaned = s
        return float(cleaned)
    except:
        return None

def append_to_excel(token, drive_id, item_id, session_id, row_data):
    """Agrega una fila nueva a la tabla TipoCambio"""
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/tables/{TABLE_NAME}/rows/add"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "workbook-session-id": session_id
    }
    values = [[
        row_data["fecha"],
        row_data["promedio_ponderado"],
        row_data["monto_total"],
        row_data["minimo"],
        row_data["maximo"],
        row_data["sesion"],
        row_data["timestamp"]
    ]]
    r = requests.post(url, headers=headers, json={"values": values})
    r.raise_for_status()
    return r.json()

def main():
    print("=" * 50)
    print("CV Finanzas - Scraper MONEX")
    print(f"Hora CR: {datetime.now(CR_TZ).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    # 1. Extraer datos del BCCR
    print("\n[1/3] Extrayendo datos del BCCR...")
    datos = scrape_bccr()
    if datos is None:
        print("\n⚠ Sin datos para guardar. El script se ejecutará de nuevo en el próximo horario.")
        return
    print(f"  ✓ Fecha: {datos['fecha']}")
    print(f"  ✓ Promedio Ponderado: ₡{datos['promedio_ponderado']:.2f}")
    print(f"  ✓ Monto Total: ${datos['monto_total']:,.2f}")
    print(f"  ✓ Sesión: {datos['sesion']}")

    # 2. Autenticar con Microsoft
    print("\n[2/3] Autenticando con Microsoft Graph...")
    token = get_token()
    print("  ✓ Token obtenido")

    # 3. Guardar en Excel
    print("\n[3/3] Guardando en Excel Online...")
    drive_id, item_id = find_excel_item(token)
    session_id = get_excel_session(token, drive_id, item_id)
    append_to_excel(token, drive_id, item_id, session_id, datos)
    print("  ✓ Fila agregada exitosamente")

    print("\n✅ Completado exitosamente")
    print(json.dumps(datos, indent=2, ensure_ascii=False))

if __name__ == "__main__":
    main()
