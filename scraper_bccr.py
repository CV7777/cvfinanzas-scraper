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
    r = requests.get(url, headers=headers, timeout=60)
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

def excel_serial_to_iso(val):
    """Convierte serial de Excel (número) o string DD/MM/YYYY a YYYY-MM-DD.
    Si la fecha resultante es futura (>hoy), intenta invertir día/mes."""
    from datetime import date
    hoy = date.today().isoformat()

    if val is None or val == "":
        return ""
    try:
        num = float(val)
        if num > 40000:
            from datetime import timedelta
            epoch = date(1899, 12, 30)
            return str(epoch + timedelta(days=int(num)))
    except (ValueError, TypeError):
        pass
    s = str(val).strip()
    # Puede venir como "D/M/YYYY", "DD/MM/YYYY", o "YYYY-MM-DD"
    if "/" in s:
        parts = s.split("/")
        if len(parts) == 3:
            # Puede ser D/M/YYYY (sin ceros) — normalizar
            d, m, y = parts[0].zfill(2), parts[1].zfill(2), parts[2]
            if len(y) == 2:
                y = "20" + y
            result = f"{y}-{m}-{d}"
            # Si la fecha es futura, probablemente está invertida (M/D/YYYY)
            if result > hoy:
                result_inv = f"{y}-{d}-{m}"
                if result_inv <= hoy:
                    return result_inv
            return result
    return s[:10]

def excel_serial_to_time(val):
    """Convierte fracción de día de Excel a string HH:MM. 0.7083 = 17:00"""
    if val is None or val == "":
        return ""
    try:
        frac = float(val)
        if 0 < frac < 1:
            total_minutes = round(frac * 24 * 60)
            h = total_minutes // 60
            m = total_minutes % 60
            return f"{h:02d}:{m:02d}"
        # Si es mayor que 1, es un timestamp completo — extraer la parte decimal
        if frac > 1:
            frac = frac - int(frac)
            total_minutes = round(frac * 24 * 60)
            h = total_minutes // 60
            m = total_minutes % 60
            return f"{h:02d}:{m:02d}"
    except (ValueError, TypeError):
        pass
    return str(val)

def read_all_rows(token, drive_id, item_id, session_id):
    """Lee todas las filas de la tabla TipoCambio"""
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/tables/{TABLE_NAME}/rows"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "workbook-session-id": session_id
    }
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    rows = r.json().get("value", [])
    result = []
    for row in rows:
        vals = row.get("values", [[]])[0]
        if len(vals) >= 5 and vals[0]:
            fecha_iso = excel_serial_to_iso(vals[0])
            sesion_raw = vals[5] if len(vals) > 5 else ""
            ts_raw = vals[6] if len(vals) > 6 else ""

            # Convertir sesión: puede ser "17:00", "13:05" o serial de hora
            sesion_str = str(sesion_raw).strip() if sesion_raw else ""
            if sesion_str and ":" not in sesion_str:
                sesion_str = excel_serial_to_time(sesion_raw)

            # Convertir timestamp: siempre usar fecha_iso (ya corregida) + hora
            if ts_raw and ":" not in str(ts_raw):
                # Es serial de Excel — extraer la hora de la fracción
                hora_str = excel_serial_to_time(ts_raw)
                ts_str = fecha_iso + " " + hora_str
            elif ts_raw:
                # Es string tipo "2/4/2026 17:00" o "2026-02-04 17:00"
                # Siempre usar fecha_iso para la parte de fecha (ya está corregida)
                ts_parts = str(ts_raw).strip().split(" ")
                hora_part = ts_parts[1] if len(ts_parts) > 1 else (sesion_str or "17:00")
                ts_str = fecha_iso + " " + hora_part
            else:
                ts_str = fecha_iso + " " + (sesion_str or "17:00") + ":00"

            result.append({
                "fecha": fecha_iso,
                "promedio_ponderado": vals[1],
                "monto_total": vals[2],
                "minimo": vals[3],
                "maximo": vals[4],
                "sesion": sesion_str,
                "timestamp": ts_str
            })
    return result

def fix_future_date(fecha_str, sesion_str):
    """Si la fecha es futura, intenta invertir mes y día para corregirla."""
    from datetime import date
    hoy = date.today().isoformat()
    if fecha_str > hoy:
        partes = fecha_str.split("-")
        if len(partes) == 3:
            y, m, d = partes
            invertida = f"{y}-{d}-{m}"
            if invertida <= hoy:
                hora = sesion_str if sesion_str else "17:00"
                return invertida, f"{invertida} {hora}"
    return fecha_str, None

def generate_json(all_rows):
    """Genera datos.json con el historial completo"""
    # Corregir fechas futuras (día/mes invertidos por error del scraper viejo)
    for r in all_rows:
        fecha_corregida, ts_corregido = fix_future_date(r["fecha"], r.get("sesion", "17:00"))
        if fecha_corregida != r["fecha"]:
            print(f"  Corrigiendo fecha: {r['fecha']} → {fecha_corregida}")
            r["fecha"] = fecha_corregida
            r["timestamp"] = ts_corregido

    # Ordenar por timestamp para deduplicar correctamente
    sorted_rows = sorted(all_rows, key=lambda x: str(x.get("timestamp", x.get("fecha", ""))))
    # Deduplicar: si hay dos entradas del mismo día, quedarse con la de 17:00
    by_date = {}
    for r in sorted_rows:
        fecha = r["fecha"]
        sesion = r.get("sesion", "")
        if fecha not in by_date or sesion == "17:00":
            by_date[fecha] = r
    # Ordenar el resultado final por timestamp ascendente (más antiguo primero)
    deduped = sorted(by_date.values(), key=lambda x: str(x.get("timestamp", x.get("fecha", ""))))
    output = {
        "actualizado": datetime.now(CR_TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "datos": deduped
    }
    with open("datos.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"  Generado datos.json con {len(deduped)} registros")

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

    # 1. Autenticar con Microsoft (siempre, para poder generar el JSON)
    print("\n[1/4] Autenticando con Microsoft Graph...")
    token = get_token()
    print("  ✓ Token obtenido")

    drive_id, item_id = find_excel_item(token)
    session_id = get_excel_session(token, drive_id, item_id)

    # 2. Extraer datos del BCCR
    print("\n[2/4] Extrayendo datos del BCCR...")
    datos = None
    try:
        datos = scrape_bccr()
    except Exception as e:
        print(f"  ⚠ Error al consultar BCCR: {e}")
        print("  Continuando con historial existente...")

    if datos is None:
        print("  Sin datos nuevos. Generando JSON con historial existente...")
    else:
        print(f"  ✓ Fecha: {datos['fecha']}")
        print(f"  ✓ Promedio Ponderado: {datos['promedio_ponderado']:.2f}")
        print(f"  ✓ Monto Total: {datos['monto_total']:,.2f}")
        print(f"  ✓ Sesion: {datos['sesion']}")

        # 3. Guardar en Excel solo si hay datos nuevos
        print("\n[3/4] Guardando en Excel Online...")
        append_to_excel(token, drive_id, item_id, session_id, datos)
        print("  ✓ Fila agregada exitosamente")

    # 4. Siempre generar datos.json con el historial completo
    print("\n[4/4] Generando datos.json...")
    all_rows = read_all_rows(token, drive_id, item_id, session_id)
    generate_json(all_rows)
    print("  ✓ datos.json generado")

    print("\n✅ Completado exitosamente")
    if datos:
        print(json.dumps(datos, indent=2, ensure_ascii=False))

if __name__ == "__main__":
    main()
