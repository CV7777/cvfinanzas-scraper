"""
CV Finanzas - Importador MONEX (BCCR)
Consume la API publica SDDE del BCCR y guarda el resultado en Excel Online.
Ejecutar: 2 veces al dia (13:05 y 17:00 hora Costa Rica).
"""

import os
import requests
from datetime import datetime
import pytz
import json

# ── CONFIGURACIÓN ──────────────────────────────────────────
SHAREPOINT_SITE   = "cvfinanzas-my.sharepoint.com"
SHAREPOINT_USER   = "carlos@cvfinanzas.com"
EXCEL_FILE_NAME   = "CV Finanzas - Tipo de Cambio.xlsx"
TABLE_NAME        = "TipoCambio"

# Estos valores los obtenés en Azure (instrucciones abajo)
TENANT_ID     = os.environ.get("AZURE_TENANT_ID")
CLIENT_ID     = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")
# ───────────────────────────────────────────────────────────

CR_TZ = pytz.timezone("America/Costa_Rica")
BCCR_API_URL = (
    "https://apim.bccr.fi.cr/SDDE/api/"
    "Bccr.GE.SDDE.Publico.Indicadores.API/cuadro/219/series"
)
BCCR_CODIGO_CUADRO = "219"
BCCR_API_BEARER_TOKEN = os.environ.get("BCCR_API_BEARER_TOKEN")
BCCR_INDICADORES = {
    "3436": "minimo",
    "3437": "maximo",
    "3439": "promedio_ponderado",
    "3446": "monto_total",
}

# Feriados Costa Rica 2026 (MONEX no opera en feriados)
FERIADOS_2026 = [
    "2026-01-01",  # Año Nuevo
    "2026-04-02",  # Jueves Santo
    "2026-04-03",  # Viernes Santo
    "2026-04-11",  # Día de Juan Santamaría
    "2026-05-01",  # Día Internacional del Trabajo
    "2026-07-25",  # Anexión del Partido de Nicoya
    "2026-08-02",  # Día de la Virgen de los Ángeles
    "2026-08-15",  # Día de la Madre
    "2026-08-31",  # Día de la Persona Negra y Cultura Afrocostarricense
    "2026-09-15",  # Día de la Independencia
    "2026-12-01",  # Día de la Abolición del Ejército
    "2026-12-25",  # Navidad
]

def is_feriado(fecha_str):
    """Valida si la fecha es un feriado (MONEX no opera)"""
    return fecha_str in FERIADOS_2026

def get_token():
    """Obtiene token de acceso a Microsoft Graph API"""
    if not all((TENANT_ID, CLIENT_ID, CLIENT_SECRET)):
        raise RuntimeError(
            "Faltan AZURE_TENANT_ID, AZURE_CLIENT_ID o AZURE_CLIENT_SECRET."
        )
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

def parse_bccr_response(payload, fecha_label, sesion, timestamp):
    """Convierte la respuesta del cuadro 219 al formato interno de MONEX."""
    if not isinstance(payload, dict) or payload.get("estado") is not True:
        mensaje = (
            payload.get("mensaje", "respuesta invalida")
            if isinstance(payload, dict)
            else "respuesta invalida"
        )
        raise ValueError(f"La API del BCCR rechazo la consulta: {mensaje}")

    valores = {}
    bloques = payload.get("datos") or []
    for bloque in bloques:
        for indicador in bloque.get("indicadores") or []:
            campo = BCCR_INDICADORES.get(str(indicador.get("codigoIndicador")))
            if not campo:
                continue
            for punto in indicador.get("series") or []:
                if str(punto.get("fecha", ""))[:10] == fecha_label:
                    valor = punto.get("valorDatoPorPeriodo")
                    if not isinstance(valor, (int, float)) or isinstance(valor, bool):
                        raise ValueError(
                            f"El indicador {campo} no contiene un valor numerico."
                        )
                    valores[campo] = valor
                    break

    faltantes = [campo for campo in BCCR_INDICADORES.values() if campo not in valores]
    if faltantes:
        print(f"  Sin datos completos para {fecha_label}: faltan {', '.join(faltantes)}.")
        return None

    if valores["promedio_ponderado"] <= 0 or valores["monto_total"] <= 0:
        print(f"  Sin negociacion publicada para {fecha_label}.")
        return None
    if not (
        0
        < valores["minimo"]
        <= valores["promedio_ponderado"]
        <= valores["maximo"]
    ):
        raise ValueError(
            "El minimo, promedio ponderado y maximo del BCCR son inconsistentes."
        )

    return {
        "fecha": fecha_label,
        "promedio_ponderado": valores["promedio_ponderado"],
        "monto_total": valores["monto_total"],
        "minimo": valores["minimo"],
        "maximo": valores["maximo"],
        "sesion": sesion,
        "timestamp": timestamp,
    }


def scrape_bccr(now_cr=None):
    """Consulta la API SDDE del BCCR para el dia actual de Costa Rica."""
    now_cr = now_cr or datetime.now(CR_TZ)
    fecha_api = now_cr.strftime("%Y/%m/%d")
    fecha_label = now_cr.strftime("%Y-%m-%d")
    sesion = "13:05" if now_cr.hour < 15 else "17:00"

    if is_feriado(fecha_label):
        print(f"  ⓘ Hoy es feriado ({fecha_label}). MONEX no opera.")
        return None

    if not BCCR_API_BEARER_TOKEN:
        raise RuntimeError(
            "Falta configurar la variable secreta BCCR_API_BEARER_TOKEN."
        )

    response = requests.get(
        BCCR_API_URL,
        params={
            "codigo": BCCR_CODIGO_CUADRO,
            "fechaInicio": fecha_api,
            "fechafin": fecha_api,
            "idioma": "ES",
        },
        headers={
            "Accept": "application/json",
            "Authorization": f"Bearer {BCCR_API_BEARER_TOKEN}",
            "User-Agent": "CVFinanzas/1.0",
        },
        timeout=60,
    )
    response.raise_for_status()

    try:
        payload = response.json()
    except requests.exceptions.JSONDecodeError as exc:
        raise ValueError("La API del BCCR no devolvio JSON valido.") from exc

    return parse_bccr_response(
        payload,
        fecha_label=fecha_label,
        sesion=sesion,
        timestamp=now_cr.strftime("%Y-%m-%d %H:%M:%S"),
    )

def excel_serial_to_iso(val):
    """Convierte serial de Excel (número) o string de fecha a YYYY-MM-DD.
    Microsoft Graph API devuelve fechas en formato M/D/YYYY (formato americano)."""
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

    # Ya viene en formato YYYY-MM-DD — devolver directo
    if len(s) >= 10 and s[4] == '-' and s[7] == '-':
        return s[:10]

    # Microsoft Graph API devuelve M/D/YYYY (formato americano: mes primero)
    if "/" in s:
        parts = s.split("/")
        if len(parts) == 3:
            m, d, y = parts[0].zfill(2), parts[1].zfill(2), parts[2]
            if len(y) == 2:
                y = "20" + y
            # Validar que mes y día sean razonables
            mi, di = int(m), int(d)
            if 1 <= mi <= 12 and 1 <= di <= 31:
                return f"{y}-{m}-{d}"
            # Si mes > 12, probablemente viene como D/M/YYYY (invertido)
            if mi > 12 and 1 <= di <= 12:
                return f"{y}-{d}-{m}"
            return f"{y}-{m}-{d}"
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
    """Corrige fechas con día/mes invertido.
    Detecta: fechas futuras y fechas que caen en fin de semana (MONEX no opera sáb/dom).
    Si invertir día/mes produce una fecha válida en día laboral, la corrige.
    También valida que no sea feriado."""
    from datetime import date
    hoy = date.today().isoformat()
    try:
        partes = fecha_str.split("-")
        if len(partes) != 3:
            return fecha_str, None
        y, m, d = int(partes[0]), int(partes[1]), int(partes[2])
        dt = date(y, m, d)
        es_finde = dt.weekday() >= 5  # 5=sáb, 6=dom
        es_futura = fecha_str > hoy
        es_feriado_actual = is_feriado(fecha_str)

        # Solo intentar invertir si día <= 12 (sino no puede ser un mes válido)
        if (es_futura or es_finde or es_feriado_actual) and d <= 12:
            try:
                invertida = date(y, d, m)
                inv_str = invertida.isoformat()
                inv_laboral = invertida.weekday() < 5
                inv_pasada = inv_str <= hoy
                inv_no_feriado = not is_feriado(inv_str)

                if inv_pasada and inv_laboral and inv_no_feriado:
                    hora = sesion_str if sesion_str else "17:00"
                    return inv_str, f"{inv_str} {hora}"
            except ValueError:
                pass  # fecha invertida inválida (ej: mes 13)

    except (ValueError, IndexError):
        pass
    return fecha_str, None

def fix_ambiguous_dates(rows):
    """Corrige fechas ambiguas (mes<=12, día<=12) cuyo valor es un outlier
    respecto a sus vecinos. Si invertir mes/día hace que el valor encaje
    mejor en la serie, se intercambian o reubican los registros.
    Solo actúa sobre registros donde ambas interpretaciones caen en día laboral."""
    from datetime import date
    MAX_PASSES = 3  # máximo de pasadas para resolver intercambios cruzados

    for _pass in range(MAX_PASSES):
        cambios = 0
        n = len(rows)
        i = 1
        while i < n - 1:
            r = rows[i]
            prev_r = rows[i - 1]
            next_r = rows[i + 1]
            val = r["promedio_ponderado"]
            val_prev = prev_r["promedio_ponderado"]
            val_next = next_r["promedio_ponderado"]

            diff_prev = abs(val - val_prev)
            diff_next = abs(val - val_next)
            diff_neighbors = abs(val_prev - val_next)

            # Detectar outlier: vecinos similares entre sí pero este difiere mucho
            if diff_neighbors < 5 and (diff_prev > 10 or diff_next > 10):
                parts = r["fecha"].split("-")
                y, m, d = int(parts[0]), int(parts[1]), int(parts[2])

                # Solo ambiguas: mes y día ambos <= 12
                if m <= 12 and d <= 12 and m != d:
                    try:
                        inv = date(y, d, m)
                        inv_str = inv.isoformat()
                        inv_laboral = inv.weekday() < 5
                        hoy = date.today().isoformat()
                        inv_pasada = inv_str <= hoy

                        if inv_laboral and inv_pasada:
                            # Buscar si la fecha invertida ya existe
                            idx_inv = None
                            for j, rj in enumerate(rows):
                                if rj["fecha"] == inv_str:
                                    idx_inv = j
                                    break

                            if idx_inv is not None:
                                # Ambas fechas existen: intercambiar valores
                                campos = ["promedio_ponderado", "monto_total", "minimo", "maximo", "sesion"]
                                for c in campos:
                                    r[c], rows[idx_inv][c] = rows[idx_inv][c], r[c]
                                r["timestamp"] = r["fecha"] + " " + r.get("sesion", "17:00")
                                rows[idx_inv]["timestamp"] = rows[idx_inv]["fecha"] + " " + rows[idx_inv].get("sesion", "17:00")
                                print(f"  Intercambiando valores: {r['fecha']} <-> {inv_str}")
                                cambios += 1
                            else:
                                # La fecha invertida no existe: mover el registro (si no es feriado)
                                if not is_feriado(inv_str):
                                    hora = r.get("sesion", "17:00")
                                    print(f"  Moviendo fecha: {r['fecha']} → {inv_str}")
                                    r["fecha"] = inv_str
                                    r["timestamp"] = f"{inv_str} {hora}"
                                    cambios += 1
                    except ValueError:
                        pass
            i += 1

        # Reordenar después de cada pasada
        rows.sort(key=lambda x: str(x.get("timestamp", x.get("fecha", ""))))

        if cambios == 0:
            break
        print(f"  Pasada {_pass + 1}: {cambios} correcciones por outlier")

    return rows


def generate_json(all_rows):
    """Genera datos.json con el historial completo"""
    # Paso 1: Corregir fechas futuras y de fin de semana (día/mes invertidos)
    for r in all_rows:
        fecha_corregida, ts_corregido = fix_future_date(r["fecha"], r.get("sesion", "17:00"))
        if fecha_corregida != r["fecha"]:
            print(f"  Corrigiendo fecha (finde/futura): {r['fecha']} → {fecha_corregida}")
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

    # Paso 2: Corregir fechas ambiguas por continuidad de valores (outliers)
    deduped = fix_ambiguous_dates(deduped)
    output = {
        "actualizado": datetime.now(CR_TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "datos": deduped
    }
    with open("datos-json/datos.json", "w", encoding="utf-8") as f:
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
        row_data["timestamp"],
        row_data["fecha"]
    ]]
    r = requests.post(url, headers=headers, json={"values": values})
    if not r.ok:
        print(f"  ✗ Error {r.status_code}: {r.text}")
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
