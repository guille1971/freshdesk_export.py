"""
freshdesk_export.py
Extrae todos los tickets de Freshdesk del día anterior y los guarda
en un archivo Excel (.xlsx) con la fecha en el nombre.

Uso:
    python freshdesk_export.py

Variables de entorno requeridas:
    FRESHDESK_DOMAIN   → tu subdominio (ej: bankinplay → bankinplay.freshdesk.com)
    FRESHDESK_API_KEY  → tu API key de Freshdesk
"""

import os
import time
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

DOMAIN    = os.environ["FRESHDESK_DOMAIN"]       # ej: "bankinplay"
API_KEY   = os.environ["FRESHDESK_API_KEY"]
BASE_URL  = f"https://{DOMAIN}.freshdesk.com/api/v2"
AUTH      = (API_KEY, "X")                        # Freshdesk usa API key como usuario

# Rango: ayer 00:00 → hoy 00:00 (UTC)
HOY       = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
AYER      = HOY - timedelta(days=1)

# Mapas de traducción de códigos → texto legible
ESTADO = {1: "Abierto", 2: "Pendiente", 3: "Resuelto", 4: "Cerrado", 5: "Esperando respuesta"}
PRIORIDAD = {1: "Baja", 2: "Media", 3: "Alta", 4: "Urgente"}
TIPO_ORIGEN = {
    1: "Email", 2: "Portal", 3: "Teléfono", 7: "Chat",
    9: "Feedback", 10: "Twitter", 11: "Facebook", 13: "WhatsApp"
}

# ---------------------------------------------------------------------------
# Funciones de extracción
# ---------------------------------------------------------------------------

def get_tickets_page(page: int, per_page: int = 100) -> list[dict]:
    """Obtiene una página de tickets actualizados desde AYER."""
    params = {
        "updated_since": AYER.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "order_by":      "updated_at",
        "order_type":    "asc",
        "per_page":      per_page,
        "page":          page,
        # Incluir campos extra
        "include":       "requester,stats",
    }
    res = requests.get(f"{BASE_URL}/tickets", auth=AUTH, params=params, timeout=30)
    res.raise_for_status()
    return res.json()


def get_all_tickets() -> list[dict]:
    """Pagina automáticamente hasta obtener todos los tickets del rango."""
    all_tickets = []
    page = 1

    print(f"Extrayendo tickets desde {AYER.date()} hasta {HOY.date()}...")

    while True:
        tickets = get_tickets_page(page)
        if not tickets:
            break

        # Filtrar solo los tickets actualizados dentro del rango (ayer)
        en_rango = [
            t for t in tickets
            if AYER <= datetime.fromisoformat(t["updated_at"].replace("Z", "+00:00")) < HOY
        ]
        all_tickets.extend(en_rango)

        print(f"  Página {page}: {len(tickets)} tickets obtenidos, {len(en_rango)} en rango")

        # Freshdesk limita a 300 páginas y aplica rate limit
        if len(tickets) < 100:
            break

        page += 1
        time.sleep(0.5)  # Respetar rate limit (500 req/min en plan Growth+)

    print(f"Total tickets extraídos: {len(all_tickets)}")
    return all_tickets


def get_agents() -> dict[int, str]:
    """Devuelve un mapa {agent_id: nombre} para enriquecer los tickets."""
    res = requests.get(f"{BASE_URL}/agents", auth=AUTH, timeout=30)
    res.raise_for_status()
    return {a["id"]: a["contact"]["name"] for a in res.json()}


def get_groups() -> dict[int, str]:
    """Devuelve un mapa {group_id: nombre}."""
    res = requests.get(f"{BASE_URL}/groups", auth=AUTH, timeout=30)
    res.raise_for_status()
    return {g["id"]: g["name"] for g in res.json()}


# ---------------------------------------------------------------------------
# Transformación
# ---------------------------------------------------------------------------

def flatten_ticket(ticket: dict, agents: dict, groups: dict) -> dict:
    """Convierte un ticket de la API en una fila plana para el Excel."""
    stats    = ticket.get("stats", {})
    req      = ticket.get("requester", {})

    return {
        "ID":                   ticket.get("id"),
        "Asunto":               ticket.get("subject"),
        "Estado":               ESTADO.get(ticket.get("status"), ticket.get("status")),
        "Prioridad":            PRIORIDAD.get(ticket.get("priority"), ticket.get("priority")),
        "Origen":               TIPO_ORIGEN.get(ticket.get("source"), ticket.get("source")),
        "Agente":               agents.get(ticket.get("responder_id"), "Sin asignar"),
        "Grupo":                groups.get(ticket.get("group_id"), "Sin grupo"),
        "Solicitante":          req.get("name", ""),
        "Email solicitante":    req.get("email", ""),
        "Etiquetas":            ", ".join(ticket.get("tags", [])),
        "Creado":               ticket.get("created_at", "")[:19].replace("T", " "),
        "Actualizado":          ticket.get("updated_at", "")[:19].replace("T", " "),
        "Primera respuesta":    (stats.get("first_responded_at") or "")[:19].replace("T", " "),
        "Resuelto en":          (stats.get("resolved_at") or "")[:19].replace("T", " "),
        "Cerrado en":           (stats.get("closed_at") or "")[:19].replace("T", " "),
        "Reabierto":            stats.get("reopened_at") is not None,
        "URL":                  f"https://{DOMAIN}.freshdesk.com/helpdesk/tickets/{ticket.get('id')}",
    }


# ---------------------------------------------------------------------------
# Exportación a Excel
# ---------------------------------------------------------------------------

def export_to_excel(rows: list[dict]) -> str:
    """Guarda los tickets en un Excel con formato y devuelve la ruta del archivo."""
    fecha_str = AYER.strftime("%Y-%m-%d")
    filename  = f"freshdesk_tickets_{fecha_str}.xlsx"

    df = pd.DataFrame(rows)

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tickets")

        # Ajustar ancho de columnas automáticamente
        ws = writer.sheets["Tickets"]
        for col in ws.columns:
            max_len = max(
                len(str(col[0].value or "")),
                *(len(str(cell.value or "")) for cell in col[1:])
            )
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    print(f"Archivo generado: {filename}")
    return filename


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    # 1. Obtener datos de referencia
    print("Cargando agentes y grupos...")
    agents = get_agents()
    groups = get_groups()

    # 2. Extraer tickets
    tickets = get_all_tickets()

    if not tickets:
        print(f"No hay tickets actualizados el {AYER.date()}. No se genera archivo.")
        return

    # 3. Aplanar datos
    rows = [flatten_ticket(t, agents, groups) for t in tickets]

    # 4. Exportar
    export_to_excel(rows)


if __name__ == "__main__":
    main()
