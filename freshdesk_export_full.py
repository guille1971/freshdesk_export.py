"""
freshdesk_export_full.py
Descarga TODOS los tickets de Freshdesk y los guarda en un Excel.
El archivo se sube como artefacto en GitHub Actions para descarga directa.

Variables de entorno requeridas:
    FRESHDESK_DOMAIN   → tu subdominio (ej: bankinplay)
    FRESHDESK_API_KEY  → tu API key de Freshdesk
"""

import os
import time
import requests
import pandas as pd
from datetime import datetime, timezone

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

DOMAIN   = os.environ["FRESHDESK_DOMAIN"]
API_KEY  = os.environ["FRESHDESK_API_KEY"]
BASE_URL = f"https://{DOMAIN}.freshdesk.com/api/v2"
AUTH     = (API_KEY, "X")

AHORA = datetime.now(timezone.utc)

ESTADO      = {1: "Abierto", 2: "Pendiente", 3: "Resuelto", 4: "Cerrado", 5: "Esperando respuesta"}
PRIORIDAD   = {1: "Baja", 2: "Media", 3: "Alta", 4: "Urgente"}
TIPO_ORIGEN = {
    1: "Email", 2: "Portal", 3: "Teléfono", 7: "Chat",
    9: "Feedback", 10: "Twitter", 11: "Facebook", 13: "WhatsApp"
}

# ---------------------------------------------------------------------------
# Extracción
# ---------------------------------------------------------------------------

def get_all_tickets() -> list[dict]:
    all_tickets = []
    page = 1
    print("Extrayendo TODOS los tickets de Freshdesk...")

    while True:
        params = {
            "order_by":   "created_at",
            "order_type": "asc",
            "per_page":   100,
            "page":       page,
            "include":    "requester,stats",
        }
        res = requests.get(f"{BASE_URL}/tickets", auth=AUTH, params=params, timeout=30)
        res.raise_for_status()
        tickets = res.json()

        if not tickets:
            break

        all_tickets.extend(tickets)
        print(f"  Página {page}: {len(tickets)} tickets | Total: {len(all_tickets)}")

        if len(tickets) < 100:
            break

        page += 1
        time.sleep(0.5)

    print(f"✓ Total tickets extraídos: {len(all_tickets)}")
    return all_tickets


def get_agents() -> dict[int, str]:
    res = requests.get(f"{BASE_URL}/agents", auth=AUTH, timeout=30)
    res.raise_for_status()
    return {a["id"]: a["contact"]["name"] for a in res.json()}


def get_groups() -> dict[int, str]:
    res = requests.get(f"{BASE_URL}/groups", auth=AUTH, timeout=30)
    res.raise_for_status()
    return {g["id"]: g["name"] for g in res.json()}

# ---------------------------------------------------------------------------
# Transformación
# ---------------------------------------------------------------------------

def flatten_ticket(ticket: dict, agents: dict, groups: dict) -> dict:
    stats = ticket.get("stats", {})
    req   = ticket.get("requester", {})
    return {
        "ID":                ticket.get("id"),
        "Asunto":            ticket.get("subject"),
        "Estado":            ESTADO.get(ticket.get("status"), ticket.get("status")),
        "Prioridad":         PRIORIDAD.get(ticket.get("priority"), ticket.get("priority")),
        "Origen":            TIPO_ORIGEN.get(ticket.get("source"), ticket.get("source")),
        "Agente":            agents.get(ticket.get("responder_id"), "Sin asignar"),
        "Grupo":             groups.get(ticket.get("group_id"), "Sin grupo"),
        "Solicitante":       req.get("name", ""),
        "Email solicitante": req.get("email", ""),
        "Etiquetas":         ", ".join(ticket.get("tags", [])),
        "Creado":            ticket.get("created_at", "")[:19].replace("T", " "),
        "Actualizado":       ticket.get("updated_at", "")[:19].replace("T", " "),
        "Primera respuesta": (stats.get("first_responded_at") or "")[:19].replace("T", " "),
        "Resuelto en":       (stats.get("resolved_at") or "")[:19].replace("T", " "),
        "Cerrado en":        (stats.get("closed_at") or "")[:19].replace("T", " "),
        "Reabierto":         stats.get("reopened_at") is not None,
        "URL":               f"https://{DOMAIN}.freshdesk.com/helpdesk/tickets/{ticket.get('id')}",
    }

# ---------------------------------------------------------------------------
# Exportación a Excel
# ---------------------------------------------------------------------------

def export_to_excel(rows: list[dict]) -> str:
    fecha_str = AHORA.strftime("%Y-%m-%d")
    filename  = f"freshdesk_todos_tickets_{fecha_str}.xlsx"
    df = pd.DataFrame(rows)

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tickets")
        ws = writer.sheets["Tickets"]
        for col in ws.columns:
            max_len = max(
                len(str(col[0].value or "")),
                *(len(str(cell.value or "")) for cell in col[1:])
            )
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    print(f"✓ Archivo generado: {filename}")
    return filename

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Cargando agentes y grupos...")
    agents = get_agents()
    groups = get_groups()

    tickets  = get_all_tickets()
    rows     = [flatten_ticket(t, agents, groups) for t in tickets]
    export_to_excel(rows)


if __name__ == "__main__":
    main()

