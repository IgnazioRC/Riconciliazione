"""
ms_constants.py — Costanti, etichette e configurazione default
Parte di: Moneyspire Reconciler
"""

import re
from datetime import date, datetime
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# COSTANTI MATCH
# ─────────────────────────────────────────────────────────────────────────────

MATCH_EXACT  = "exact"
MATCH_FUZZY  = "fuzzy"
MATCH_SPLIT  = "split"
MATCH_MERGE  = "merge"
MATCH_NONE    = "none"
MATCH_SKIP    = "skip"
MATCH_PENDING = "pending"   # transazione a cavallo mese — riconciliare il mese dopo

STATO_LABELS = {
    MATCH_EXACT:  ("✅ Trovata",         "#d4edda"),
    MATCH_FUZZY:  ("🔶 Trovata (±gg)",   "#fff3cd"),
    MATCH_SPLIT:  ("🔀 Split",           "#cce5ff"),
    MATCH_MERGE:  ("🔗 Merge/Composta",  "#e2ccff"),
    MATCH_NONE:    ("❌ Mancante",        "#f8d7da"),
    MATCH_SKIP:    ("📝 Solo Money",      "#e2e3e5"),
    MATCH_PENDING: ("⏳ In attesa",       "#fde8c8"),
}

DEFAULT_CONFIG = {
    "date_tolerance_days": 3,
    "amount_tolerance": 0.01,
    "conti": {
        "Fineco": {
            "ffd_account_id": 536,
            "tipo": "conto_corrente",
            "excel_sheet": "Movimenti",
            "excel_columns": {
                "data": "Data", "entrate": "Entrate", "uscite": "Uscite",
                "descrizione": "Descrizione",
                "descrizione_completa": "Descrizione_Completa"
            },
            "frequenza": "mensile"
        },
        "Lombard": {
            "ffd_account_id": 577,
            "tipo": "conto_corrente",
            "excel_sheet": "Lombard",
            "excel_columns": {
                "data": "Data", "entrate": "Entrate", "uscite": "Uscite",
                "descrizione": "Descrizione",
                "descrizione_completa": "Descrizione_Completa"
            },
            "frequenza": "mensile"
        },
        "USD": {
            "ffd_account_id": 554,
            "tipo": "conto_corrente",
            "excel_sheet": "USD",
            "excel_columns": {
                "data": "Data", "entrate": "Entrate", "uscite": "Uscite",
                "descrizione": "Descrizione",
                "descrizione_completa": "Descrizione_Completa"
            },
            "frequenza": "mensile"
        },
        "MC Fineco": {
            "ffd_account_id": 531,
            "tipo": "carta_credito",
            "numero_carta": "5260",
            "frequenza": "mensile"
        },
        "Visa Fineco": {
            "ffd_account_id": 532,
            "tipo": "carta_credito",
            "numero_carta": "6421",
            "frequenza": "mensile"
        },
        "Unicredit CCM": {
            "ffd_account_id": 548,
            "tipo": "conto_corrente",
            "tipo_file_analisi": "originale_unicredit",
            "frequenza": "variabile",
            "includi_in_riconciliazione": True
        }
    }
}


# ─────────────────────────────────────────────────────────────────────────────
# UTILITÀ
# ─────────────────────────────────────────────────────────────────────────────

def _to_date(val) -> date | None:
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
            try:
                return datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                continue
    return None


def _to_float(val) -> float:
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if re.match(r'^-?[\d\.]+,\d{1,2}$', s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def fmt_eur(val: float) -> str:
    """Formatta importo in stile italiano: +1.234,56 €"""
    sign = "+" if val >= 0 else "-"
    s = f"{abs(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sign}{s} €"


def _amt_eq(a: float, b: float, tol: float = 0.01) -> bool:
    """Confronto importi con tolleranza assoluta."""
    return abs(abs(a) - abs(b)) <= tol


def _date_ok(d1: "date", d2: "date", tol: int = 3) -> bool:
    """Verifica che due date siano entro la tolleranza in giorni."""
    return abs((d1 - d2).days) <= tol
