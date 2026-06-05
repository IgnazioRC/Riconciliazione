"""
ms_engine.py — Wrapper di compatibilità
Re-esporta tutti i simboli pubblici dai sottomoduli.
Il reconciler e qualsiasi altro script possono continuare a importare
da ms_engine senza modifiche.

Struttura interna:
  ms_constants.py  — costanti MATCH_*, STATO_LABELS, DEFAULT_CONFIG, utility
  ms_db.py         — MoneyspireDB, MoneyWriter, backup/verifica/finalizza
  ms_parsers.py    — parse_fineco_*, leggi_saldo_*, valida_file_banca
  ms_matching.py   — ReconcileEngine, RulesEngine, raggruppa_*, export_*
  ms_excel.py      — ExcelIntegrator, leggi_excel_elaborato_*
"""

# ── Costanti e utility ────────────────────────────────────────────────────────
from ms_constants import (
    MATCH_EXACT, MATCH_FUZZY, MATCH_SPLIT, MATCH_MERGE,
    MATCH_NONE, MATCH_SKIP, MATCH_PENDING,
    STATO_LABELS, DEFAULT_CONFIG,
    _to_date, _to_float, fmt_eur,
)

# ── Database Moneyspire ───────────────────────────────────────────────────────
from ms_db import (
    MoneyspireDB,
    MoneyWriter,
    prepara_db_scrittura,
    verifica_scrittura,
    finalizza_db,
)

# ── Parser file banca ─────────────────────────────────────────────────────────
from ms_parsers import (
    parse_fineco_conto_originale,
    parse_fineco_conto,
    parse_fineco_conto_usd,
    parse_fineco_cc,
    parse_fineco_cc_cartella,
    parse_unicredit_ccm,
    leggi_intestazione_unicredit,
    leggi_numero_conto_fineco,
    leggi_saldo_fineco,
    leggi_saldo_money,
    leggi_variazione_mensile_fineco,
    leggi_variazione_mensile_money,
)

# ── Parser profilo Silvia (SC) ────────────────────────────────────────────────
from ms_parsers_silvia import (
    parse_bper,
    leggi_intestazione_bper,
    parse_nexi_xlsx,
    leggi_intestazione_nexi,
)

# ── Motore di matching e regole ───────────────────────────────────────────────
from ms_matching import (
    ReconcileEngine,
    RulesEngine,
    valida_file_banca,
    marca_in_attesa,
    marca_solo_money_in_attesa,
    raggruppa_cedole_ritenute,
    fondi_cedole_per_match,
    raggruppa_transazioni_correlate,
    costruisci_transazioni_da_risultati,
    export_full_csv,
    export_missing_csv,
)

# ── Fase 2: Excel elaborati ───────────────────────────────────────────────────
from ms_excel import (
    ExcelIntegrator,
    MESI_IT_NOMI,
    leggi_excel_elaborato_cc,
    leggi_excel_elaborato_cc_mensile,
)

__all__ = [
    # Costanti
    "MATCH_EXACT", "MATCH_FUZZY", "MATCH_SPLIT", "MATCH_MERGE",
    "MATCH_NONE", "MATCH_SKIP", "MATCH_PENDING",
    "STATO_LABELS", "DEFAULT_CONFIG", "fmt_eur",
    # DB
    "MoneyspireDB", "MoneyWriter",
    "prepara_db_scrittura", "verifica_scrittura", "finalizza_db",
    # Parsers Fineco / Unicredit
    "parse_fineco_conto_originale", "parse_fineco_conto",
    "parse_fineco_conto_usd", "parse_fineco_cc", "parse_fineco_cc_cartella",
    "parse_unicredit_ccm", "leggi_intestazione_unicredit",
    "leggi_numero_conto_fineco",
    "leggi_saldo_fineco", "leggi_saldo_money",
    "leggi_variazione_mensile_fineco", "leggi_variazione_mensile_money",
    # Parsers Silvia (SC): BPER e Nexi xlsx
    "parse_bper", "leggi_intestazione_bper",
    "parse_nexi_xlsx", "leggi_intestazione_nexi",
    # Matching
    "ReconcileEngine", "RulesEngine",
    "valida_file_banca", "marca_in_attesa", "marca_solo_money_in_attesa",
    "raggruppa_cedole_ritenute", "fondi_cedole_per_match",
    "raggruppa_transazioni_correlate",
    "costruisci_transazioni_da_risultati",
    "export_full_csv", "export_missing_csv",
    # Excel
    "ExcelIntegrator", "MESI_IT_NOMI",
    "leggi_excel_elaborato_cc", "leggi_excel_elaborato_cc_mensile",
]
