#!/usr/bin/env python3
"""
Verifica rapida della riconciliazione del conto USD.
Da lanciare sul Mac, nella cartella dell'app (dove stanno i moduli ms_*.py),
con la venv attiva:

    pystable                       # attiva la venv
    python3 verifica_usd.py  /percorso/Ignazio.ffd  /percorso/usd.xlsx  554

Dice subito (1) se le patch sono caricate e (2) quante righe banca trovano match.
Non scrive nulla: legge soltanto.
"""
import sys
from datetime import date

def main():
    if len(sys.argv) < 4:
        print("Uso: python3 verifica_usd.py <file.ffd> <usd.xlsx> <account_id_USD>")
        print("     (account_id del conto 'Fineco USD' — nel tuo DB e' 554)")
        return

    ffd, xlsx, acct = sys.argv[1], sys.argv[2], int(sys.argv[3])

    # 1) Le patch sono presenti nei moduli?
    import os, inspect
    import ms_matching
    ok_fondi = hasattr(ms_matching, "fondi_cedole_per_match")

    recon_path = os.path.join(os.path.dirname(os.path.abspath(ms_matching.__file__)),
                              "ms_reconciler.py")
    try:
        with open(recon_path, encoding="utf-8") as f:
            src_recon = f.read()
    except Exception:
        src_recon = ""
    ok_chiamata = "fondi_cedole_per_match(bank_txns)" in src_recon
    ok_ramo_usd = 'tipo_file_analisi == "originale_fineco_usd"' in src_recon

    print("── Stato patch ──────────────────────────────")
    print(f"  fondi_cedole_per_match definita : {'SI' if ok_fondi else 'NO  <-- modulo vecchio!'}")
    print(f"  chiamata nel reconciler         : {'SI' if ok_chiamata else 'NO  <-- reconciler vecchio!'}")
    print(f"  ramo parser USD nel reconciler  : {'SI' if ok_ramo_usd else 'NO  <-- reconciler vecchio!'}")
    print(f"  modulo ms_matching da           : {ms_matching.__file__}")
    print(f"  ms_reconciler letto da          : {recon_path}")
    if not (ok_fondi and ok_chiamata and ok_ramo_usd):
        print("\n  >> L'app sta usando file NON aggiornati. Sostituisci i .py,")
        print("     cancella le cartelle __pycache__, e se usi un .app ricompila.")
        return

    # 2) Prova di matching sui dati reali
    from ms_db import MoneyspireDB
    from ms_parsers import parse_fineco_conto_usd
    from ms_matching import (raggruppa_cedole_ritenute, fondi_cedole_per_match,
                             ReconcileEngine, MATCH_NONE)

    db = MoneyspireDB(ffd)
    # finestra ampia: tutto il primo semestre, poi filtra il parser per mese
    money = db.get_transactions(acct, date(2026, 1, 1), date(2026, 6, 30))

    print("\n── Prova matching (febbraio 2026) ───────────")
    bank = parse_fineco_conto_usd(xlsx, 2, 2026)
    bank = raggruppa_cedole_ritenute(bank)
    bank = fondi_cedole_per_match(bank)
    eng = ReconcileEngine(db, {})
    res = eng.reconcile(bank, money, account_id=acct)

    n_ok = sum(1 for r in res if r["bank_txn"] and r["match_type"] != MATCH_NONE)
    n_no = sum(1 for r in res if r["bank_txn"] and r["match_type"] == MATCH_NONE)
    print(f"  Righe banca: {n_ok+n_no}  →  abbinate: {n_ok}  |  non trovate: {n_no}")
    for r in res:
        bt = r["bank_txn"]
        if not bt:
            continue
        esito = "—" if r["match_type"] == MATCH_NONE else "OK"
        print(f"    {bt['date']}  {bt['amount']:>10.2f}  {esito}")

if __name__ == "__main__":
    main()
