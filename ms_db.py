"""
ms_db.py — Accesso al database Moneyspire (.ffd / SQLite)
          MoneyspireDB  — lettura sola lettura
          MoneyWriter   — scrittura su copia di lavoro
          prepara_db_scrittura / verifica_scrittura / finalizza_db
Parte di: Moneyspire Reconciler
"""

import sqlite3
import json
import re
import shutil as _shutil
from pathlib import Path
from datetime import date, datetime
from collections import Counter

from ms_constants import (
    MATCH_EXACT, MATCH_FUZZY, MATCH_SPLIT, MATCH_MERGE,
    MATCH_NONE, MATCH_SKIP, MATCH_PENDING,
    DEFAULT_CONFIG, fmt_eur,
)

# ─────────────────────────────────────────────────────────────────────────────
# DATABASE MONEYSPIRE
# ─────────────────────────────────────────────────────────────────────────────

class MoneyspireDB:
    """Accesso in SOLA LETTURA al database .ffd (SQLite) di Moneyspire."""

    def __init__(self, ffd_path: str):
        self.path = ffd_path
        self._conn: sqlite3.Connection | None = None
        self._cat_map: dict[int, str] = {}
        self._payee_map: dict[int, str] = {}
        self._account_map: dict[int, str] = {}
        self._load_maps()

    def _get_conn(self) -> sqlite3.Connection:
        if self._conn is None:
            uri = f"file:{self.path}?mode=ro"
            self._conn = sqlite3.connect(uri, uri=True)
            self._conn.row_factory = sqlite3.Row
        return self._conn

    def close(self):
        if self._conn:
            self._conn.close()
            self._conn = None

    def _load_maps(self):
        conn = self._get_conn()
        cur = conn.cursor()
        cur.execute("SELECT ID, Name, ParentCategoryID FROM Accounts WHERE Expense=1 OR Type=10")
        rows = cur.fetchall()
        raw: dict[int, tuple] = {r["ID"]: (r["Name"] or "", r["ParentCategoryID"]) for r in rows}
        for cid, (name, parent) in raw.items():
            self._cat_map[cid] = f"{raw[parent][0]}:{name}" if parent and parent in raw else name
        cur.execute("SELECT ID, Name FROM Accounts WHERE Type BETWEEN 0 AND 8")
        self._account_map = {r["ID"]: (r["Name"] or "") for r in cur.fetchall()}
        cur.execute("SELECT ID, Name FROM Payees")
        self._payee_map = {r["ID"]: (r["Name"] or "") for r in cur.fetchall()}

    def cat_name(self, cid) -> str:
        if cid is None:
            return ""
        cid = int(cid)
        # Prima cerca nelle categorie, poi nei conti (trasferimenti)
        if cid in self._cat_map:
            return self._cat_map[cid]
        if cid in self._account_map:
            return f"→ {self._account_map[cid]}"  # es. "→ Visa Fineco"
        return f"[{cid}]"

    def payee_name(self, pid) -> str:
        return "" if pid is None else self._payee_map.get(int(pid), "")

    def get_categories(self) -> dict[int, str]:
        return dict(self._cat_map)

    def get_payees(self) -> dict[int, str]:
        return dict(self._payee_map)

    def get_accounts(self) -> dict[int, str]:
        return dict(self._account_map)

    def get_payee_categories(self) -> dict[str, int]:
        """
        Ritorna mappa payee_name → LastCategoryID.
        Usata da DialogModificaCategoria per pre-compilare la categoria
        quando si seleziona un payee già noto in Money.
        """
        conn = self._get_conn()
        cur  = conn.cursor()
        cur.execute("""
            SELECT p.Name, p.LastCategoryID
            FROM Payees p
            WHERE p.LastCategoryID IS NOT NULL
        """)
        result = {}
        for row in cur.fetchall():
            name = row[0] or ""
            if name:
                result[name] = int(row[1])
        return result

    def get_transactions(self, account_id: int,
                         date_from: date | None = None,
                         date_to: date | None = None) -> list[dict]:
        conn = self._get_conn()
        cur = conn.cursor()
        where = ["t.AccountID = ?"]
        params: list = [account_id]
        if date_from:
            where.append("t.TransactionDate >= ?")
            params.append(date_from.isoformat())
        if date_to:
            where.append("t.TransactionDate <= ?")
            params.append(date_to.isoformat())
        # Usa DATE() per normalizzare il confronto — il DB può avere
        # timestamp '2026-03-31 00:00:00' che confrontato come stringa
        # risulta > '2026-03-31' escludendo erroneamente l'ultimo giorno
        where_sql = " AND ".join(where)
        where_sql = where_sql.replace(
            "t.TransactionDate >=", "DATE(t.TransactionDate) >=").replace(
            "t.TransactionDate <=", "DATE(t.TransactionDate) <=")
        cur.execute(f"""
            SELECT t.ID, t.TransactionDate, t.Withdrawal, t.Deposit,
                   t.Memo, t.CategoryID, t.PayeeID, t.Status
            FROM Transactions t
            WHERE {where_sql}
            ORDER BY t.TransactionDate, t.ID
        """, params)
        result = []
        for row in cur.fetchall():
            cur2 = conn.cursor()
            cur2.execute("""
                SELECT ID, CategoryID, Withdrawal, Deposit, Memo
                FROM Splits WHERE TransactionID = ?
            """, (row["ID"],))
            splits = [{
                "id": s["ID"], "category": self.cat_name(s["CategoryID"]),
                "category_id": s["CategoryID"],
                "withdrawal": float(s["Withdrawal"] or 0),
                "deposit": float(s["Deposit"] or 0),
                "memo": s["Memo"] or ""
            } for s in cur2.fetchall()]

            raw_dt = row["TransactionDate"]
            if isinstance(raw_dt, str):
                dt = date.fromisoformat(raw_dt[:10])
            elif isinstance(raw_dt, datetime):
                dt = raw_dt.date()
            else:
                dt = raw_dt
            if not isinstance(dt, date):
                continue

            w = float(row["Withdrawal"] or 0)
            d = float(row["Deposit"] or 0)
            result.append({
                "id": row["ID"], "date": dt,
                "withdrawal": w, "deposit": d, "amount": d - w,
                "memo": row["Memo"] or "",
                "category": self.cat_name(row["CategoryID"]),
                "category_id": row["CategoryID"],
                "payee": self.payee_name(row["PayeeID"]),
                "payee_id": row["PayeeID"],
                "status": row["Status"],
                "splits": splits, "has_splits": len(splits) > 0
            })
        return result




# ─────────────────────────────────────────────────────────────────────────────
# SCRITTURA SU DATABASE MONEYSPIRE
# ─────────────────────────────────────────────────────────────────────────────

import shutil as _shutil

# Nomi canonici categorie per splits automatici
# Gli ID vengono risolti a runtime da MoneyWriter._resolve_special_cats()
_CAT_CEDOLE_NAME    = "cedole"
_CAT_DIVIDENDI_NAME = "Dividendi"
_CAT_RITENUTA_NAME  = "ritenuta 26%"

# Valori di fallback (ID noti nel DB di Ignazio, usati se la ricerca per nome fallisce)
_CAT_CEDOLE    = 670
_CAT_DIVIDENDI = 292
_CAT_RITENUTA  = 570


def _next_id(cur, table: str) -> int:
    cur.execute(f"SELECT COALESCE(MAX(ID),0)+1 FROM {table}")
    return cur.fetchone()[0]


def prepara_db_scrittura(ffd_path: str, profilo: str = "") -> tuple[str, str]:
    """
    Prepara il DB per la scrittura sicura:
    1. Verifica che Moneyspire non sia in esecuzione
    2. Crea backup timestampato del file originale dentro la sottocartella
       _Backup_<PROFILO> (creata se non esiste) accanto al file originale.
       Se profilo è vuoto, usa "_Backup" senza suffisso.
    3. Crea copia di lavoro su cui scrivere (nella cartella originale)
    Ritorna (path_copia_lavoro, path_backup).
    """
    import subprocess, datetime
    from pathlib import Path

    result = subprocess.run(["pgrep", "-ix", "Moneyspire"],
                            capture_output=True, text=True)
    if result.returncode == 0:
        raise RuntimeError(
            "Moneyspire è in esecuzione. Chiudilo prima di procedere.")

    p  = Path(ffd_path)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    backup_dir_name = f"_Backup_{profilo}" if profilo else "_Backup"
    backup_dir = p.parent / backup_dir_name
    backup_dir.mkdir(exist_ok=True)

    backup = backup_dir / f"{p.stem}_backup_{ts}{p.suffix}"
    work   = p.parent  / f"{p.stem}_work_{ts}{p.suffix}"

    _shutil.copy2(ffd_path, backup)  # backup intatto
    _shutil.copy2(ffd_path, work)    # copia di lavoro
    return str(work), str(backup)


def verifica_scrittura(work_path: str, account_id: int,
                       date_from: "date", date_to: "date",
                       n_attese: int,
                       expected_txns: list[dict] | None = None) -> tuple[bool, str]:
    """
    Verifica la copia di lavoro dopo la scrittura Moneyspire.

    Compatibile con la chiamata storica della GUI:
        verifica_scrittura(work_path, account_id, date_from, date_to, n_attese)

    Se expected_txns viene passato, esegue anche una verifica forte
    sulle singole transazioni attese usando conto, data, importo e, quando
    disponibile, una parte del memo/descrizione.

    Ritorna:
        (ok, messaggio)
    """
    import sqlite3
    from datetime import date as _date, datetime as _datetime

    def _as_iso_day(value) -> str:
        if isinstance(value, _datetime):
            return value.date().isoformat()
        if isinstance(value, _date):
            return value.isoformat()
        txt = str(value or "")
        return txt[:10]

    def _amount_of(row) -> float:
        return round(float(row["Deposit"] or 0) - float(row["Withdrawal"] or 0), 2)

    def _expected_amount(t: dict) -> float:
        return round(float(t.get("deposit") or 0) - float(t.get("withdrawal") or 0), 2)

    conn = sqlite3.connect(work_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    try:
        cur.execute("""
            SELECT ID, AccountID, TransactionDate, Withdrawal, Deposit,
                   Memo, CategoryID, PayeeID
            FROM Transactions
            WHERE AccountID=?
              AND DATE(TransactionDate)>=?
              AND DATE(TransactionDate)<=?
            ORDER BY TransactionDate, ID
        """, (account_id, date_from.isoformat(), date_to.isoformat()))
        rows = cur.fetchall()

        n = len(rows)

        # Estratti carte di credito: l'estratto di un mese può contenere operazioni
        # degli ultimi giorni del mese precedente, registrate/addebittate nel mese corrente.
        # La GUI storica verifica solo il mese selezionato; quindi, se il conteggio
        # stretto fallisce, riproviamo allargando prudentemente l'inizio periodo di 5 giorni.
        expanded_rows = rows
        expanded_from = date_from
        if n < n_attese:
            from datetime import timedelta
            expanded_from = date_from - timedelta(days=5)
            cur.execute("""
                SELECT ID, AccountID, TransactionDate, Withdrawal, Deposit,
                       Memo, CategoryID, PayeeID
                FROM Transactions
                WHERE AccountID=?
                  AND DATE(TransactionDate)>=?
                  AND DATE(TransactionDate)<=?
                ORDER BY TransactionDate, ID
            """, (account_id, expanded_from.isoformat(), date_to.isoformat()))
            expanded_rows = cur.fetchall()
            n_expanded = len(expanded_rows)
            if n_expanded < n_attese:
                return False, (
                    f"Verifica FALLITA: nel periodo {date_from.isoformat()}–"
                    f"{date_to.isoformat()} attese almeno {n_attese}, trovate {n}. "
                    f"Anche estendendo dal {expanded_from.isoformat()} trovate {n_expanded}."
                )
            rows = expanded_rows
            n = n_expanded

        # Modalità storica: controllo quantitativo minimo.
        if not expected_txns:
            if expanded_from != date_from:
                return True, (
                    f"Verifica OK: {n} transazioni trovate nel periodo esteso "
                    f"{expanded_from.isoformat()}–{date_to.isoformat()} "
                    f"(attese almeno {n_attese}). Nota: alcune operazioni carta "
                    f"appartengono agli ultimi giorni del mese precedente."
                )
            return True, (
                f"Verifica OK: {n} transazioni trovate nel periodo "
                f"(attese almeno {n_attese})"
            )

        unmatched = []
        used_ids: set[int] = set()

        for t in expected_txns:
            exp_date = _as_iso_day(t.get("txn_date") or t.get("date"))
            exp_amount = _expected_amount(t)
            exp_memo = (
                t.get("memo")
                or t.get("_descrizione_banca")
                or t.get("descrizione")
                or ""
            ).strip().lower()
            exp_cat = t.get("category_id")
            exp_payee = t.get("payee_id")

            candidates = []
            for row in rows:
                rid = int(row["ID"])
                if rid in used_ids:
                    continue
                if _as_iso_day(row["TransactionDate"]) != exp_date:
                    continue
                if abs(_amount_of(row) - exp_amount) > 0.01:
                    continue
                if exp_cat is not None and row["CategoryID"] != exp_cat:
                    continue
                if exp_payee is not None and row["PayeeID"] != exp_payee:
                    continue

                # Il memo Moneyspire può essere più corto/lungo della descrizione
                # Fineco: usiamo solo un controllo morbido, non bloccante.
                db_memo = str(row["Memo"] or "").strip().lower()
                memo_score = 0
                if exp_memo and db_memo:
                    if exp_memo in db_memo or db_memo in exp_memo:
                        memo_score = 2
                    elif exp_memo[:20] and exp_memo[:20] in db_memo:
                        memo_score = 1
                candidates.append((memo_score, rid))

            if candidates:
                candidates.sort(reverse=True)
                used_ids.add(candidates[0][1])
            else:
                unmatched.append(t)

        if unmatched:
            esempi = []
            for t in unmatched[:5]:
                esempi.append(
                    f"{_as_iso_day(t.get('txn_date') or t.get('date'))} "
                    f"{_expected_amount(t):.2f} "
                    f"{str(t.get('_descrizione_banca') or t.get('memo') or '')[:40]}"
                )
            extra = "; ".join(esempi)
            return False, (
                f"Verifica FALLITA: {len(unmatched)} transazioni attese "
                f"non ritrovate puntualmente. Esempi: {extra}"
            )

        return True, (
            f"Verifica OK forte: {len(expected_txns)} transazioni attese "
            f"ritrovate puntualmente; {n} transazioni totali nel periodo"
        )

    finally:
        conn.close()

def finalizza_db(ffd_originale: str, work_path: str) -> bool:
    """
    Sostituisce il file originale con la copia di lavoro verificata.
    Il backup rimane intatto come sicurezza.
    """
    from pathlib import Path
    Path(work_path).replace(ffd_originale)
    return True


class MoneyWriter:
    """
    Scrive transazioni sul database Moneyspire (.ffd).
    Opera sempre su una copia di lavoro, mai sull'originale direttamente.
    """

    def __init__(self, db_path: str):
        self.path = db_path
        self._conn = sqlite3.connect(db_path)
        self._conn.row_factory = sqlite3.Row
        # Carica mappa categoria nome→ID
        cur = self._conn.cursor()
        cur.execute("SELECT ID, Name, ParentCategoryID FROM Accounts WHERE Expense=1 OR Type=10")
        rows = cur.fetchall()
        raw = {r["ID"]: (r["Name"], r["ParentCategoryID"]) for r in rows}
        self._cat_by_name: dict[str, int] = {}
        for cid, (name, parent) in raw.items():
            full = f"{raw[parent][0]}:{name}" if parent and parent in raw else name
            self._cat_by_name[full.lower()] = cid
            self._cat_by_name[name.lower()]  = cid
        # Mappa conti
        cur.execute("SELECT ID, Name FROM Accounts WHERE Type BETWEEN 0 AND 8")
        self._account_by_name = {r["Name"].lower(): r["ID"] for r in cur.fetchall()}
        # Risolve ID categorie speciali per nome (sovrascrive i valori hardcoded)
        self._resolve_special_cats()

    def _resolve_special_cats(self):
        """Risolve gli ID delle categorie speciali cercando per nome nel DB."""
        global _CAT_CEDOLE, _CAT_DIVIDENDI, _CAT_RITENUTA
        if cid := self._cat_by_name.get(_CAT_CEDOLE_NAME.lower()):
            _CAT_CEDOLE = cid
        if cid := self._cat_by_name.get(_CAT_DIVIDENDI_NAME.lower()):
            _CAT_DIVIDENDI = cid
        if cid := self._cat_by_name.get(_CAT_RITENUTA_NAME.lower()):
            _CAT_RITENUTA = cid

    def close(self):
        self._conn.close()

    def cat_id(self, name: str) -> int | None:
        """Ritorna ID categoria da nome (case-insensitive)."""
        return self._cat_by_name.get((name or "").lower())

    def account_id(self, name: str) -> int | None:
        """Ritorna ID conto da nome (case-insensitive)."""
        return self._account_by_name.get((name or "").lower())

    def inserisci_transazione(self,
                               account_id: int,
                               txn_date: "date",
                               deposit: float = 0.0,
                               withdrawal: float = 0.0,
                               memo: str = "",
                               category_id: int | None = None,
                               payee_id: int | None = None,
                               status: int = 0,
                               splits: list[dict] | None = None) -> int:
        """
        Inserisce una transazione con eventuali splits.
        splits: lista di dict con chiavi deposit, withdrawal, category_id, memo
        Ritorna l'ID della transazione inserita.
        """
        cur = self._conn.cursor()
        txn_id = _next_id(cur, "Transactions")

        if self._has_sync_edit():
            cur.execute("""
                INSERT INTO Transactions
                  (ID, AccountID, TransactionDate, Withdrawal, Deposit,
                   Memo, CategoryID, PayeeID, Status,
                   SyncAdd, SyncEdit)
                VALUES (?,?,?,?,?,?,?,?,?,1,0)
            """, (
                txn_id, account_id,
                txn_date.isoformat(),
                withdrawal if withdrawal else None,
                deposit    if deposit    else None,
                memo or None,
                category_id,
                payee_id,
                status
            ))
        else:
            cur.execute("""
                INSERT INTO Transactions
                  (ID, AccountID, TransactionDate, Withdrawal, Deposit,
                   Memo, CategoryID, PayeeID, Status)
                VALUES (?,?,?,?,?,?,?,?,?)
            """, (
                txn_id, account_id,
                txn_date.isoformat(),
                withdrawal if withdrawal else None,
                deposit    if deposit    else None,
                memo or None,
                category_id,
                payee_id,
                status
            ))

        if splits:
            # Verifica SyncAdd/SyncEdit anche per la tabella Splits
            cur.execute("PRAGMA table_info(Splits)")
            splits_cols = {row[1] for row in cur.fetchall()}
            splits_has_sync = "SyncEdit" in splits_cols
            for sp in splits:
                sp_id = _next_id(cur, "Splits")
                if splits_has_sync:
                    cur.execute("""
                        INSERT INTO Splits
                          (ID, TransactionID, CategoryID, Withdrawal, Deposit, Memo,
                           SyncAdd, SyncEdit)
                        VALUES (?,?,?,?,?,?,1,0)
                    """, (
                        sp_id, txn_id,
                        sp.get("category_id"),
                        sp.get("withdrawal") or None,
                        sp.get("deposit")    or None,
                        sp.get("memo")       or None
                    ))
                else:
                    cur.execute("""
                        INSERT INTO Splits
                          (ID, TransactionID, CategoryID, Withdrawal, Deposit, Memo)
                        VALUES (?,?,?,?,?,?)
                    """, (
                        sp_id, txn_id,
                        sp.get("category_id"),
                        sp.get("withdrawal") or None,
                        sp.get("deposit")    or None,
                        sp.get("memo")       or None
                    ))

        self._conn.commit()
        return txn_id

    def _has_sync_edit(self) -> bool:
        """Verifica se la colonna SyncEdit esiste nella tabella Transactions."""
        if not hasattr(self, "_sync_edit_checked"):
            cur = self._conn.cursor()
            cur.execute("PRAGMA table_info(Transactions)")
            cols = {row[1] for row in cur.fetchall()}
            self._sync_edit_ok = "SyncEdit" in cols
            self._sync_edit_checked = True
        return self._sync_edit_ok

    def correggi_importo(self, txn_id: int,
                          deposit: float | None = None,
                          withdrawal: float | None = None) -> bool:
        """Aggiorna importo di una transazione esistente."""
        cur = self._conn.cursor()
        sync = ", SyncEdit=1" if self._has_sync_edit() else ""
        if deposit is not None:
            cur.execute(f"UPDATE Transactions SET Deposit=?{sync} WHERE ID=?",
                        (deposit, txn_id))
        if withdrawal is not None:
            cur.execute(f"UPDATE Transactions SET Withdrawal=?{sync} WHERE ID=?",
                        (withdrawal, txn_id))
        self._conn.commit()
        return cur.rowcount > 0

    def correggi_data(self, txn_id: int, nuova_data: "date") -> bool:
        """Aggiorna la data di una transazione esistente."""
        cur = self._conn.cursor()
        sync = ", SyncEdit=1" if self._has_sync_edit() else ""
        cur.execute(
            f"UPDATE Transactions SET TransactionDate=?{sync} WHERE ID=?",
            (nuova_data.isoformat(), txn_id)
        )
        self._conn.commit()
        return cur.rowcount > 0