# Moneyspire Reconciler — Documentazione Completa

*Versione corrente: v2.1.3 — Aggiornato: 2026-05-22*

---

## Indice

1. [Scopo e architettura](#1-scopo-e-architettura)
2. [File del progetto](#2-file-del-progetto)
3. [Configurazione multi-profilo](#3-configurazione-multi-profilo)
4. [ms_engine.py — Modulo motore](#4-ms_enginepy--modulo-motore)
5. [ms_reconciler.py — GUI](#5-ms_reconcilerpy--gui)
6. [Fase 1: Riconciliazione Money ↔ Banca](#6-fase-1-riconciliazione-money--banca)
7. [Fase 2: Integrazione file Excel elaborati](#7-fase-2-integrazione-file-excel-elaborati)
8. [Regole di categorizzazione](#8-regole-di-categorizzazione)
9. [Struttura file Excel elaborati](#9-struttura-file-excel-elaborati)
10. [Flusso scrittura sicura (backup)](#10-flusso-scrittura-sicura-backup)
11. [Log di audit](#11-log-di-audit)
12. [Problemi noti e soluzioni](#12-problemi-noti-e-soluzioni)
13. [Aggiornamenti futuri previsti](#13-aggiornamenti-futuri-previsti)

---

## 1. Scopo e architettura

**Moneyspire Reconciler** confronta i movimenti bancari (file esportati da Fineco, BPER/ex Popolare di Sondrio, Nexi) con le transazioni già inserite in Moneyspire (database `.ffd`, formato SQLite). Permette di:

- Identificare transazioni bancarie non ancora caricate in Money (❌ Mancanti)
- Rilevare transazioni in Money senza corrispondenza in banca (📝 Solo Money)
- Inserire automaticamente i mancanti nel database Moneyspire (Fase 1)
- Aggiornare data e/o importo di transazioni con match fuzzy (🔶 Trovata ±gg)
- Mantenere aggiornati i file Excel elaborati annuali (Fase 2)

**Architettura modulare:**

```
ms_engine.py           — wrapper: ri-esporta tutti i simboli dai sottomoduli
ms_parsers.py          — parser Fineco (CC, carte, cartella multi-file) e Unicredit
ms_parsers_silvia.py   — parser BPER (.xls) e Nexi (.xlsx) — profilo SC
ms_matching.py         — ReconcileEngine, RulesEngine, valida_file_banca
ms_db.py               — MoneyspireDB, MoneyWriter, backup/verifica
ms_excel.py            — ExcelIntegrator, leggi_excel_elaborato_*
ms_constants.py        — costanti MATCH_*, STATO_LABELS, DEFAULT_CONFIG
ms_reconciler.py       — GUI tkinter: combobox profilo + 4 tab
```

**Percorso canonico:**
```
~/Library/CloudStorage/Dropbox/Documenti_IRC/Python/stable/Riconciliazione Money.CLD/
```

---

## 2. File del progetto

| File | Ruolo |
|------|-------|
| `ms_reconciler.py` | GUI tkinter con selezione profilo via combobox |
| `ms_engine.py` | Wrapper: ri-esporta tutti i simboli pubblici |
| `ms_parsers.py` | Parser Fineco CC, carte singolo/cartella, USD, Unicredit |
| `ms_parsers_silvia.py` | Parser BPER (.xls) e Nexi (.xlsx) — profilo SC |
| `ms_matching.py` | Motore matching, regole, validazione file banca |
| `ms_db.py` | Accesso database .ffd e scrittura |
| `ms_excel.py` | Integrazione file Excel elaborati |
| `ms_constants.py` | Costanti, etichette, DEFAULT_CONFIG |
| `path_widgets.py` | Widget PathEntry condiviso (`Python/shared/`) |

**File di stato:**
```
_Config/Riconciliazione Moneyspire/ultimo_profilo.json   — ultimo ambiente usato
```

**Lancio:**
```bash
python3 ms_reconciler.py
```

---

## 3. Configurazione multi-profilo

### 3.1 Struttura cartelle

Due profili indipendenti: **IRC** (Ignazio Rusconi Clerici) e **SC** (Silvia Corinaldi).

```
~/Library/CloudStorage/Dropbox/Documenti_IRC/Python/
├── _Config/
│   └── Riconciliazione Moneyspire/
│       ├── IRC_build.json
│       ├── IRC_config.json
│       ├── IRC_rules.json
│       ├── SC_build.json
│       ├── SC_config.json
│       ├── SC_rules.json
│       └── ultimo_profilo.json       ← profilo usato all'ultimo avvio
│
└── stable/
    └── Riconciliazione Money.CLD/
        ├── ms_reconciler.py
        ├── ms_engine.py
        ├── ms_parsers.py
        ├── ms_parsers_silvia.py
        ├── ms_matching.py
        ├── ms_db.py
        ├── ms_excel.py
        └── ms_constants.py
```

File `.ffd`:
```
Moneyspire.26/Data/Ignazio.ffd   ← profilo IRC
Moneyspire.26/Data/Silvia.ffd    ← profilo SC
```

### 3.2 Convenzione nomi file

| Profilo | Config | Rules | Build | Prefisso output |
|---------|--------|-------|-------|-----------------|
| IRC (Ignazio) | `IRC_config.json` | `IRC_rules.json` | `IRC_build.json` | `IRC_` |
| SC (Silvia) | `SC_config.json` | `SC_rules.json` | `SC_build.json` | `SC_` |

I file CSV di export includono prefisso profilo e nome conto:
```
IRC_MC_Fineco_riconciliazione_completa.csv
SC_Visa_Oro_Popso_(Nexi)_transazioni_mancanti.csv
```

### 3.3 Struttura IRC_config.json / SC_config.json

```json
{
  "_profilo": "IRC",
  "_nome": "Ignazio Rusconi Clerici",
  "date_tolerance_days": 5,
  "amount_tolerance": 0.01,
  "tolleranze_fuzzy": {
    "importo_max_eur": 1.0,
    "commissione_max_eur": 0.50,
    "commissione_giorni": 2
  },
  "giorni_fine_mese_pending": 4,
  "soglia_aggiustamento_saldo": 5.0,
  "contropartita_default": "Da classificare",
  "prefisso_output": "IRC_",
  "conti": { ... },
  "conti_contanti": { ... },
  "last_paths": { ... }
}
```

### 3.4 Conti profilo IRC (Ignazio)

| Nome | ffd_account_id | Tipo | tipo_file_analisi | Note |
|------|---------------|------|-------------------|------|
| Fineco | 536 | conto_corrente | originale_fineco | Conto principale |
| Fineco Lombard | 577 | conto_corrente | originale_fineco | |
| Fineco USD | 554 | conto_corrente | originale_fineco | Conto dollari |
| Unicredit CCM | 548 | conto_corrente | originale_unicredit | File .xls |
| MC Fineco | 531 | carta_credito | estratto_cc | Carta 5260 |
| Visa Fineco | 532 | carta_credito | estratto_cc | Carte 6421, 6553 |
| Fidaty Oro | 557 | — | — | Manuale, escluso |
| Fineco GBP / CHF | 573/597 | conto_corrente | originale_fineco | Inattivi, esclusi |

### 3.5 Conti profilo SC (Silvia)

| Nome | ffd_account_id | Tipo | tipo_file_analisi | Note |
|------|---------------|------|-------------------|------|
| Popolare Sondrio (BPER) | 433 | conto_corrente | originale_bper | File .xls |
| Fineco personale | 539 | conto_corrente | originale_fineco | Stesso formato IRC |
| Visa Oro Popso (Nexi) | 476 | carta_credito | originale_nexi_xlsx | File .xlsx da Nexi Pay |
| Fineco titoli | 542 | conto_corrente | originale_fineco | Sporadico, escluso |

### 3.6 Parser BPER (profilo SC)

File `.xls` esportato dal portale BPER (Excel 97-2003, letto con `xlrd`).

Struttura:
- Righe 0–15: intestazione (IBAN, saldi, titolare, filtri)
- Riga 16: header colonne (`Data operazione | Data valuta | Descrizione | Entrate | Uscite | Categoria | Stato`)
- Riga 17+: movimenti; ultima riga = totali (skippata)

Date in formato italiano (`"05 maggio 2026"`). Stato: `"Contabilizzato"` o `"Da contabilizzare"` (incluso per default, escludibile con `includi_non_contabilizzati: false` nel config).

Il parser restituisce i campi standard del motore di matching (`date`, `amount`, `raw_text`, `deposit`, `withdrawal`) più i campi specifici BPER (`data_valuta`, `categoria_banca`, `stato`, `fonte`).

Il confronto saldi post-riconciliazione **non è disponibile** per BPER (solo per Fineco).

### 3.7 Parser Nexi xlsx (profilo SC)

File `.xlsx` esportato da **Nexi Pay** (`Estratto conto → Export`).

Struttura:
- Righe 0–8: intestazione (numero carta, periodo, tipo)
- Riga 9: header (`Mese | Data | Riferimento | Categorie | Descrizione | Stato | ... | Importo (€)`)
- Riga 10+: movimenti

Tutti i movimenti sono uscite (importo positivo nel file → negativo nel dict). Il file può coprire più mesi; il filtro mese/anno è applicato automaticamente. Supporta modalità **cartella multi-file** (vedi sezione 5.4).

**Nota Nexi ↔ BPER:** l'addebito mensile Nexi sul conto BPER appare come `"ADDEBITO SDD Nexi Payments SpA"` il giorno 15–16 del mese successivo. La regola in `SC_rules.json` lo mappa su `"Trasferimenti:Carta Nexi"`.

### 3.8 last_paths — chiavi

| Chiave | Descrizione |
|--------|-------------|
| `ffd` | File Moneyspire `.ffd` |
| `xls_cc` | Ultimo file banca (Fase 1) |
| `xls_cc_cartella` | Ultima cartella multi-file |
| `xls_cc_<NomeConto>` | Ultimo file per conto specifico |
| `conto` | Conto selezionato in Fase 1 |
| `fase2_orig` | File banca originale (Fase 2) |
| `xlsx_elaborato` | File elaborato annuale (Fase 2) |
| `fase2_conto` | Conto selezionato in Fase 2 |

I `last_paths` sono nel rispettivo config (IRC/SC). L'ultimo profilo usato è in `ultimo_profilo.json`.

---

## 4. ms_engine.py — Modulo motore

Wrapper che ri-esporta tutti i simboli dai sottomoduli:

```python
from ms_parsers import (
    parse_fineco_conto_originale, parse_fineco_cc,
    parse_fineco_cc_cartella,    # multi-file da cartella
    parse_unicredit_ccm, ...
)
from ms_parsers_silvia import (
    parse_bper, leggi_intestazione_bper,
    parse_nexi_xlsx, leggi_intestazione_nexi,
)
```

### 4.1 Costanti match

| Costante | Label UI | Colore |
|----------|----------|--------|
| `MATCH_EXACT` | ✅ Trovata | verde |
| `MATCH_FUZZY` | 🔶 Trovata (±gg) | giallo |
| `MATCH_SPLIT` | 🔀 Split | azzurro |
| `MATCH_MERGE` | 🔗 Merge/Composta | viola |
| `MATCH_NONE` | ❌ Mancante | rosso |
| `MATCH_SKIP` | 📝 Solo Money | grigio |
| `MATCH_PENDING` | ⏳ In attesa | arancio |

### 4.2 MoneyspireDB

```python
db = MoneyspireDB(ffd_path)
db.get_transactions(account_id, date_from, date_to)
db.get_categories()        # dict[int, str]
db.get_payees()            # dict[int, str]
db.get_payee_categories()  # dict[str, int]
```

### 4.3 Parser file banca

**`parse_fineco_conto_originale(xlsx, mese, anno)`** — IRC: file CC Fineco, header riga 13, filtra `Stato == "Contabilizzato"`.

**`parse_fineco_cc(xlsx, numero_carta, mese, anno)`** — IRC: estratto carte MC/Visa, filtra per data di registrazione.

**`parse_fineco_cc_cartella(cartella, numero_carta, mese, anno)`** — IRC: legge tutti i `.xlsx` nella cartella, unifica e deduplicca per `(data, importo, descrizione, carta)`.

**`parse_unicredit_ccm(xls, mese, anno)`** — IRC: file `.xls` Unicredit con `xlrd`.

**`parse_bper(xls, mese, anno)`** — SC: file `.xls` BPER con `xlrd`. Vedi sezione 3.6.

**`parse_nexi_xlsx(xlsx, mese, anno)`** — SC: file `.xlsx` Nexi Pay. Vedi sezione 3.7.

### 4.4 ReconcileEngine

Algoritmo di matching a 7 livelli:

| Livello | Criterio | Confidenza |
|---------|----------|-----------|
| 1 | Importo esatto + data identica | 1.00 |
| 1b | Importo quasi identico (Δ ≤ tolleranza) + data identica | 0.92 |
| 1c | Importo quasi identico (prob. commissione) + data ±2 gg | 0.88 |
| 2 | Importo esatto + data ±N giorni | 0.80–0.95 |
| 3 | Banca = netto di split in Money | 0.88 |
| 4 | Banca = uno split di transazione composta | 0.72 |
| 5 | Categoria con tolleranza % | 0.60–0.85 |
| 6 | Giroconto/cambio valuta cross-conto | 0.85 |
| 7 | Nessun match → MATCH_NONE | 0.00 |

**Post-processing `_correggi_fuzzy_invertiti`:** per coppie di match fuzzy con importi quasi identici (Δ ≤ `importo_max_eur`, default 1€), confronta la similarità testo-banca vs categoria+memo+payee Money per entrambe le combinazioni. Se la combinazione inversa è ≥ 10% migliore, scambia gli abbinamenti e aggiunge "[corretto inversione]" alla nota.

### 4.5 RulesEngine

```json
{
  "pattern": "FARMACIA",
  "regex": false,
  "category": "Spese mediche:Farmacia",
  "payee": "",
  "source": "manual",
  "hits": 8
}
```

### 4.6 valida_file_banca

| tipo validazione | Metodo | Libreria |
|-----------------|--------|---------|
| `foglio_excel` | Verifica presenza foglio | openpyxl |
| `numero_carta` | Verifica ultime 4 cifre carta | openpyxl |
| `intestazione_bper` | Parole chiave BPER prime 20 righe | xlrd |
| `intestazione_nexi` | Parole chiave Nexi (xlsx) | openpyxl |
| `intestazione_conto` | Testo Unicredit prime 5 righe | xlrd/pandas |

Il confronto saldi post-riconciliazione è attivo **solo** per `originale_fineco` e `originale_fineco_usd`.

### 4.7 MoneyWriter

```python
work_path, backup_path = prepara_db_scrittura(ffd_path, profilo)  # profilo "IRC" o "SC"
writer = MoneyWriter(work_path)
txn_id = writer.inserisci_transazione(...)
writer.correggi_data(txn_id, nuova_data)
writer.correggi_importo(txn_id, deposit=val)
writer.close()
ok, msg = verifica_scrittura(work_path, account_id, date_from, date_to,
                              n_attese, expected_txns=lista_dict)
if ok:
    finalizza_db(ffd_originale, work_path)
```

---

## 5. ms_reconciler.py — GUI v2.1

### 5.1 Selezione ambiente

All'avvio si apre direttamente con l'ultimo ambiente usato (da `ultimo_profilo.json`, default IRC).

**Toolbar in cima al tab Riconcilia:**
```
Ambiente: [IRC — Ignazio  ▼]   Prefisso output: IRC_  |  47 regole caricate
```

Al cambio ambiente dalla combobox vengono automaticamente: caricati config e rules, aggiornati i combobox conti, puliti tabella risultati e log, aggiornato il titolo, salvato il nuovo profilo in `ultimo_profilo.json`.

### 5.2 Tab 1 — Riconcilia

**Selezione file banca:** riga "Movimenti originali banca" con due pulsanti:
- **📂** → file singolo (si riposiziona sulla cartella dell'ultimo file per quel conto)
- **📁** → cartella multi-file

Il dialogo file accetta `.xlsx`, `.xls` e `.pdf`. Al cambio conto il PathVar viene pulito se l'estensione del file corrente è incompatibile col nuovo tipo (mappa estensioni per tipo: BPER → `.xls`, Nexi → `.xlsx`, Fineco → `.xlsx`, Unicredit → `.xls`).

**Opzione "Tutto l'anno"** nel dialogo periodo: disponibile per tutti i conti correnti e per le carte quando è selezionata una cartella multi-file.

**Export CSV:** si apre su `Documents/download/` con nome `{prefisso}{NomeConto}_riconciliazione_completa.csv`.

### 5.3 Barra inferiore

```
[status operativa]  |  📋 Copia  💾 Salva  🗑 Pulisci  |  ✖ Esci
```

### 5.4 Modalità cartella multi-file

Per le **carte di credito Fineco** (MC, Visa): selezionando una cartella con 📁, `parse_fineco_cc_cartella` legge tutti i `.xlsx` presenti, li unifica e deduplicca. Risolve il problema delle transazioni di fine mese che compaiono in estratti di mesi diversi.

Per la **Visa Nexi** (SC): la cartella deve contenere file `.xlsx` esportati da Nexi Pay. I file vengono letti, unificati e deduplicati con la stessa logica.

### 5.5 DialogRevisione

Mostra le transazioni da inserire con categoria suggerita. Doppio click → `DialogModificaCategoria` con filtro live categorie e payee → categoria automatica.

### 5.6 DialogAggiornaFuzzy

Aggiorna selettivamente data e/o importo dei match fuzzy. Transazioni split → importo mai modificabile (🔒). Segno invertito → importo mai modificabile (⚠️).

---

## 6. Fase 1: Riconciliazione Money ↔ Banca

### Flusso completo

```
File banca (Fineco / BPER / Nexi / Unicredit) — singolo o cartella
        ↓
  parse_*() / parse_*_cartella()       → bank_txns
        ↓
  raggruppa_cedole_ritenute()
  raggruppa_transazioni_correlate()
        ↓
  ReconcileEngine.reconcile()          → results
        ↓
  _correggi_fuzzy_invertiti()          → corregge abbinamenti capovolti
        ↓
  Post-processing cedole/dividendi
  marca_solo_money_in_attesa() (carte)
        ↓
  Visualizzazione Treeview con filtri
        ↓
  costruisci_transazioni_da_risultati() → da_inserire
        ↓
  DialogRevisione → conferma
        ↓
  MoneyWriter → copia .ffd → verifica forte → finalizza_db()
        ↓
  write_audit_log()
```

### Confronto saldi

Disponibile **solo** per `originale_fineco` / `originale_fineco_usd`, solo per analisi mensile (non annuale). Se `|Δ variazione| > 5€` → avviso nel log.

### Gestione "In attesa"

Transazioni degli ultimi `giorni_fine_mese_pending` (default 4) giorni del mese → `⏳ In attesa`.

---

## 7. Fase 2: Integrazione file Excel elaborati

Disponibile per profilo IRC. Mantiene aggiornati i file elaborati annuali Fineco.

| File elaborato | Fogli |
|---------------|-------|
| `2026.xlsx` | Movimenti, Lombard, USD — saldo a catena |
| `2026 mc.xlsx` | Fogli mensili MC |
| `2026 visa.xlsx` | Fogli mensili Visa (con Intestatario) |

> **Attenzione:** il file Visa usa `"gennazio"` (typo) per gennaio. Da correggere nel file Excel.

---

## 8. Regole di categorizzazione

### Struttura

```json
{
  "pattern": "FARMACIA",
  "regex": false,
  "category": "Spese mediche:Farmacia",
  "payee": "",
  "source": "manual",
  "hits": 8
}
```

### Categorie speciali

- `"Ignora"` → non inserita in Money
- `"Trasferimento:carta"` → addebito estratto CC
- `"Trasferimento:conto"` → giroconto
- `contropartita_default` → transazione senza categoria specifica

### Regole pre-configurate SC (Silvia)

| Pattern | Categoria | Note |
|---------|-----------|------|
| `PENSIONE INPS` | Retribuzione:Pensione | |
| `ADDEBITO SDD.*WIND TRE` | Casa:Telefono | |
| `ADDEBITO SDD.*FASTWEB` | Casa:Internet | |
| `ADDEBITO SDD.*IREN` | Casa:Utenze | Gas/luce/acqua |
| `PAGAMENTO PAGOPA` | Imposte:PagoPA | |
| `QUOTA ANNUA` | Spese bancarie:Quota annua carta | Quota Nexi 103,29€/anno |
| `ADDEBITO SDD.*NEXI` | Trasferimenti:Carta Nexi | Addebito mensile su BPER |
| `COMMISSIONI BONIFICI` | Spese bancarie:Commissioni | Commissioni BPER |

---

## 9. Struttura file Excel elaborati

Riconoscimento automatico colonna saldo:
```python
has_moneymap = "moneymap" in [str(v or "").lower() for v in hdr]
col_saldo = 7 if has_moneymap else 6   # G per Movimenti, F per Lombard/USD
```

---

## 10. Flusso scrittura sicura (backup)

```
1. Verifica Moneyspire non in esecuzione (solo .ffd)
2. Backup timestampato in sottocartella _Backup_<PROFILO>/:
     <dir_originale>/_Backup_<PROFILO>/NomeFile_backup_YYYYMMDD_HHMMSS.ext
   La sottocartella viene creata se non esiste (es. _Backup_IRC, _Backup_SC).
3. Copia di lavoro nella cartella originale:
     <dir_originale>/NomeFile_work_YYYYMMDD_HHMMSS.ext
4. Scrittura sulla copia (mai sull'originale)
5. Verifica forte (data + importo + categoria + payee + memo parziale)
6. Conferma utente
7a. Conferma → Path.replace (sostituzione atomica)
7b. Annulla → elimina copia di lavoro
```

**Note:**
- Il parametro `profilo` è passato dalla GUI come variabile globale `_PROFILO` (`"IRC"` o `"SC"`).
- Il backup `.ffd` viene quindi creato in `<cartella_Moneyspire>/_Backup_<PROFILO>/`, separato dalla cartella `Data` per evitare rumore tra i file `.ffd` operativi.
- Il backup `.xlsx` viene creato in `<cartella_Excel>/_Backup_<PROFILO>/`, accanto agli Excel originali (che vivono in una cartella diversa dai `.ffd`).
- Non è prevista retention automatica: i backup si accumulano e vanno eventualmente puliti a mano.

---

## 11. Log di audit

```
_Config/Riconciliazione Moneyspire/Logs/ms_reconciler_YYYY.log
```

Append-only, un file per anno.

```
════════════════════════════════════════════════════════════════════════
SESSIONE  2026-05-06 11:00:00   Moneyspire Reconciler  v2.1 — Silvia (SC)
File .ffd : Silvia.ffd
Operazione: INSERIMENTO   Conto: Popolare Sondrio (BPER) (ID 433)
────────────────────────────────────────────────────────────────────────
  INS   2026-04-20  +500,00 €  Cat: Trasferimenti  ID:7050
  SKIP  2026-04-17   -0,40 €  Cat: Ignora
────────────────────────────────────────────────────────────────────────
Riepilogo: 1 inserite  1 ignorate  0 errori  — Verifica: OK forte (1/1)
════════════════════════════════════════════════════════════════════════
```

| Tipo | Significato |
|------|-------------|
| `INS` | Transazione inserita |
| `SKIP` | Ignorata |
| `ERR` | Errore di scrittura |
| `UPD_DATA` | Data aggiornata |
| `UPD_IMP` | Importo aggiornato |
| `UPD_ENTRAMBI` | Data e importo aggiornati |
| `WARN` | Aggiornamento bloccato |

---

## 12. Problemi noti e soluzioni

### `error messaging the mach port for IMKCFRunLoopWakeUpReliable`
Bug macOS/IMKit con tkinter. Ignorare, oppure `python3 ms_reconciler.py 2>/dev/null`.

### File BPER o Unicredit: errore openpyxl "does not support file format"
Il codice tenta di aprire un `.xls` con openpyxl. Verificare di avere la versione corrente di `ms_matching.py` e `ms_reconciler.py`, che usa `xlrd` per i `.xls` e limita il confronto saldi ai soli file Fineco.

### Opzione "Tutto l'anno" non appare per le carte
Con file singolo, l'opzione anno è disponibile solo per i conti correnti. Selezionare una cartella con 📁 abilita l'opzione anno anche per le carte.

### PathVar mostra il file del conto precedente al cambio conto
Al cambio conto, il PathVar viene pulito automaticamente se l'estensione è incompatibile col tipo del nuovo conto. Il path specifico per ogni conto è memorizzato in `last_paths["xls_cc_<NomeConto>"]`.

### File BPER: movimenti "Da contabilizzare" inclusi
I bonifici del giorno stesso appaiono come "Da contabilizzare". Inclusi per default. Per escluderli: `includi_non_contabilizzati: false` nel SC_config.json.

### Carta Nexi: addebito BPER non trovato come trasferimento
L'addebito mensile Nexi sul conto BPER va riconciliato come trasferimento interno. La regola `ADDEBITO SDD.*NEXI` in `SC_rules.json` lo mappa su `"Trasferimenti:Carta Nexi"`.

### Match fuzzy invertito tra importi quasi identici
Il post-processing `_correggi_fuzzy_invertiti` rileva e corregge automaticamente gli abbinamenti capovolti confrontando la similarità testo-banca vs categoria+memo+payee Money. La correzione appare nella colonna Nota come "[corretto inversione]".

---

## 13. Aggiornamenti futuri previsti

- [ ] **Test Visa Nexi xlsx** — verificare con il file .xlsx reale di Silvia quando disponibile; se ok, rimuovere codice PDF residuo
- [ ] **Confronto saldo fine mese** — transazione di aggiustamento automatica se Δ > soglia
- [ ] **Aggiornamento Moneymap** in blocco nel file elaborato
- [ ] **PyInstaller bundle `.app`** — aggiornare AppBuilder per profili IRC/SC e `ultimo_profilo.json`
- [ ] **Visa gennaio** — correggere typo `"gennazio"` nel file xlsx elaborato

### Architettura multi-modulo (implementata)

| Modulo | Contenuto |
|--------|-----------|
| `ms_db.py` | `MoneyspireDB`, `MoneyWriter`, backup/verifica |
| `ms_parsers.py` | `parse_fineco_*`, `parse_fineco_cc_cartella`, `parse_unicredit_ccm` |
| `ms_parsers_silvia.py` | `parse_bper`, `parse_nexi_xlsx` |
| `ms_matching.py` | `ReconcileEngine`, `RulesEngine`, `valida_file_banca`, `_correggi_fuzzy_invertiti` |
| `ms_excel.py` | `ExcelIntegrator`, `export_*` |
| `ms_constants.py` | Costanti, `DEFAULT_CONFIG`, utility |
| `ms_engine.py` | Wrapper che ri-esporta tutto |

---

*Documentazione generata il 2026-05-01 da Claude Sonnet 4.6*
*Aggiornata il 2026-05-03 — v2.0: separazione moduli*
*Aggiornata il 2026-05-05 — v2.1: multi-profilo IRC/SC, parser BPER e Nexi xlsx*
*Aggiornata il 2026-05-06 — v2.1 rev.2: combobox profilo, cartella multi-file, fix fuzzy invertito, export CSV con nome conto, pulizia al cambio ambiente, ultimo_profilo.json, compatibilità .xls BPER/Unicredit*
*Aggiornata il 2026-05-22 — v2.1.1: fix NameError `_CAT_DIVIDENDI` in `costruisci_transazioni_da_risultati` (causa del dialog "Revisione" vuoto su mese maggio); backup in sottocartella `_Backup_<PROFILO>/` accanto al file originale, sia per `.ffd` sia per `.xlsx`*
*Aggiornata il 2026-05-22 — v2.1.2: guard segno coerente in tutti i livelli di match fuzzy (1a esatto, 1b stessa-data-importo-quasi-identico, 1c commissione, 2 fuzzy importo+data±N, 3 split). Una transazione banca non viene più abbinata a una Money con importo di segno opposto (es. +105 € entrata vs -105 € uscita). Il livello 6 trasferimenti/cambio valuta resta esente: lì il segno opposto è proprio il pattern cercato.*
*Aggiornata il 2026-05-22 — v2.1.3: il dialog "Seleziona periodo" propone come default il mese precedente solo nei primi 10 giorni del mese, dopodiché propone il mese corrente.*
