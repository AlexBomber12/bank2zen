# bank2zen.py  – Fineco → ZenMoney   (2025-06-30)

import pandas as pd, json, re, sys, os
from pathlib import Path

# ── ваши счета ────────────────────────────────────────────────────────────────
ACC_DEBIT   = "Fineco Debit"
ACC_CREDIT  = "Fineco Credit"
ACC_CASH    = "Наличные Евро"

# ── файлы самообучения ────────────────────────────────────────────────────────
CATS_FILE = "categories.json"
ACC_FILE  = "accounts_to.json"

# ── паттерны ──────────────────────────────────────────────────────────────────
# 1. точный шаблон кредитки – смотрим ТОЛЬКО колонку «Descrizione»
rx_credit  = re.compile(r"MONOFUNZIONE\s+CONTACTLESS\s+CHIP\s+5100.*3142",
                        re.I)

# 2. снятие наличных (осталось без изменений)
rx_cashout = re.compile(r"Prelievo\s+Bancomat\s+Unicredit", re.I)

# ── helpers ───────────────────────────────────────────────────────────────────
def normalize(t: str) -> str:
    t = re.sub(r"\d+", "", t)
    t = re.sub(r"\s{2,}", " ", t)
    return t.strip().lower()

def _norm_cell(value) -> str:
    value = "" if value is None else str(value)
    return re.sub(r"[^a-z0-9]", "", value.strip().lower())

_ALIASES = {
    "Data_Valuta": {
        "datavaluta", "dataval", "datavalut", "data"
    },
    "Entrate": {
        "entrate", "ebt", "income", "crediti", "accredito"
    },
    "Uscite": {
        "uscite", "usc", "expense", "addebiti", "debito", "spese"
    },
    "Descrizione": {
        "descrizione", "descr", "discr", "description", "causale"
    },
    "Descrizione_Completa": {
        "descrizionecompleta", "descrcompleta", "discrcompleta",
        "descrizionecomplet", "fulldescription", "dettaglio"
    },
}

def _read_movements_xlsx(path: str) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, engine="openpyxl")
    hdr_idx = None
    target_keys = _ALIASES["Data_Valuta"]

    for idx in range(len(raw)):
        row_vals = {
            _norm_cell(v)
            for v in raw.iloc[idx].tolist()
            if pd.notna(v) and str(v).strip()
        }
        if row_vals & target_keys:
            hdr_idx = idx
            break

    if hdr_idx is None:
        raise ValueError("Не удалось найти заголовок 'Data_Valuta' в XLSX. Проверьте шаблон выгрузки.")

    if hdr_idx:
        print(f"Пропущено строк до шапки: {hdr_idx}")

    df = pd.read_excel(path, header=hdr_idx, engine="openpyxl")

    rename_map = {}
    for col in df.columns:
        key = _norm_cell(col)
        for canon, keys in _ALIASES.items():
            if key in keys:
                rename_map[col] = canon
                break

    df = df.rename(columns=rename_map)

    if "Data_Valuta" not in df.columns:
        raise ValueError("Колонка даты Data_Valuta не найдена.")

    if not ({"Entrate", "Uscite"} & set(df.columns)):
        raise ValueError("В XLSX нет колонок сумм (Entrate/Uscite).")

    if not ({"Descrizione", "Descrizione_Completa"} & set(df.columns)):
        raise ValueError("В XLSX нет колонок описания (Descrizione/Descrizione_Completa).")

    if "Descrizione_Completa" not in df.columns and "Descrizione" in df.columns:
        df["Descrizione_Completa"] = df["Descrizione"]

    keep = ["Data_Valuta", "Entrate", "Uscite", "Descrizione", "Descrizione_Completa"]
    present = [c for c in keep if c in df.columns]
    df = df[present].copy()

    df["Data_Valuta"] = pd.to_datetime(df["Data_Valuta"], dayfirst=True, errors="coerce")
    missing_dates = df["Data_Valuta"].isna().sum()
    if missing_dates:
        print(f"Предупреждение: отброшено строк без даты: {missing_dates}")
        df = df[df["Data_Valuta"].notna()].copy()

    for col in ("Entrate", "Uscite"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).abs()

    return df

def _load(p):
    if Path(p).exists():
        try:
            return json.load(open(p, encoding="utf-8"))
        except json.JSONDecodeError:
            pass
    return {}

def _lookup(nrm, table):
    for key, patterns in table.items():
        if any(p in nrm for p in patterns):
            return key
    return ""

# ── конвертер ────────────────────────────────────────────────────────────────
def convert(xlsx):
    df = _read_movements_xlsx(xlsx)

    df = df.copy()
    df["Data"] = df["Data_Valuta"]
    if "Descrizione" not in df.columns:
        df["Descrizione"] = df["Descrizione_Completa"]
    if "Entrate" not in df.columns:
        df["Entrate"] = 0.0
    if "Uscite" not in df.columns:
        df["Uscite"] = 0.0

    df["Norm"] = df["Descrizione_Completa"].apply(lambda s: normalize(str(s)))

    cat_map = _load(CATS_FILE)
    acc_map = _load(ACC_FILE)

    df["Category"]  = df["Norm"].apply(lambda n: _lookup(n, cat_map))
    df["AccountTo"] = df["Norm"].apply(lambda n: _lookup(n, acc_map))

    # ── пост-обработка строк ─────────────────────────────
    def split(row):
        desc_short = str(row.get("Descrizione", ""))  # ← короткое описание
        full       = str(row["Descrizione_Completa"]).lower()
        inc  = row["Entrate"] if pd.notna(row["Entrate"]) else ""
        out  = abs(row["Uscite"]) if pd.notna(row["Uscite"]) else ""
        acc_from, acc_to = ACC_DEBIT, row["AccountTo"]

        if rx_credit.search(desc_short):              # кредитка
            acc_from = ACC_CREDIT
        elif rx_cashout.search(full):                 # снятие кеша
            acc_to   = ACC_CASH
            inc      = out                            # приход на Cash

        return pd.Series([acc_from, acc_to, inc, out])

    df[["Account", "AccountTo", "Income", "Expense"]] = df.apply(split, axis=1)

    # ── незнакомые категории → new_desc.xlsx ────────────
    unknown = df[df["Category"] == ""][["Data", "Entrate", "Uscite", "Descrizione_Completa"]]
    if len(unknown):
        unknown[["Data", "Entrate", "Uscite", "Descrizione_Completa"]].to_excel(
            "new_desc.xlsx", index=False
        )
        return "need_class"

    # ── итоговый CSV ────────────────────────────────────
    cols = ["Data","Category","Descrizione_Completa",
            "Account","AccountTo","Income","Expense"]
    df.to_csv("out_zenmoney.csv", columns=cols,
              index=False, encoding="utf-8-sig", sep=";")
    return "ok"

# ── CLI ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else \
          next((f for f in os.listdir() if f.lower().endswith(".xlsx")), None)
    if not src:
        sys.exit("Usage: python bank2zen.py file.xlsx")
    res = convert(src)
    print("out_zenmoney.csv ready" if res == "ok"
          else "fill new_desc.xlsx and run again")
