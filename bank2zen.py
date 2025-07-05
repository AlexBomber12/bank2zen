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
    # ищем строку, где начинается шапка
    hdr = next(i for i, v in enumerate(
        pd.read_excel(xlsx, header=None).iloc[:, 0]) if str(v).strip() == "Data")

    df = (
        pd.read_excel(xlsx, header=hdr,
                      dtype={"Entrate": float, "Uscite": float})
        .dropna(subset=["Data"])
    )

    df["Norm"] = df["Descrizione_Completa"].apply(lambda s: normalize(str(s)))

    cat_map = _load(CATS_FILE)
    acc_map = _load(ACC_FILE)

    df["Category"]  = df["Norm"].apply(lambda n: _lookup(n, cat_map))
    df["AccountTo"] = df["Norm"].apply(lambda n: _lookup(n, acc_map))

    # ── пост-обработка строк ─────────────────────────────
    def split(row):
        desc_short = str(row["Descrizione"])          # ← короткое описание
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
