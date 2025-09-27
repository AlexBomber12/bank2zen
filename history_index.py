from __future__ import annotations
import os
import re
import sqlite3
import time
import hashlib
from pathlib import Path


def _state_dir() -> Path:
    if os.name == "nt":
        base = os.getenv("APPDATA") or str(Path.home())
        p = Path(base) / "bank2zen"
    else:
        p = Path.home() / ".bank2zen"
    p.mkdir(parents=True, exist_ok=True)
    return p


def db_path() -> Path:
    return _state_dir() / "seen.sqlite"


def ensure_db() -> sqlite3.Connection:
    con = sqlite3.connect(db_path())
    con.execute(
        """
        CREATE TABLE IF NOT EXISTS seen(
            key TEXT PRIMARY KEY,
            date TEXT,
            direction TEXT,
            amount REAL,
            account TEXT,
            token1 TEXT,
            token2 TEXT,
            source TEXT,
            created_at REAL
        )
    """
    )
    return con


def _norm_text(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def _tokens(s: str):
    # optional stable IDs from text: last4, CRO/TRN, generic ID
    last4 = None
    m = re.search(r"carta\s*n\.?\s*\*+\s*(\d{3,4})", s, re.I)
    if m:
        last4 = m.group(1)

    crotrn = None
    m = re.search(r"\b(CRO|TRN)\s*[: ]?\s*([A-Z0-9]{6,})", s, re.I)
    if m:
        crotrn = f"{m.group(1).upper()}:{m.group(2).upper()}"

    ppid = None
    m = re.search(r"\bID\s*[:#]?\s*([A-Z0-9]{10,})", s, re.I)
    if m:
        ppid = m.group(1).upper()
    return (last4 or ""), (crotrn or ""), (ppid or "")


def fingerprint(
    date_iso: str,
    direction: str,
    amount_cents: int,
    descr_full: str,
    account: str = "",
    extra: str = "",
) -> str:
    dnorm = _norm_text(descr_full)
    t1, t2, t3 = _tokens(dnorm)
    key_str = "|".join(
        [
            date_iso,
            direction,
            str(amount_cents),
            (account or ""),
            t1,
            t2,
            t3,
            dnorm,
        ]
    )
    return hashlib.sha1(key_str.encode("utf-8")).hexdigest()


def seen_lookup(con: sqlite3.Connection, keys: list[str]) -> set[str]:
    if not keys:
        return set()
    found = set()
    for i in range(0, len(keys), 500):
        chunk = keys[i : i + 500]
        q = "SELECT key FROM seen WHERE key IN (%s)" % ",".join("?" * len(chunk))
        for (k,) in con.execute(q, chunk):
            found.add(k)
    return found


def mark_seen(con: sqlite3.Connection, rows):
    """
    rows: iterable of tuples (key, date_iso, direction, amount, account, token1, token2, source)
    """
    now = time.time()
    con.executemany(
        "INSERT OR IGNORE INTO seen(key,date,direction,amount,account,token1,token2,source,created_at) VALUES (?,?,?,?,?,?,?,?,?)",
        [(k, d, dr, amt, acc, t1, t2, src, now) for (k, d, dr, amt, acc, t1, t2, src) in rows],
    )
    con.commit()
