"""Bank statement parser supporting common CSV formats."""

import csv
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional


@dataclass
class Transaction:
    date: datetime
    description: str
    amount: float  # negative = debit, positive = credit
    balance: Optional[float] = None
    category: Optional[str] = None
    raw: dict = field(default_factory=dict)

    @property
    def is_debit(self) -> bool:
        return self.amount < 0

    @property
    def is_credit(self) -> bool:
        return self.amount > 0


# Column name aliases used across different bank exports
_DATE_COLS = {"date", "transaction date", "trans date", "value date", "posted date", "transaction_date"}
_DESC_COLS = {"description", "details", "narrative", "transaction details", "memo", "payee", "merchant"}
_AMOUNT_COLS = {"amount", "transaction amount", "value"}
_DEBIT_COLS = {"debit", "debit amount", "withdrawals", "payment"}
_CREDIT_COLS = {"credit", "credit amount", "deposits", "receipts"}
_BALANCE_COLS = {"balance", "running balance", "available balance"}

_DATE_FORMATS = [
    "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y",
    "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%Y/%m/%d",
]


def _parse_date(value: str) -> datetime:
    value = value.strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    raise ValueError(f"Unrecognised date format: {value!r}")


def _parse_amount(value: str) -> float:
    value = re.sub(r"[£$€,\s]", "", value.strip())
    if value in ("", "-", "–"):
        return 0.0
    # Handle parentheses as negatives: (123.45) → -123.45
    if value.startswith("(") and value.endswith(")"):
        value = "-" + value[1:-1]
    return float(value)


def _normalise_header(h: str) -> str:
    return h.strip().lower().replace("_", " ")


def _match_col(headers: list[str], candidates: set[str]) -> Optional[int]:
    for i, h in enumerate(headers):
        if _normalise_header(h) in candidates:
            return i
    return None


def parse_csv(path: str | Path) -> list[Transaction]:
    path = Path(path)
    with path.open(newline="", encoding="utf-8-sig") as fh:
        # Detect dialect
        sample = fh.read(4096)
        fh.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample)
        except csv.Error:
            dialect = csv.excel
        reader = csv.DictReader(fh, dialect=dialect)
        rows = list(reader)

    if not rows:
        raise ValueError(f"No data found in {path}")

    headers = list(rows[0].keys())
    norm = [_normalise_header(h) for h in headers]

    date_i = _match_col(headers, _DATE_COLS)
    desc_i = _match_col(headers, _DESC_COLS)
    amount_i = _match_col(headers, _AMOUNT_COLS)
    debit_i = _match_col(headers, _DEBIT_COLS)
    credit_i = _match_col(headers, _CREDIT_COLS)
    balance_i = _match_col(headers, _BALANCE_COLS)

    if date_i is None:
        raise ValueError(f"Could not find a date column. Headers: {headers}")
    if desc_i is None:
        raise ValueError(f"Could not find a description column. Headers: {headers}")
    if amount_i is None and (debit_i is None and credit_i is None):
        raise ValueError(f"Could not find amount columns. Headers: {headers}")

    date_col = headers[date_i]
    desc_col = headers[desc_i]
    amount_col = headers[amount_i] if amount_i is not None else None
    debit_col = headers[debit_i] if debit_i is not None else None
    credit_col = headers[credit_i] if credit_i is not None else None
    balance_col = headers[balance_i] if balance_i is not None else None

    transactions: list[Transaction] = []
    for row in rows:
        raw_date = row.get(date_col, "").strip()
        if not raw_date:
            continue
        try:
            date = _parse_date(raw_date)
        except ValueError:
            continue

        description = row.get(desc_col, "").strip()

        if amount_col:
            raw_amt = row.get(amount_col, "0")
            amount = _parse_amount(raw_amt)
        else:
            debit = _parse_amount(row.get(debit_col, "") or "0")
            credit = _parse_amount(row.get(credit_col, "") or "0")
            amount = credit - debit  # credits positive, debits negative

        balance = None
        if balance_col:
            raw_bal = row.get(balance_col, "").strip()
            if raw_bal:
                try:
                    balance = _parse_amount(raw_bal)
                except ValueError:
                    pass

        transactions.append(Transaction(
            date=date,
            description=description,
            amount=amount,
            balance=balance,
            raw=dict(row),
        ))

    transactions.sort(key=lambda t: t.date)
    return transactions
