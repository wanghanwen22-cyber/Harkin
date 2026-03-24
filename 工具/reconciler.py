#!/usr/bin/env python3
"""Simple automated reconciliation tool for two CSV files.

Compares "bank" records with "ledger" records and outputs:
- matched.csv
- unmatched_bank.csv
- unmatched_ledger.csv
- summary.json
"""

from __future__ import annotations

import argparse
import csv
import json
from dataclasses import dataclass
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

DATE_FORMATS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y%m%d",
    "%d/%m/%Y",
    "%m/%d/%Y",
)


@dataclass
class NormalizedRow:
    index: int
    raw: dict[str, str]
    amount: Decimal
    tx_date: date
    tx_id: str | None


def parse_amount(value: str) -> Decimal:
    cleaned = value.replace(",", "").strip()
    try:
        return Decimal(cleaned)
    except InvalidOperation as exc:
        raise ValueError(f"invalid amount: {value!r}") from exc


def parse_date(value: str) -> date:
    v = value.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(v, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"invalid date: {value!r}, supported formats: {', '.join(DATE_FORMATS)}")


def load_csv(
    path: Path,
    amount_field: str,
    date_field: str,
    id_field: str | None,
    encoding: str,
) -> list[NormalizedRow]:
    rows: list[NormalizedRow] = []
    with path.open("r", newline="", encoding=encoding) as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise ValueError(f"{path} has no header row")

        missing = [field for field in (amount_field, date_field) if field not in reader.fieldnames]
        if missing:
            raise ValueError(f"{path} missing fields: {', '.join(missing)}")
        if id_field and id_field not in reader.fieldnames:
            raise ValueError(f"{path} missing id field: {id_field}")

        for i, raw in enumerate(reader, start=1):
            amount = parse_amount(raw[amount_field])
            tx_date = parse_date(raw[date_field])
            tx_id = raw[id_field].strip() if id_field and raw[id_field].strip() else None
            rows.append(NormalizedRow(index=i, raw=raw, amount=amount, tx_date=tx_date, tx_id=tx_id))
    return rows


def date_delta_days(a: date, b: date) -> int:
    return abs((a - b).days)


def reconcile(
    bank_rows: list[NormalizedRow],
    ledger_rows: list[NormalizedRow],
    amount_tolerance: Decimal,
    date_tolerance_days: int,
) -> tuple[list[tuple[NormalizedRow, NormalizedRow, str]], list[NormalizedRow], list[NormalizedRow]]:
    matched: list[tuple[NormalizedRow, NormalizedRow, str]] = []

    bank_unmatched: set[int] = set(range(len(bank_rows)))
    ledger_unmatched: set[int] = set(range(len(ledger_rows)))

    # Pass 1: ID exact match (if ID exists on both sides)
    ledger_id_map: dict[str, list[int]] = {}
    for idx, row in enumerate(ledger_rows):
        if row.tx_id:
            ledger_id_map.setdefault(row.tx_id, []).append(idx)

    for b_idx, b in enumerate(bank_rows):
        if b_idx not in bank_unmatched or not b.tx_id:
            continue
        candidate_ids = ledger_id_map.get(b.tx_id, [])
        for l_idx in candidate_ids:
            if l_idx not in ledger_unmatched:
                continue
            l = ledger_rows[l_idx]
            if abs(b.amount - l.amount) <= amount_tolerance and date_delta_days(b.tx_date, l.tx_date) <= date_tolerance_days:
                matched.append((b, l, "id_match"))
                bank_unmatched.remove(b_idx)
                ledger_unmatched.remove(l_idx)
                break

    # Pass 2: Greedy amount + date match
    for b_idx in list(bank_unmatched):
        b = bank_rows[b_idx]
        best_l_idx: int | None = None
        best_score: tuple[Any, Any] | None = None

        for l_idx in ledger_unmatched:
            l = ledger_rows[l_idx]
            amount_diff = abs(b.amount - l.amount)
            date_diff = date_delta_days(b.tx_date, l.tx_date)
            if amount_diff > amount_tolerance or date_diff > date_tolerance_days:
                continue

            score = (amount_diff, date_diff)
            if best_score is None or score < best_score:
                best_score = score
                best_l_idx = l_idx

        if best_l_idx is not None:
            l = ledger_rows[best_l_idx]
            matched.append((b, l, "amount_date_match"))
            bank_unmatched.remove(b_idx)
            ledger_unmatched.remove(best_l_idx)

    return (
        matched,
        [bank_rows[i] for i in sorted(bank_unmatched)],
        [ledger_rows[i] for i in sorted(ledger_unmatched)],
    )


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Automated reconciliation for two CSV files")
    parser.add_argument("--bank", required=True, help="Path to bank CSV")
    parser.add_argument("--ledger", required=True, help="Path to ledger CSV")
    parser.add_argument("--out-dir", default="./reconcile_output", help="Output directory")

    parser.add_argument("--bank-amount-field", default="amount")
    parser.add_argument("--bank-date-field", default="date")
    parser.add_argument("--bank-id-field", default="tx_id")

    parser.add_argument("--ledger-amount-field", default="amount")
    parser.add_argument("--ledger-date-field", default="date")
    parser.add_argument("--ledger-id-field", default="tx_id")

    parser.add_argument("--amount-tolerance", default="0.00", help="Decimal amount tolerance, e.g. 0.01")
    parser.add_argument("--date-tolerance-days", type=int, default=0, help="Allowed date delta in days")
    parser.add_argument("--encoding", default="utf-8")

    args = parser.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    amount_tolerance = parse_amount(args.amount_tolerance)

    bank_rows = load_csv(
        Path(args.bank),
        amount_field=args.bank_amount_field,
        date_field=args.bank_date_field,
        id_field=args.bank_id_field,
        encoding=args.encoding,
    )
    ledger_rows = load_csv(
        Path(args.ledger),
        amount_field=args.ledger_amount_field,
        date_field=args.ledger_date_field,
        id_field=args.ledger_id_field,
        encoding=args.encoding,
    )

    matched, bank_unmatched, ledger_unmatched = reconcile(
        bank_rows,
        ledger_rows,
        amount_tolerance=amount_tolerance,
        date_tolerance_days=args.date_tolerance_days,
    )

    # matched.csv
    matched_records: list[dict[str, Any]] = []
    for b, l, match_type in matched:
        record: dict[str, Any] = {
            "match_type": match_type,
            "bank_row": b.index,
            "ledger_row": l.index,
            "bank_amount": str(b.amount),
            "ledger_amount": str(l.amount),
            "bank_date": b.tx_date.isoformat(),
            "ledger_date": l.tx_date.isoformat(),
            "bank_tx_id": b.tx_id or "",
            "ledger_tx_id": l.tx_id or "",
        }
        matched_records.append(record)

    write_csv(
        out_dir / "matched.csv",
        matched_records,
        [
            "match_type",
            "bank_row",
            "ledger_row",
            "bank_amount",
            "ledger_amount",
            "bank_date",
            "ledger_date",
            "bank_tx_id",
            "ledger_tx_id",
        ],
    )

    # unmatched files keep original row plus parsed columns
    bank_unmatched_records = [
        {**r.raw, "_row": r.index, "_amount": str(r.amount), "_date": r.tx_date.isoformat(), "_tx_id": r.tx_id or ""}
        for r in bank_unmatched
    ]
    ledger_unmatched_records = [
        {**r.raw, "_row": r.index, "_amount": str(r.amount), "_date": r.tx_date.isoformat(), "_tx_id": r.tx_id or ""}
        for r in ledger_unmatched
    ]

    if bank_unmatched_records:
        bank_fields = list(bank_unmatched_records[0].keys())
        write_csv(out_dir / "unmatched_bank.csv", bank_unmatched_records, bank_fields)
    else:
        write_csv(out_dir / "unmatched_bank.csv", [], ["_row", "_amount", "_date", "_tx_id"])

    if ledger_unmatched_records:
        ledger_fields = list(ledger_unmatched_records[0].keys())
        write_csv(out_dir / "unmatched_ledger.csv", ledger_unmatched_records, ledger_fields)
    else:
        write_csv(out_dir / "unmatched_ledger.csv", [], ["_row", "_amount", "_date", "_tx_id"])

    summary = {
        "bank_total": len(bank_rows),
        "ledger_total": len(ledger_rows),
        "matched": len(matched),
        "bank_unmatched": len(bank_unmatched),
        "ledger_unmatched": len(ledger_unmatched),
        "amount_tolerance": str(amount_tolerance),
        "date_tolerance_days": args.date_tolerance_days,
    }

    with (out_dir / "summary.json").open("w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
