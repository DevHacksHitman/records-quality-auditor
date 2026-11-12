#!/usr/bin/env python3
"""records_audit — a zero-dependency records validator with an audit trail.

Purpose: check a records file (CSV or Excel) for the problems that turn into
audit findings — duplicates, missing required fields, malformed formats,
out-of-range values, unexpected codes — and leave behind a machine-readable
audit log of what was checked, when, and against which file.

Design notes:
* Standard library only. Excel (.xlsx) is read directly with ``zipfile`` +
  ``ElementTree`` — no pandas, no openpyxl, nothing to install.
* Read-only: the tool never writes to the input file. Its only output is a
  findings report and an append-only audit log.
* Exit codes are CI-friendly: 0 = clean, 1 = findings, 2 = usage/parse error.

Usage:
    python3 records_audit.py records.csv --rules rules.json
    python3 records_audit.py records.xlsx --rules rules.json --json
    python3 records_audit.py records.csv --rules rules.json --audit-log audit.jsonl

Rules file (JSON):
    {
      "required": ["surname", "first_name"],
      "unique": ["file_number"],
      "format": {"date_of_birth": "date", "email": "email", "phone": "phone_ng"},
      "allowed": {"gender": ["Male", "Female"]},
      "range": {"age": [0, 120]},
      "pii": true
    }
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree

VERSION = "0.1.0"
TOOL = "records-audit"

XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Named format patterns — deliberately permissive: the goal is to catch
# malformed entries, not to reject unusual-but-valid ones.
NAMED_PATTERNS: Dict[str, str] = {
    "date": r"^\d{4}-\d{2}-\d{2}$|^\d{1,2}/\d{1,2}/\d{4}$",
    "email": r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$",
    "phone_ng": r"^(\+?234|0)[789]\d{9}$",
    "integer": r"^-?\d+$",
    "decimal": r"^-?\d+(\.\d+)?$",
    "alpha": r"^[A-Za-z][A-Za-z \-']*$",
    "alnum": r"^[A-Za-z0-9][A-Za-z0-9 \-/]*$",
}

# Patterns used by the PII sweep: columns that look like personal data are
# reported so the data owner can decide on retention and access controls.
PII_PATTERNS: Dict[str, str] = {
    "email": NAMED_PATTERNS["email"],
    "phone": r"^(\+?\d[\d \-]{7,}\d)$",
    "national_id": r"^\d{11}$",
}

RULE_KEYS = {"required", "unique", "format", "allowed", "range", "pii"}


class RecordsError(Exception):
    """A usage, configuration or parsing problem (exit code 2)."""


# ---------------------------------------------------------------------------
# Reading records
# ---------------------------------------------------------------------------
def _column_index(ref: str) -> int:
    """'A' -> 0, 'B' -> 1, 'AA' -> 26."""
    index = 0
    for char in ref:
        if char.isalpha():
            index = index * 26 + (ord(char.upper()) - ord("A") + 1)
    return max(index - 1, 0)


def read_xlsx(path: Path) -> Tuple[List[str], List[List[str]]]:
    """Minimal, dependency-free .xlsx reader (first worksheet).

    Handles shared strings, inline strings and literal values — the three ways
    a spreadsheet stores cell text. Formulas are read as their cached value.
    """
    try:
        archive = zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        raise RecordsError(f"{path} is not a readable .xlsx file: {exc}") from exc

    with archive:
        names = archive.namelist()
        shared: List[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ElementTree.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall(f"{XLSX_NS}si"):
                shared.append("".join(node.text or "" for node in item.iter(f"{XLSX_NS}t")))

        sheets = sorted(
            (
                name
                for name in names
                if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
            ),
            key=lambda name: int(re.search(r"sheet(\d+)", name).group(1)),
        )
        if not sheets:
            raise RecordsError(f"{path} contains no worksheet")
        sheet = ElementTree.fromstring(archive.read(sheets[0]))

    table: List[List[str]] = []
    for row in sheet.iter(f"{XLSX_NS}row"):
        cells: Dict[int, str] = {}
        for cell in row.findall(f"{XLSX_NS}c"):
            reference = cell.get("r") or ""
            kind = cell.get("t")
            value = cell.find(f"{XLSX_NS}v")
            inline = cell.find(f"{XLSX_NS}is")
            if kind == "s" and value is not None and value.text is not None:
                position = int(value.text)
                text = shared[position] if 0 <= position < len(shared) else ""
            elif kind == "inlineStr" and inline is not None:
                text = "".join(node.text or "" for node in inline.iter(f"{XLSX_NS}t"))
            elif value is not None:
                text = value.text or ""
            else:
                text = ""
            index = _column_index(reference) if reference else len(cells)
            cells[index] = text.strip()
        if cells:
            width = max(cells) + 1
            table.append([cells.get(index, "") for index in range(width)])

    if not table:
        raise RecordsError(f"{path} contains no rows")
    header = [name or f"column_{index + 1}" for index, name in enumerate(table[0])]
    rows = [row + [""] * (len(header) - len(row)) for row in table[1:]]
    return header, rows


def read_csv(path: Path) -> Tuple[List[str], List[List[str]]]:
    """Read a CSV/TSV file; BOM-tolerant, blank lines skipped."""
    delimiter = "\t" if path.suffix.lower() == ".tsv" else ","
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle, delimiter=delimiter)
        table = [row for row in reader if any(cell.strip() for cell in row)]
    if not table:
        raise RecordsError(f"{path} contains no rows")
    header = [name.strip() or f"column_{index + 1}" for index, name in enumerate(table[0])]
    rows = [
        [cell.strip() for cell in row] + [""] * (len(header) - len(row))
        for row in table[1:]
    ]
    return header, rows


def load_records(path: Any) -> Tuple[List[str], List[List[str]]]:
    """Dispatch on extension: .xlsx via the built-in reader, else CSV/TSV."""
    path = Path(path)
    if not path.exists():
        raise RecordsError(f"no such file: {path}")
    if path.suffix.lower() == ".xlsx":
        return read_xlsx(path)
    return read_csv(path)


# ---------------------------------------------------------------------------
# Rules
# ---------------------------------------------------------------------------
def load_rules(path: Any) -> Dict[str, Any]:
    path = Path(path)
    if not path.exists():
        raise RecordsError(f"no such rules file: {path}")
    try:
        rules = json.loads(path.read_text(encoding="utf-8"))
    except ValueError as exc:
        raise RecordsError(f"{path} is not valid JSON: {exc}") from exc
    if not isinstance(rules, dict):
        raise RecordsError("rules file must contain a JSON object")
    unknown = sorted(set(rules) - RULE_KEYS)
    if unknown:
        raise RecordsError(
            f"unknown rule(s): {', '.join(unknown)} — allowed: {', '.join(sorted(RULE_KEYS))}"
        )
    for field, spec in (rules.get("format") or {}).items():
        try:
            re.compile(_pattern_for(spec))
        except re.error as exc:
            raise RecordsError(
                f"format rule for {field!r} is not a valid regular expression: {exc}"
            ) from exc
    return rules


def _pattern_for(spec: str) -> str:
    if spec in NAMED_PATTERNS:
        return NAMED_PATTERNS[spec]
    return spec  # a custom regular expression


def _finding(row: int, field: str, code: str, message: str, value: Any) -> Dict[str, Any]:
    return {"row": row, "field": field, "code": code, "message": message, "value": value}


def validate(
    header: Sequence[str], rows: Sequence[Sequence[str]], rules: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """Run every rule over the records; return findings in row order.

    ``row`` is the data-row number counting the header as row 1, so the first
    data row is 2 — the same numbering a spreadsheet shows.
    """
    findings: List[Dict[str, Any]] = []
    positions = {name: index for index, name in enumerate(header)}

    def missing_fields(fields: Iterable[str]) -> None:
        for field in fields:
            if field not in positions:
                raise RecordsError(f"rule refers to unknown column: {field!r}")

    missing_fields(rules.get("required", []))
    for spec in rules.get("unique", []):
        missing_fields(spec if isinstance(spec, list) else [spec])
    missing_fields(rules.get("format", {}))
    missing_fields(rules.get("allowed", {}))
    missing_fields(rules.get("range", {}))

    seen: Dict[Any, int] = {}

    for offset, row in enumerate(rows):
        number = offset + 2
        for field in rules.get("required", []):
            if not str(row[positions[field]]).strip():
                findings.append(
                    _finding(number, field, "required", f"{field} is empty", "")
                )

        for field in rules.get("format", {}):
            value = str(row[positions[field]]).strip()
            if not value:
                continue  # required-ness is the 'required' rule's job
            pattern = _pattern_for(rules["format"][field])
            if not re.match(pattern, value):
                findings.append(
                    _finding(number, field, "format", f"{field} does not match {rules['format'][field]}", value)
                )

        for field, allowed in rules.get("allowed", {}).items():
            value = str(row[positions[field]]).strip()
            if value and value not in allowed:
                findings.append(
                    _finding(number, field, "allowed", f"{field} is not one of: {', '.join(map(str, allowed))}", value)
                )

        for field, bounds in rules.get("range", {}).items():
            value = str(row[positions[field]]).strip()
            if not value:
                continue
            try:
                number_value = float(value)
            except ValueError:
                findings.append(_finding(number, field, "range", f"{field} is not numeric", value))
                continue
            low, high = bounds
            if number_value < low or number_value > high:
                findings.append(
                    _finding(number, field, "range", f"{field} is outside {low}–{high}", value)
                )

        for spec in rules.get("unique", []):
            fields = spec if isinstance(spec, list) else [spec]
            key = tuple(str(row[positions[field]]).strip().lower() for field in fields)
            if not any(key):
                continue
            label = " + ".join(fields)
            if key in seen:
                findings.append(
                    _finding(number, label, "unique", f"duplicate of row {seen[key]}", " / ".join(key))
                )
            else:
                seen[key] = number

    if rules.get("pii"):
        for field, index in positions.items():
            hits = sum(
                1
                for row in rows
                if any(re.match(pattern, str(row[index]).strip()) for pattern in PII_PATTERNS.values())
            )
            if hits:
                findings.append(
                    _finding(0, field, "pii", f"column looks like personal data ({hits} row(s))", "")
                )

    return findings


# ---------------------------------------------------------------------------
# Reporting and the audit trail
# ---------------------------------------------------------------------------
def summarise(findings: Sequence[Dict[str, Any]]) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for finding in findings:
        counts[finding["code"]] = counts.get(finding["code"], 0) + 1
    return counts


def render_text(
    path: Path, header: Sequence[str], rows: Sequence[Sequence[str]],
    findings: Sequence[Dict[str, Any]], limit: int = 50,
) -> str:
    if not findings:
        return f"PASS — {len(rows)} record(s) in {path.name} with no findings."
    lines = [
        f"FINDINGS — {len(findings)} in {path.name} ({len(rows)} record(s), "
        f"{len(header)} column(s))",
        "",
    ]
    for finding in findings[:limit]:
        where = f"row {finding['row']}" if finding["row"] else "column"
        lines.append(f"  {where:>8}  [{finding['code']}] {finding['message']}")
        if finding["value"]:
            lines.append(f"            value: {finding['value']}")
    if len(findings) > limit:
        lines.append(f"  … {len(findings) - limit} more (use --json for the full list)")
    counts = summarise(findings)
    lines.append("")
    lines.append("Summary: " + ", ".join(f"{code}={count}" for code, count in sorted(counts.items())))
    return "\n".join(lines)


def file_digest(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_audit_log(
    log_path: Path, path: Path, header: Sequence[str], rows: Sequence[Sequence[str]],
    rules: Dict[str, Any], findings: Sequence[Dict[str, Any]],
) -> Dict[str, Any]:
    """Append one JSON line recording what was checked, against what."""
    entry = {
        "tool": TOOL,
        "version": VERSION,
        "checked_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "file": path.name,
        "sha256": file_digest(path),
        "columns": len(header),
        "records": len(rows),
        "rules": sorted(rules.keys()),
        "findings": len(findings),
        "findings_by_code": summarise(findings),
        "result": "FAIL" if findings else "PASS",
    }
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(entry, separators=(",", ":")) + "\n")
    return entry


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="records_audit.py",
        description="Validate CSV/Excel records and write an audit trail (zero dependencies).",
    )
    parser.add_argument("records", help="CSV, TSV or .xlsx file to check")
    parser.add_argument("--rules", required=True, help="JSON rules file")
    parser.add_argument("--audit-log", metavar="PATH", help="append a JSON audit line to this file")
    parser.add_argument("--json", action="store_true", help="print findings as JSON")
    parser.add_argument("--limit", type=int, default=50, help="max findings to print (text mode)")
    parser.add_argument("--version", action="version", version=f"{TOOL} {VERSION}")
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    path = Path(args.records)
    try:
        header, rows = load_records(path)
        rules = load_rules(args.rules)
        findings = validate(header, rows, rules)
    except RecordsError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    if args.audit_log:
        write_audit_log(Path(args.audit_log), path, header, rows, rules, findings)

    if args.json:
        print(json.dumps(
            {
                "file": path.name,
                "records": len(rows),
                "columns": len(header),
                "findings": findings,
                "summary": summarise(findings),
            },
            indent=2,
        ))
    else:
        print(render_text(path, header, rows, findings, args.limit))

    return 1 if findings else 0


if __name__ == "__main__":
    raise SystemExit(main())
