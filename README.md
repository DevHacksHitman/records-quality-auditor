# Records Quality Auditor

A **zero-dependency** records validator for CSV and Excel files that leaves
behind a machine-readable **audit trail**.

Point it at a records file and a rules file. It reports duplicates, missing
required fields, malformed values, out-of-range numbers and unexpected codes —
and appends one JSON line recording exactly what was checked, when, and against
which file (with a SHA-256 of the input).

| | |
|---|---|
| Dependencies | **0** — standard library only (Excel is read with `zipfile` + `ElementTree`) |
| Code | **1 file**, 412 lines (333 logic) |
| Tests | **35**, zero-network, no packages installed in CI |
| Inputs | `.csv`, `.tsv`, `.xlsx` |
| Outputs | text report · `--json` · append-only `--audit-log` |
| Exit codes | `0` clean · `1` findings · `2` usage/parse error (CI-friendly) |

## Why this exists

Records work fails quietly. A duplicated file number, a date in the wrong
format, a missing surname or a birth year typed as `200` will pass a manual
skim and surface months later as an audit finding. The fix is a check that runs
the same way every time and **records that it ran** — which is exactly what a
records office, a registry or an audit function needs to be able to show.

This is the same discipline as the other repositories: one file, no
dependencies, tests that never touch the network, and a README that states its
limits instead of overclaiming.

## Quickstart

```bash
# validate a CSV against the example rules
python3 records_audit.py records.csv --rules rules.example.json

# same, but write an audit line (append-only JSONL)
python3 records_audit.py records.csv --rules rules.example.json --audit-log audit.jsonl

# machine-readable output for a pipeline
python3 records_audit.py records.xlsx --rules rules.json --json
```

Example run:

```
FINDINGS — 8 in records.csv (4 record(s), 7 column(s))

     row 4  [format] date_of_birth does not match date
            value: not-a-date
     row 4  [allowed] gender is not one of: Male, Female
            value: Unknown
     row 4  [range] age is not numeric
            value: not-a-number
     row 5  [required] first_name is empty
     row 5  [range] age is outside 0–120
            value: 200
     row 5  [unique] duplicate of row 3
            value: 1002
    column  [pii] column looks like personal data (2 row(s))

Summary: allowed=1, format=1, pii=2, range=2, required=1, unique=1
```

Row numbers count the header as row 1, so they match what you see in Excel.

## Rules

`--rules` takes a JSON object. Every rule is optional.

| Key | Shape | Checks |
|---|---|---|
| `required` | `["surname", "file_number"]` | field is not empty |
| `unique` | `["file_number"]` or `[["surname","file_number"]]` | no duplicates (case-insensitive; composite keys supported) |
| `format` | `{"date_of_birth": "date"}` | value matches a named pattern or your own regular expression |
| `allowed` | `{"gender": ["Male","Female"]}` | value is in the permitted set |
| `range` | `{"age": [0, 120]}` | numeric and within bounds |
| `pii` | `true` | reports which columns look like personal data (emails, phone numbers, 11-digit IDs) |

Named formats: `date`, `email`, `phone_ng`, `integer`, `decimal`, `alpha`,
`alnum`. Anything else is treated as a raw regular expression.

## The audit trail

```json
{"tool":"records-audit","version":"0.1.0","checked_at":"2026-11-12T09:44:46Z","file":"records.csv","sha256":"af2f8197…","columns":7,"records":4,"rules":["allowed","format","pii","range","required","unique"],"findings":8,"findings_by_code":{"format":1,"allowed":1,"range":2,"required":1,"unique":1,"pii":2},"result":"FAIL"}
```

Append-only by design: entries are added, never rewritten, so the log itself is
evidence of what was run and against which bytes. The `sha256` field is what
makes it worth anything — it ties a result to one exact file, not to a
filename.

## Design decisions

- **No dependencies, including for Excel.** `.xlsx` is a ZIP of XML, so the
  tool reads `sharedStrings.xml` and the first worksheet directly. Nothing to
  install means nothing to keep patched on a locked-down machine.
- **Read-only.** The input file is never modified and never rewritten. Output
  is a report plus the audit log.
- **Exit codes designed for automation:** a cron job or CI step can gate on `1`.
- **PII sweep as a governance signal**, not a compliance claim: it tells a data
  owner where personal data sits so retention and access can be decided
  deliberately.
- **Case-insensitive duplicate detection** — `Okafor` and `okafor` are the same
  record in practice.

## Honest limitations

- **`.xlsx` support is deliberately minimal:** the first worksheet only, cached
  formula values rather than evaluation, no `.xls` (legacy binary) support, and
  no styling or merged-cell awareness. For multi-sheet workbooks, split or
  export first.
- **No fuzzy matching.** `Okafor` vs `Okafor ` (trailing space) is handled by
  trimming, but `Okafor` vs `Okaforr` is a different value. Near-duplicate
  detection is not implemented.
- **The PII sweep is heuristic.** It flags columns by pattern; it does not
  prove a column is or is not personal data, and it does not redact or encrypt.
- **This is not a compliance certification.** It produces evidence a human can
  review; it does not make a dataset legally compliant.
- **UTF-8 (with BOM) and delimiter-based files only** — no fixed-width formats.

## Tests

```bash
python3 -m unittest discover -s tests -t . -v
```

35 tests, zero network, no third-party packages: both readers (including a real
`.xlsx` built in-test with `zipfile`), every rule type, row numbering, the
audit log's append/FAIL/PASS behaviour, the CLI exit codes, and an AST test
asserting the tool imports only the standard library.

## Repository hygiene

- Commits use the GitHub noreply identity; no personal addresses in history.
- `githooks/pre-push` is a blocking sensitivity gate. Enable it once per clone:

      git config core.hooksPath githooks

  Its needles are written self-avoidantly (character-class breaks) so the hook
  can never match its own source and block every push forever.
- `.gitignore` excludes `*.csv`, `*.tsv`, `*.xlsx` and `audit.jsonl` so a real
  records file — which may contain personal data — cannot be committed by
  accident.

## Roadmap

- Multi-sheet `.xlsx` selection (`--sheet NAME`).
- Near-duplicate detection (edit distance) behind an explicit flag.
- `--redact` output that writes a de-identified copy of the records.
- Optional watch mode for a records folder, so every drop is audited.

## License

MIT — see `LICENSE`.
