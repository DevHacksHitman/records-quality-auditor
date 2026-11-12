"""Tests for records_audit: readers, rules, reporting, audit log, CLI.

Zero network, zero third-party packages. The .xlsx fixture is built with
``zipfile`` so the suite proves the dependency-free Excel path for real.
"""

import ast
import io
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from contextlib import redirect_stdout
from pathlib import Path

import records_audit as ra

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "records_audit.py"

SHARED_STRINGS = """<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>surname</t></si><si><t>first_name</t></si><si><t>file_number</t></si>
</sst>"""

SHEET = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Okafor</t></is></c>
      <c r="B2" t="inlineStr"><is><t>Ada</t></is></c>
      <c r="C2"><v>1001</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Adeyemi</t></is></c>
      <c r="B3" t="inlineStr"><is><t>Tunde</t></is></c>
      <c r="C3"><v>1002</v></c>
    </row>
  </sheetData>
</worksheet>"""


def build_xlsx(path: Path) -> Path:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/sharedStrings.xml", SHARED_STRINGS)
        archive.writestr("xl/worksheets/sheet1.xml", SHEET)
    return path


class TestCsvReader(unittest.TestCase):
    def write(self, text, name="records.csv", encoding="utf-8"):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = Path(tmp.name) / name
        path.write_text(text, encoding=encoding)
        return path

    def test_basic(self):
        path = self.write("surname,first_name\nOkafor,Ada\nAdeyemi,Tunde\n")
        header, rows = ra.load_records(path)
        self.assertEqual(header, ["surname", "first_name"])
        self.assertEqual(rows, [["Okafor", "Ada"], ["Adeyemi", "Tunde"]])

    def test_bom_and_blank_lines_and_short_rows(self):
        path = self.write("\ufeffsurname,first_name,file_number\n\nOkafor,Ada\n")
        header, rows = ra.load_records(path)
        self.assertEqual(header[0], "surname")  # BOM stripped
        self.assertEqual(rows, [["Okafor", "Ada", ""]])  # padded

    def test_quoted_commas(self):
        path = self.write('surname,address\nOkafor,"12 Broad St, Lagos"\n')
        _, rows = ra.load_records(path)
        self.assertEqual(rows[0][1], "12 Broad St, Lagos")

    def test_tsv(self):
        path = self.write("a\tb\n1\t2\n", name="records.tsv")
        self.assertEqual(ra.load_records(path), (["a", "b"], [["1", "2"]]))

    def test_missing_and_empty(self):
        with self.assertRaises(ra.RecordsError):
            ra.load_records("/nonexistent/records.csv")
        with self.assertRaises(ra.RecordsError):
            ra.load_records(self.write("\n\n"))


class TestXlsxReader(unittest.TestCase):
    def test_reads_shared_and_inline_strings(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = build_xlsx(Path(tmp) / "records.xlsx")
            header, rows = ra.load_records(path)
        self.assertEqual(header, ["surname", "first_name", "file_number"])
        self.assertEqual(rows[0], ["Okafor", "Ada", "1001"])
        self.assertEqual(rows[1], ["Adeyemi", "Tunde", "1002"])

    def test_not_a_zip(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad = Path(tmp) / "records.xlsx"
            bad.write_text("not a spreadsheet")
            with self.assertRaises(ra.RecordsError):
                ra.load_records(bad)


class TestRules(unittest.TestCase):
    def write_rules(self, payload):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = Path(tmp.name) / "rules.json"
        path.write_text(json.dumps(payload), encoding="utf-8")
        return path

    def test_valid(self):
        rules = ra.load_rules(self.write_rules({"required": ["surname"]}))
        self.assertEqual(rules["required"], ["surname"])

    def test_unknown_key_rejected(self):
        with self.assertRaises(ra.RecordsError):
            ra.load_rules(self.write_rules({"requred": ["surname"]}))

    def test_bad_json_rejected(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = Path(tmp.name) / "rules.json"
        path.write_text("{not json", encoding="utf-8")
        with self.assertRaises(ra.RecordsError):
            ra.load_rules(path)

    def test_missing_file(self):
        with self.assertRaises(ra.RecordsError):
            ra.load_rules("/nonexistent/rules.json")

    def test_invalid_custom_regex_is_a_usage_error(self):
        with self.assertRaises(ra.RecordsError):
            ra.load_rules(self.write_rules({"format": {"surname": "([a-z"}}))


class TestValidate(unittest.TestCase):
    header = ["surname", "first_name", "file_number", "date_of_birth", "gender", "age"]

    def rows(self):
        return [
            ["Okafor", "Ada", "1001", "1990-05-01", "Female", "34"],
            ["Adeyemi", "Tunde", "1002", "12/11/1988", "Male", "36"],
            ["okafor", "Ada", "", "not a date", "Unknown", "not-a-number"],
            ["Bello", "", "1002", "", "Female", "200"],
        ]

    def test_required(self):
        findings = ra.validate(self.header, self.rows(), {"required": ["first_name", "file_number"]})
        codes = [(f["row"], f["field"], f["code"]) for f in findings]
        self.assertEqual(codes, [(4, "file_number", "required"), (5, "first_name", "required")])

    def test_format_named_patterns(self):
        findings = ra.validate(self.header, self.rows(), {"format": {"date_of_birth": "date"}})
        self.assertEqual([(f["row"], f["code"]) for f in findings], [(4, "format")])

    def test_format_custom_pattern(self):
        findings = ra.validate(self.header, [["AB-12"]], {"format": {"surname": r"^[A-Z]{2}-\d{2}$"}})
        self.assertEqual(findings, [])

    def test_allowed_values(self):
        findings = ra.validate(self.header, self.rows(), {"allowed": {"gender": ["Male", "Female"]}})
        self.assertEqual([(f["row"], f["value"]) for f in findings], [(4, "Unknown")])

    def test_range_numeric_and_non_numeric(self):
        findings = ra.validate(self.header, self.rows(), {"range": {"age": [0, 120]}})
        self.assertEqual([(f["row"], f["code"]) for f in findings], [(4, "range"), (5, "range")])

    def test_unique_is_case_insensitive(self):
        findings = ra.validate(self.header, self.rows(), {"unique": ["surname"]})
        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0]["row"], 4)
        self.assertIn("duplicate of row 2", findings[0]["message"])

    def test_unique_composite_key(self):
        header = ["surname", "file_number"]
        rows = [["Okafor", "1001"], ["Adeyemi", "1002"], ["Okafor", "1001"]]
        findings = ra.validate(header, rows, {"unique": [["surname", "file_number"]]})
        self.assertEqual([f["code"] for f in findings], ["unique"])
        self.assertEqual(findings[0]["row"], 4)
        self.assertEqual(findings[0]["field"], "surname + file_number")

    def test_empty_values_are_not_duplicates(self):
        findings = ra.validate(["file_number"], [[""], [""]], {"unique": ["file_number"]})
        self.assertEqual(findings, [])

    def test_row_numbering_starts_at_two(self):
        findings = ra.validate(["a"], [[""]], {"required": ["a"]})
        self.assertEqual(findings[0]["row"], 2)

    def test_unknown_column_in_rule(self):
        with self.assertRaises(ra.RecordsError):
            ra.validate(self.header, self.rows(), {"required": ["nope"]})
        with self.assertRaises(ra.RecordsError):
            ra.validate(self.header, self.rows(), {"unique": [["surname", "nope"]]})

    def test_pii_sweep(self):
        header = ["name", "email", "notes"]
        rows = [["Ada", "ada@example.com", "ok"], ["Tunde", "tunde@example.com", "ok"]]
        findings = ra.validate(header, rows, {"pii": True})
        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0]["field"], "email")
        self.assertEqual(findings[0]["code"], "pii")
        self.assertEqual(findings[0]["row"], 0)

    def test_no_findings_on_clean_data(self):
        rows = [["Okafor", "Ada", "1001", "1990-05-01", "Female", "34"]]
        rules = {
            "required": ["surname"],
            "unique": ["file_number"],
            "format": {"date_of_birth": "date"},
            "allowed": {"gender": ["Male", "Female"]},
            "range": {"age": [0, 120]},
        }
        self.assertEqual(ra.validate(self.header, rows, rules), [])


class TestReportingAndAuditLog(unittest.TestCase):
    def test_summarise(self):
        findings = [
            {"code": "required"}, {"code": "required"}, {"code": "format"},
        ]
        self.assertEqual(ra.summarise(findings), {"required": 2, "format": 1})

    def test_render_pass_and_fail(self):
        rules = {"required": ["surname"]}
        header, rows = ["surname"], [["Okafor"]]
        self.assertIn("PASS", ra.render_text(Path("r.csv"), header, rows, []))
        findings = ra.validate(header, [[""]], rules)
        text = ra.render_text(Path("r.csv"), header, rows, findings)
        self.assertIn("FINDINGS", text)
        self.assertIn("[required]", text)

    def test_render_truncates(self):
        findings = [{"row": i, "field": "a", "code": "required", "message": "m", "value": ""} for i in range(2, 12)]
        text = ra.render_text(Path("r.csv"), ["a"], [["x"]] * 10, findings, limit=3)
        self.assertIn("7 more", text)

    def test_digest_is_stable_and_content_sensitive(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "records.csv"
            path.write_text("a\n1\n", encoding="utf-8")
            first = ra.file_digest(path)
            self.assertEqual(first, ra.file_digest(path))
            path.write_text("a\n2\n", encoding="utf-8")
            self.assertNotEqual(first, ra.file_digest(path))

    def test_audit_log_appends_and_records_outcome(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "records.csv"
            source.write_text("surname,notes\n,missing name\nOkafor,fine\n", encoding="utf-8")
            log = Path(tmp) / "audit.jsonl"
            header, rows = ra.load_records(source)
            rules = {"required": ["surname"]}
            findings = ra.validate(header, rows, rules)
            self.assertEqual(len(findings), 1)
            ra.write_audit_log(log, source, header, rows, rules, findings)
            ra.write_audit_log(log, source, header, rows, rules, [])
            lines = [json.loads(line) for line in log.read_text(encoding="utf-8").splitlines()]
        self.assertEqual(len(lines), 2)
        self.assertEqual(lines[0]["result"], "FAIL")
        self.assertEqual(lines[1]["result"], "PASS")
        self.assertEqual(lines[0]["tool"], "records-audit")
        self.assertEqual(len(lines[0]["sha256"]), 64)
        self.assertEqual(lines[0]["records"], 2)


class TestCli(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)
        self.records = self.dir / "records.csv"
        self.records.write_text(
            "surname,file_number,date_of_birth\nOkafor,1001,1990-05-01\n,1001,oops\n",
            encoding="utf-8",
        )
        self.rules = self.dir / "rules.json"
        self.rules.write_text(
            json.dumps({"required": ["surname"], "unique": ["file_number"], "format": {"date_of_birth": "date"}}),
            encoding="utf-8",
        )

    def run_cli(self, *arguments):
        return subprocess.run(
            [sys.executable, str(SCRIPT), *arguments],
            capture_output=True, text=True, cwd=str(ROOT), timeout=60,
        )

    def test_exit_1_with_findings(self):
        result = self.run_cli(str(self.records), "--rules", str(self.rules))
        self.assertEqual(result.returncode, 1)
        self.assertIn("FINDINGS", result.stdout)

    def test_exit_0_when_clean(self):
        clean = self.dir / "clean.csv"
        clean.write_text("surname,file_number,date_of_birth\nOkafor,1001,1990-05-01\n", encoding="utf-8")
        result = self.run_cli(str(clean), "--rules", str(self.rules))
        self.assertEqual(result.returncode, 0)
        self.assertIn("PASS", result.stdout)

    def test_exit_2_on_usage_errors(self):
        self.assertEqual(self.run_cli("/nope.csv", "--rules", str(self.rules)).returncode, 2)
        self.assertEqual(self.run_cli(str(self.records), "--rules", "/nope.json").returncode, 2)

    def test_json_mode(self):
        result = self.run_cli(str(self.records), "--rules", str(self.rules), "--json")
        payload = json.loads(result.stdout)
        self.assertEqual(payload["records"], 2)
        self.assertEqual(payload["summary"]["unique"], 1)

    def test_audit_log_written(self):
        log = self.dir / "audit.jsonl"
        self.run_cli(str(self.records), "--rules", str(self.rules), "--audit-log", str(log))
        entry = json.loads(log.read_text(encoding="utf-8").strip())
        self.assertEqual(entry["result"], "FAIL")
        self.assertEqual(entry["file"], "records.csv")

    def test_invalid_regex_exits_two_not_one(self):
        bad = self.dir / "bad-rules.json"
        bad.write_text(json.dumps({"format": {"surname": "([a-z"}}), encoding="utf-8")
        result = self.run_cli(str(self.records), "--rules", str(bad))
        self.assertEqual(result.returncode, 2)
        self.assertIn("not a valid regular expression", result.stderr)

    def test_main_is_callable_in_process(self):
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            code = ra.main([str(self.records), "--rules", str(self.rules)])
        self.assertEqual(code, 1)
        self.assertIn("[required]", buffer.getvalue())


class TestZeroDependencies(unittest.TestCase):
    def test_only_standard_library_imports(self):
        tree = ast.parse(SCRIPT.read_text(encoding="utf-8"))
        names = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                names.update(alias.name.split(".")[0] for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
                names.add(node.module.split(".")[0])
        external = sorted(n for n in names if n not in sys.stdlib_module_names)
        self.assertEqual(external, [], f"third-party imports found: {external}")


if __name__ == "__main__":
    unittest.main()
