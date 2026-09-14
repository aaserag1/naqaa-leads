"""Focused public-API regression tests for the assembled app and its core.

Run against app.py once assembled. SIEVE_CORE_MODULE=core_segment allows a
core-only test run without starting or importing the main UI during development.
"""
from __future__ import annotations

import copy
import importlib
import importlib.util
import io
import json
import os
from pathlib import Path
import sys
import unittest
from unittest import mock
import zipfile

import openpyxl
import pandas as pd
import phonenumbers


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
MODULE = os.environ.get("SIEVE_CORE_MODULE") or ("app" if (ROOT / "app.py").exists() else "core_segment")
core = importlib.import_module(MODULE)

EXPECTED_COLUMNS = [
    "lead_id", "source", "source_row", "name", "phone_raw", "phone_e164", "phone_display",
    "phone_valid", "phone_region", "phone_error", "tel_url", "whatsapp_raw", "whatsapp_e164",
    "whatsapp_display", "whatsapp_valid", "whatsapp_error", "wa_url", "whatsapp_origin", "job",
    "budget_raw", "budget_value", "budget_currency", "vip", "timeline", "readiness", "city",
    "details", "original", "issues", "duplicate",
]
EXPECTED_KPIS = {"total", "valid_phone", "valid_wa", "vip_count", "ready_count", "ready_rate",
                 "reachable_count", "top_contact", "duplicates"}
EG_PHONE = "+201012345678"
EG_SECOND = "+201112345678"


def xlsx_bytes(rows, title="العملاء"):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = title
    for row in rows:
        sheet.append(row)
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def make_frame(rows=None):
    rows = rows if rows is not None else [
        ["أحمد", EG_PHONE, "", "مهندس", "١٢٠ ألف جنيه مصري", "جاهز الآن", "القاهرة", "ملاحظة", ""],
        ["", "not a phone", "explicit invalid", "", "100 USD", "مش دلوقتي", "", "", ""],
    ]
    return pd.DataFrame(rows, columns=["Name", "Phone", "WhatsApp", "Job", "Budget", "Timeline", "City", "Notes", "Empty"])


class ContractTests(unittest.TestCase):
    def test_constants_and_field_keys(self):
        self.assertEqual(core.MAX_FILE_BYTES, 20 * 1024 * 1024)
        self.assertEqual((core.MAX_FILES, core.MAX_ROWS, core.MAX_COLUMNS, core.MAX_CELLS, core.PDF_BATCH_SIZE),
                         (5, 25000, 200, 1500000, 500))
        self.assertEqual(set(core.FIELD_LABELS), {"name", "phone", "whatsapp", "job", "budget", "timeline", "city"})
        self.assertEqual(set(core.COUNTRY_OPTIONS), {"EG", "SA", "AE", "QA", "KW", "BH", "OM", "US", "GB", "JO", "LB", "IQ"})
        self.assertEqual(core.CURRENCY_OPTIONS, ["EGP", "SAR", "AED", "QAR", "KWD", "BHD", "OMR", "USD", "EUR"])
        self.assertTrue(issubclass(core.UserDataError, ValueError))

    def test_empty_clean_frame_and_kpis_are_complete(self):
        cleaned = core.clean_leads(pd.DataFrame(), {}, "فارغ")
        self.assertEqual(list(cleaned.columns), EXPECTED_COLUMNS)
        self.assertTrue(cleaned.empty)
        for column in ("phone_valid", "whatsapp_valid", "vip", "duplicate"):
            self.assertEqual(str(cleaned[column].dtype), "bool")
        kpis = core.calculate_kpis(cleaned)
        self.assertEqual(set(kpis), EXPECTED_KPIS)
        self.assertEqual(kpis["total"], 0)
        self.assertEqual(kpis["ready_rate"], 0)
        self.assertEqual(kpis["reachable_count"], 0)
        self.assertIsInstance(kpis["top_contact"], str)


class IngestionTests(unittest.TestCase):
    @unittest.skipUnless(importlib.util.find_spec("xlwt"), "يتطلب إنشاء عينة XLS مكتبة xlwt الاختبارية.")
    def test_xls_values_and_formula_cache_fail_closed(self):
        import xlwt
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Leads")
        for row, values in enumerate([["Name", "Phone"], ["Ali", EG_PHONE]]):
            for column, value in enumerate(values):
                sheet.write(row, column, value)
        output = io.BytesIO()
        workbook.save(output)
        self.assertEqual(core.workbook_sheets(output.getvalue(), "legacy.xls"), ["Leads"])
        source = core.read_source(output.getvalue(), "legacy.xls")
        self.assertEqual(source.rows[1], ["Ali", EG_PHONE])
        with mock.patch.object(core, "MAX_ROWS", 1):
            with self.assertRaises(core.UserDataError):
                core.read_source(output.getvalue(), "large.xls")
        sheet.write(1, 2, xlwt.Formula("100000+1"))
        output = io.BytesIO()
        workbook.save(output)
        with self.assertRaisesRegex(core.UserDataError, "XLSX"):
            core.read_source(output.getvalue(), "formula.xls")

    def test_sheet_names_and_csv_contract(self):
        data = xlsx_bytes([["Name", "Phone"], ["Ali", EG_PHONE]])
        self.assertEqual(core.workbook_sheets(data, "leads.xlsx"), ["العملاء"])
        self.assertEqual(core.workbook_sheets(b"Name\nAli\n", "leads.csv"), [""])
        with self.assertRaises(core.UserDataError):
            core.read_source(data, "leads.xlsx", sheet="غير موجود")

    def test_preamble_delimiter_source_rows_and_metadata(self):
        data = ("تقرير العملاء\nجهة التصدير\n\nالاسم;الهاتف;الميزانية\nأحمد;" + EG_PHONE
                + ";100000\n\nمنى;" + EG_SECOND + ";200000\n").encode("utf-8")
        source = core.read_source(data, "customers.csv")
        self.assertEqual(source.delimiter, ";")
        self.assertEqual(source.header_row, 3)
        retained = copy.deepcopy(source.rows)
        frame = core.table_from_source(source, source.header_row)
        self.assertEqual(frame.attrs["source_rows"], [5, 7])
        self.assertEqual(len(frame.attrs["metadata_rows"]), 3)
        self.assertEqual(source.rows, retained)
        self.assertEqual(len(frame), 2)

    def test_sep_directive_retained(self):
        source = core.read_source(b"sep=;\nName;Phone\nAli;+201012345678\n", "x.csv")
        self.assertEqual(source.delimiter, ";")
        self.assertEqual(source.header_row, 1)
        self.assertTrue(source.rows[0][0].startswith("sep="))
        self.assertEqual(core.table_from_source(source, 1).attrs["source_rows"], [3])

    def test_unicode_boms_and_strict_manual_encodings(self):
        text = "الاسم,الهاتف\nأحمد,٠١٠١٢٣٤٥٦٧٨\n"
        for encoding in ("utf-8-sig", "utf-16", "utf-32"):
            with self.subTest(encoding=encoding):
                source = core.read_source(text.encode(encoding), "unicode.csv")
                self.assertEqual(source.rows[1][0], "أحمد")
                self.assertEqual(source.rows[1][1], "٠١٠١٢٣٤٥٦٧٨")
        legacy_text = "الاسم,الهاتف\nأحمد,01012345678\n"
        source = core.read_source(legacy_text.encode("cp1256"), "arabic.csv", encoding="windows-1256")
        self.assertEqual(source.rows[1][0], "أحمد")
        self.assertEqual(source.encoding, "cp1256")
        with self.assertRaises(core.UserDataError):
            core.read_source(legacy_text.encode("cp1256"), "arabic.csv", encoding="utf-8")

    def test_auto_legacy_encoding_is_warned_and_not_replaced(self):
        arabic = "الاسم,الهاتف\nأحمد محمد,01012345678\n"
        source = core.read_source(arabic.encode("cp1256"), "legacy.csv")
        self.assertEqual(source.rows[1][0], "أحمد محمد")
        self.assertTrue(any("ترميز" in note or "الترميز" in note for note in source.warnings))
        latin = core.read_source("Name,City\nRené,Zürich\n".encode("cp1252"), "latin.csv", encoding="cp1252")
        self.assertEqual(latin.rows[1], ["René", "Zürich"])
        with self.assertRaises(core.UserDataError):
            core.read_source(b"Name,City\nA,\x81\n", "bad.csv", encoding="cp1252")

    def test_bad_signatures_password_container_and_corruption_rejected(self):
        for data, filename in ((b"not a zip", "a.xlsx"), (b"not ole", "a.xls"),
                               (b"PK\x03\x04junk", "a.xlsx"), (b"%PDF-1.7", "a.csv"),
                               (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1encrypted", "a.xlsx"),
                               (b"", "a.csv"), (b"abc", "a.xlsm")):
            with self.subTest(filename=filename, data=data[:8]):
                with self.assertRaises(core.UserDataError):
                    core.read_source(data, filename)

    def test_archive_ratio_member_count_and_xml_entities_rejected(self):
        base = xlsx_bytes([["Name", "Phone"], ["Ali", EG_PHONE]])
        compressed = io.BytesIO(base)
        with zipfile.ZipFile(compressed, "a", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("xl/padding.bin", b"0" * (2 * 1024 * 1024))
        with self.assertRaises(core.UserDataError):
            core.workbook_sheets(compressed.getvalue(), "bomb.xlsx")
        with mock.patch.object(core, "_ZIP_MAX_MEMBERS", 1):
            with self.assertRaises(core.UserDataError):
                core.workbook_sheets(base, "many.xlsx")
        malicious = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(base)) as original, zipfile.ZipFile(malicious, "w", zipfile.ZIP_DEFLATED) as result:
            for entry in original.infolist():
                payload = original.read(entry.filename)
                if entry.filename == "xl/workbook.xml":
                    payload = b'<!DOCTYPE workbook [<!ENTITY x "boom">]>' + payload
                result.writestr(entry.filename, payload)
        with self.assertRaises(core.UserDataError):
            core.workbook_sheets(malicious.getvalue(), "entities.xlsx")

    def test_rows_columns_cells_and_bytes_raise_instead_of_truncating(self):
        limits = [("MAX_ROWS", 2, b"Name,Phone\nA,1\nB,2\n"),
                  ("MAX_COLUMNS", 2, b"Name,Phone,City\nA,1,C\n"),
                  ("MAX_CELLS", 5, b"Name,Phone\nA,1\nB,2\n"),
                  ("MAX_FILE_BYTES", 5, b"Name\nAli\n")]
        for name, limit, data in limits:
            with self.subTest(limit=name), mock.patch.object(core, name, limit):
                with self.assertRaises(core.UserDataError):
                    core.read_source(data, "large.csv", delimiter=",")
        with mock.patch.object(core, "MAX_ROWS", 2):
            with self.assertRaises(core.UserDataError):
                core.read_source(xlsx_bytes([["Name"], ["A"], ["B"]]), "large.xlsx")

    def test_understated_xlsx_dimensions_do_not_hide_data(self):
        base = xlsx_bytes([["Name", "Phone"], ["Ali", EG_PHONE], ["Mona", EG_SECOND]])
        target = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(base)) as original, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
            for entry in original.infolist():
                payload = original.read(entry.filename)
                if entry.filename == "xl/worksheets/sheet1.xml":
                    payload = payload.replace(b'ref="A1:B3"', b'ref="A1:A1"')
                archive.writestr(entry.filename, payload)
        source = core.read_source(target.getvalue(), "dimensions.xlsx")
        self.assertEqual(len(source.rows), 3)
        self.assertEqual(source.rows[2][1], EG_SECOND)

    def test_malformed_csv_raises_and_quoted_multiline_is_retained(self):
        with self.assertRaises(core.UserDataError):
            core.read_source(b'Name,Phone\n"unclosed,+201012345678', "broken.csv", delimiter=",")
        source = core.read_source(b'Name,Notes\nAli,"first\nsecond"\n', "quoted.csv")
        self.assertEqual(source.rows[1][1], "first\nsecond")
        self.assertEqual(core.table_from_source(source, source.header_row).attrs["source_rows"], [2])

    def test_inconsistent_rows_duplicates_blanks_and_long_cells_preserved(self):
        long_value = "x" * 140000
        data = f"Name,Phone,Notes\nAli,{EG_PHONE}\nMona,{EG_SECOND},{long_value},extra\n,,,\n".encode()
        source = core.read_source(data, "inconsistent.csv", delimiter=",")
        self.assertEqual({len(row) for row in source.rows}, {4})
        frame = core.table_from_source(source, source.header_row)
        self.assertEqual(len(frame), 2)
        self.assertEqual(frame.iloc[1]["Notes"], long_value)
        self.assertEqual(frame.iloc[1, 3], "extra")
        self.assertTrue(source.warnings)
        manual = core.ParsedTable("test", "headers.csv", "", [["Name", "Name", "", "Name (2)", ""],
                                                                 ["A", "B", "C", "D", "E"]], "utf-8", ",", [], 0)
        unique = core.table_from_source(manual, 0)
        self.assertTrue(unique.columns.is_unique)
        self.assertEqual(len(unique.columns), 5)
        self.assertEqual(unique.iloc[0].tolist(), ["A", "B", "C", "D", "E"])

    def test_formulas_remain_strings_and_are_never_contact_numbers(self):
        data = xlsx_bytes([["Name", "Phone", "Budget"], ["Ali", "=10000000000+12345678", "=SUM(100000,200000)"]])
        source = core.read_source(data, "formula.xlsx")
        self.assertTrue(source.rows[1][1].startswith("="))
        frame = core.table_from_source(source, source.header_row)
        cleaned = core.clean_leads(frame, core.guess_mapping(frame), "صيغ")
        self.assertFalse(cleaned.iloc[0]["phone_valid"])
        self.assertFalse(cleaned.iloc[0]["vip"])
        self.assertEqual(cleaned.iloc[0]["phone_e164"], "")
        self.assertEqual(json.loads(cleaned.iloc[0]["original"])["Budget"], "=SUM(100000,200000)")


class HeaderMappingTests(unittest.TestCase):
    def test_scan_near_100th_row_and_unknown_file_never_loses_first_row(self):
        rows = [["بيانات تمهيدية"] for _ in range(97)] + [["الاسم", "رقم الهاتف"], ["أحمد", EG_PHONE]]
        index, confidence = core.detect_header(rows)
        self.assertEqual(index, 97)
        self.assertGreater(confidence, 0.8)
        source = core.read_source(b"Alice,12000\nBob,13000\n", "no-header.csv")
        self.assertEqual(source.header_row, -1)
        frame = core.table_from_source(source, -1)
        self.assertEqual(frame.iloc[0, 0], "Alice")
        self.assertEqual(frame.attrs["source_rows"], [1, 2])
        self.assertTrue(all(value is None for value in core.guess_mapping(frame).values()))

    def test_repeated_headers_are_records_with_warnings(self):
        source = core.read_source(f"Name,Phone\nAli,{EG_PHONE}\nName,Phone\nMona,{EG_SECOND}\n".encode(), "repeat.csv")
        frame = core.table_from_source(source, source.header_row)
        self.assertEqual(len(frame), 3)
        self.assertEqual(frame.attrs["source_rows"], [2, 3, 4])
        self.assertTrue(any("تكرّر" in warning for warning in source.warnings))
        clean = core.clean_leads(frame, core.guess_mapping(frame), "مكرر")
        self.assertEqual(len(clean), 3)
        self.assertIn("عناوين", clean.iloc[1]["issues"])

    def test_mapping_fuzzy_one_to_one_and_whatsapp_disambiguation(self):
        frame = pd.DataFrame(columns=["Full Name", "Phone Numbre", "WhatsApp Phone", "Job Title", "الميزانية", "جاهزية الشراء", "المدينة", "عمود 8"])
        mapping = core.guess_mapping(frame)
        self.assertEqual(mapping["name"], "Full Name")
        self.assertEqual(mapping["phone"], "Phone Numbre")
        self.assertEqual(mapping["whatsapp"], "WhatsApp Phone")
        self.assertEqual(mapping["timeline"], "جاهزية الشراء")
        assigned = [column for column in mapping.values() if column is not None]
        self.assertEqual(len(assigned), len(set(assigned)))
        self.assertNotIn("عمود 8", assigned)


class PhoneTests(unittest.TestCase):
    def test_explicit_country_wins_and_supported_regions_validate(self):
        for region in core.COUNTRY_OPTIONS:
            with self.subTest(region=region):
                example = phonenumbers.example_number_for_type(region, phonenumbers.PhoneNumberType.MOBILE)
                self.assertIsNotNone(example)
                international = phonenumbers.format_number(example, phonenumbers.PhoneNumberFormat.E164)
                result = core.normalize_phone(international, "EG")
                self.assertTrue(result["valid"], result)
                self.assertEqual(result["e164"], international)
                self.assertEqual(result["tel_url"], "tel:" + international)
                self.assertEqual(result["wa_url"], "https://wa.me/" + international[1:])
                self.assertEqual(core.normalize_phone("00" + international[1:])["e164"], international)

    def test_labels_arabic_digits_integral_artifacts_and_urls(self):
        examples = ["tel: +20 10 1234 5678", "p:٠١٠١٢٣٤٥٦٧٨", "۰۱۰۱۲۳۴۵۶۷۸", "01012345678.0",
                    "201012345678.0", "2.01012345678E+11", "https://wa.me/201012345678?text=Hello",
                    "wa.me/201012345678", "https://api.whatsapp.com/send?phone=201012345678"]
        for value in examples:
            with self.subTest(value=value):
                result = core.normalize_phone(value)
                self.assertTrue(result["valid"], result)
                self.assertEqual(result["e164"], EG_PHONE)
                self.assertEqual(result["raw"], value)
        self.assertTrue(core.normalize_phone("201012345678")["error"])
        self.assertTrue(core.normalize_phone("01012345678")["error"])

    def test_invalid_numbers_never_get_links_or_silently_concatenate(self):
        values = ["=201012345678", "@SUM(1,2)", "+SUM(1,2)", "call 01012345678 please", "201012345678.5",
                  "1.1e-3", "01012345678/01112345678", "01012345678,01112345678",
                  "01012345678 01112345678", "01012345678\n01112345678", "-01012345678",
                  "(+20 1012345678", "+9991012345678", "+201012345678 ext 123", "1e10000000",
                  "https://example.com/201012345678", "https://wa.me/201012345678/01112345678",
                  "https://api.whatsapp.com/send?phone=201012345678&phone=201112345678", ""]
        for value in values:
            with self.subTest(value=value):
                result = core.normalize_phone(value)
                self.assertFalse(result["valid"], result)
                self.assertEqual(result["e164"], "")
                self.assertEqual(result["tel_url"], "")
                self.assertEqual(result["wa_url"], "")
                self.assertTrue(result["error"])
                self.assertEqual(result["raw"], value)

    def test_ambiguous_gcc_is_not_guessed(self):
        # 8-digit mobile plans overlap among several GCC regions.
        local = "55123456"
        matching = []
        for region in ("SA", "AE", "QA", "KW", "BH", "OM"):
            try:
                number = phonenumbers.parse(local, region)
                if phonenumbers.is_valid_number(number):
                    matching.append(region)
            except phonenumbers.NumberParseException:
                pass
        self.assertGreater(len(matching), 1)
        result = core.normalize_phone(local, "EG")
        self.assertFalse(result["valid"])
        self.assertIn("أكثر من دولة", result["error"])
        self.assertEqual(result["wa_url"], "")

    def test_single_gcc_inference_is_explicitly_noted(self):
        # An ordinary 9-digit Saudi mobile is not valid as an Egyptian local.
        result = core.normalize_phone("512345678", "EG")
        self.assertTrue(result["valid"], result)
        self.assertEqual(result["region"], "SA")
        self.assertIn("أكّد", result["error"])


class BudgetReadinessTests(unittest.TestCase):
    def test_budget_digits_scales_currencies_and_range_lower_bound(self):
        examples = {
            "١٢٠ ألف جنيه مصري": (120000, "EGP"), "۲٫۵ مليون ريال سعودي": (2500000, "SAR"),
            "100-200k EGP": (100000, "EGP"), "100k-200k": (100000, "EGP"),
            "من ١٠٠ إلى ٢٠٠ ألف": (100000, "EGP"), "between 100 and 200 thousand USD": (100000, "USD"),
            "200000-100000": (100000, "EGP"), "AED 150,000": (150000, "AED"),
            "€1,200.50": (1200.5, "EUR"), "١٢٣٬٤٥٦٫٧٨ ج.م": (123456.78, "EGP"),
            "1.5m SAR": (1500000, "SAR"), "100k+": (100000, "EGP"), "0": (0, "EGP"),
            "1e5": (100000, "EGP"), "100 000": (100000, "EGP"), "مليون ونصف": (1500000, "EGP"),
        }
        for value, expected in examples.items():
            with self.subTest(value=value):
                actual = core.parse_budget(value)
                self.assertEqual(actual, expected)

    def test_unknown_negative_or_ambiguous_budget_returns_none(self):
        for value in ("", "غير معروف", "حسب الاتفاق", "-50000", "−٥٠٠", "(1000)", "100 USD / 200 EGP",
                      "100k أو 200k", "100 / 200", "1.000", "100-200-300", "100k-200", "=1e10",
                      "NaN", "Infinity", "1e1000", "100%", "100- -200", "5,00,000", "100 GBP", "100 ريال",
                      "100 USD 200", "100 جنيه 200", "100 EGP 200 EGP"):
            with self.subTest(value=value):
                self.assertIsNone(core.parse_budget(value)[0])
        self.assertEqual(core.parse_budget("100 ريال", "SAR"), (100, "SAR"))
        self.assertEqual(core.parse_budget("100 USD", "EGP"), (100, "USD"))

    def test_readiness_negation_before_positive_and_later_before_yes(self):
        for value in ("not now", "not ready", "not ready now", "غير جاهز", "مش دلوقتي", "مو جاهز", "لست مستعدًا",
                      "no, not ready", "yes, next month", "جاهز بعد شهر", "لا", "not today", "never ready", "مش عايز دلوقتي"):
            with self.subTest(value=value):
                self.assertEqual(core.classify_readiness(value), "لاحقًا")
        for value in ("نعم", "جاهز الآن", "yes", "immediately", "ready", "فوري"):
            self.assertEqual(core.classify_readiness(value), "فوري")
        for value in ("", "ربما", "unknown", "notebook", "=TRUE()"):
            self.assertEqual(core.classify_readiness(value), "غير محدد")


class CleaningDuplicateKpiTests(unittest.TestCase):
    def test_index_preserving_filter_and_sort_keep_original_row_numbers(self):
        frame = make_frame()
        frame.attrs["source_rows"] = [6, 9]
        reversed_frame = frame.iloc[::-1].copy()
        cleaned = core.clean_leads(reversed_frame, core.guess_mapping(frame), "ordered")
        self.assertEqual(cleaned["source_row"].tolist(), [9, 6])
        subset = core.clean_leads(frame.iloc[[1]].copy(), core.guess_mapping(frame), "subset")
        self.assertEqual(subset["source_row"].tolist(), [9])

    def test_zero_column_rows_are_not_silently_dropped(self):
        cleaned = core.clean_leads(pd.DataFrame(index=range(2)), {}, "empty columns")
        self.assertEqual(len(cleaned), 2)
        self.assertEqual(cleaned["name"].tolist(), ["بدون اسم", "بدون اسم"])
        self.assertEqual(cleaned["source_row"].tolist(), [1, 2])
        self.assertEqual(cleaned["original"].tolist(), ["{}", "{}"])

    def test_clean_schema_original_details_and_deterministic_identity(self):
        frame = make_frame()
        frame.attrs["source_rows"] = [6, 9]
        frame.attrs["source_id"] = "source-A"
        original = frame.copy(deep=True)
        mapping = core.guess_mapping(frame)
        cleaned = core.clean_leads(frame, mapping, "ملف العملاء")
        self.assertEqual(list(cleaned.columns), EXPECTED_COLUMNS)
        self.assertEqual(cleaned["source_row"].tolist(), [6, 9])
        self.assertEqual(cleaned.iloc[1]["name"], "بدون اسم")
        self.assertEqual(cleaned.iloc[0]["budget_value"], 120000)
        self.assertEqual(json.loads(cleaned.iloc[0]["details"]), {"Notes": "ملاحظة", "Empty": ""})
        raw = json.loads(cleaned.iloc[0]["original"])
        self.assertEqual(set(raw), set(frame.columns))
        self.assertEqual(raw["WhatsApp"], "")
        self.assertEqual(raw["Empty"], "")
        self.assertFalse(cleaned["duplicate"].any())
        self.assertEqual(cleaned["lead_id"].tolist(), core.clean_leads(frame, mapping, "ملف العملاء")["lead_id"].tolist())
        self.assertEqual(cleaned["lead_id"].nunique(), 2)
        other = core.clean_leads(frame, mapping, "ملف آخر")
        self.assertTrue(set(cleaned["lead_id"]).isdisjoint(set(other["lead_id"])))
        pd.testing.assert_frame_equal(frame, original)

    def test_whatsapp_fallback_only_empty_explicit_opt_in_and_raw_retention(self):
        frame = pd.DataFrame({"Name": ["A", "B"], "Phone": [EG_PHONE, EG_PHONE], "WhatsApp": ["", "invalid"]})
        mapping = core.guess_mapping(frame)
        disabled = core.clean_leads(frame, mapping, "x", wa_fallback=False)
        enabled = core.clean_leads(frame, mapping, "x", wa_fallback=True)
        self.assertEqual(disabled["whatsapp_valid"].tolist(), [False, False])
        self.assertEqual(enabled["whatsapp_valid"].tolist(), [True, False])
        self.assertEqual(enabled.iloc[0]["whatsapp_e164"], EG_PHONE)
        self.assertEqual(enabled.iloc[0]["whatsapp_raw"], "")
        self.assertEqual(json.loads(enabled.iloc[0]["original"])["WhatsApp"], "")
        self.assertEqual(enabled.iloc[1]["whatsapp_raw"], "invalid")
        self.assertEqual(enabled.iloc[1]["wa_url"], "")
        self.assertIn("الهاتف", enabled.iloc[0]["whatsapp_origin"])

    def test_vip_currency_scope_and_conservative_range(self):
        frame = pd.DataFrame({"Budget": ["150000 EGP", "150000 USD", "90000-200000 EGP", "-200000", "unknown"]})
        clean = core.clean_leads(frame, {"budget": "Budget"}, "budget", vip_threshold=100000, default_currency="EGP")
        self.assertEqual(clean["vip"].tolist(), [True, False, False, False, False])
        self.assertEqual(clean.iloc[1]["budget_currency"], "USD")
        self.assertIsNone(clean.iloc[3]["budget_value"])
        with self.assertRaises(core.UserDataError):
            core.clean_leads(frame, {"budget": "Budget"}, "budget", vip_threshold=float("nan"))

    def test_no_name_no_mapping_and_all_raw_data_are_preserved(self):
        long_value = "ط" * 180000
        frame = pd.DataFrame({"Unknown": [long_value, "other"], "Empty": ["", ""]})
        clean = core.clean_leads(frame, {}, "raw")
        self.assertEqual(len(clean), 2)
        self.assertEqual(clean["name"].tolist(), ["بدون اسم", "بدون اسم"])
        self.assertEqual(json.loads(clean.iloc[0]["original"])["Unknown"], long_value)
        self.assertEqual(json.loads(clean.iloc[0]["details"])["Unknown"], long_value)

    def test_duplicate_all_rows_cross_field_self_and_invalid_exclusions(self):
        frame = pd.DataFrame({"Phone": [EG_PHONE, EG_SECOND, "", "", "invalid", "invalid"],
                              "WhatsApp": [EG_PHONE, EG_PHONE, EG_SECOND, "+447400123456", "", ""]})
        clean = core.clean_leads(frame, core.guess_mapping(frame), "duplicates")
        marked = core.mark_duplicates(clean)
        self.assertEqual(marked["duplicate"].tolist(), [True, True, True, False, False, False])
        self.assertFalse(clean["duplicate"].any())
        self.assertEqual(len(marked), len(clean))
        isolated = core.mark_duplicates(clean.iloc[[0]].copy())
        self.assertEqual(isolated["duplicate"].tolist(), [False])
        pd.testing.assert_frame_equal(marked, core.mark_duplicates(marked))

    def test_kpis_reachable_union_duplicates_and_percentage(self):
        frame = pd.DataFrame({"phone_valid": [True, True, False, False],
                              "whatsapp_valid": [True, False, True, False],
                              "vip": [True, False, False, True], "duplicate": [True, True, False, False],
                              "readiness": ["فوري", "لاحقًا", "فوري", "غير محدد"]})
        result = core.calculate_kpis(frame)
        self.assertEqual(set(result), EXPECTED_KPIS)
        self.assertEqual(result["valid_phone"], 2)
        self.assertEqual(result["valid_wa"], 2)
        self.assertEqual(result["reachable_count"], 3)
        self.assertEqual(result["vip_count"], 2)
        self.assertEqual(result["ready_count"], 2)
        self.assertEqual(result["ready_rate"], 50)
        self.assertEqual(result["duplicates"], 2)

    def test_invalid_mapping_and_source_provenance_are_explicit_errors(self):
        frame = make_frame()
        for mapping in ({"phone": "missing"}, {"phone": "Phone", "whatsapp": "Phone"}):
            with self.assertRaises(core.UserDataError):
                core.clean_leads(frame, mapping, "x")
        frame.attrs["source_rows"] = [2, 2]
        with self.assertRaises(core.UserDataError):
            core.clean_leads(frame, core.guess_mapping(frame), "x")


if __name__ == "__main__":
    unittest.main()
