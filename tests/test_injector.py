import builtins
import contextlib
import hashlib
import importlib.util
import io
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET

import pandas as pd

import excelxml


GTIN = "01234567890123"


def markings(rows=None):
    return pd.DataFrame(rows if rows is not None else [(GTIN, "SYNTHETIC-CODE")],
                        columns=["GTIN", "Код маркировки"])


def reference(rows=None):
    return pd.DataFrame(rows if rows is not None else [("00042", GTIN)],
                        columns=["Баркод", "ГТИН"])


def document(barcodes=("00042",)):
    root = ET.Element("Файл")
    for barcode in barcodes:
        product = ET.SubElement(root, "СведТов")
        if barcode is not None:
            ET.SubElement(product, "ДопСведТов", КодТов=barcode)
    return root


class InjectionTests(unittest.TestCase):
    def test_reserved_text_and_literal_entity_roundtrip(self):
        values = ['synthetic&<value>"quote\'apostrophe', "literal&amp;text"]
        root = document()
        counts = excelxml.inject_markings(root, markings([(GTIN, v) for v in values]), reference())
        reparsed = ET.fromstring(ET.tostring(root, encoding="windows-1251"))
        self.assertEqual([e.text for e in reparsed.findall(".//КИЗ")], values)
        self.assertEqual(counts["added_codes"], 2)

    def test_leading_zero_and_trim_before_padding(self):
        root = document((" 00042 ",))
        counts = excelxml.inject_markings(root, markings([(" 1234567890123 ", "code")]),
                                         reference([("00042", GTIN)]))
        self.assertEqual(counts["added_codes"], 1)
        self.assertEqual(excelxml.normalize_gtin(" 1234567890123 "), GTIN)
        self.assertEqual(root.find(".//КИЗ").text, "code")

    def test_consistent_reference_and_marking_duplicates_deduplicate(self):
        root = document()
        counts = excelxml.inject_markings(root, markings([(GTIN, "one"), (GTIN, "one"), (GTIN, "two")]),
                                         reference([("00042", None), ("00042", GTIN), ("00042", GTIN)]))
        self.assertEqual(counts["added_codes"], 2)

    def test_conflicting_reference_fails_before_mutation(self):
        root = document()
        before = ET.tostring(root)
        with self.assertRaisesRegex(ValueError, "multiple GTINs"):
            excelxml.inject_markings(root, markings(), reference([("00042", GTIN), ("00042", "123")]))
        self.assertEqual(ET.tostring(root), before)

    def test_same_code_across_gtins_rejected(self):
        with self.assertRaisesRegex(ValueError, "multiple GTINs"):
            excelxml.inject_markings(document(), markings([(GTIN, "one"), ("123", "one")]), reference())

    def test_repeated_run_is_idempotent(self):
        root = document()
        excelxml.inject_markings(root, markings(), reference())
        counts = excelxml.inject_markings(root, markings(), reference())
        self.assertEqual(counts["added_codes"], 0)
        self.assertEqual(counts["existing_codes"], 1)
        self.assertEqual(len(root.findall(".//КИЗ")), 1)

    def test_missing_mappings_have_separate_counts(self):
        root = document((None, "", "unknown", "empty-gtin", "no-codes", "00042"))
        counts = excelxml.inject_markings(root, markings(), reference([
            ("empty-gtin", None), ("no-codes", "123"), ("00042", GTIN)]))
        self.assertEqual(counts["products"], 6)
        for key in ("missing_additional_info", "missing_barcode", "missing_reference",
                    "missing_reference_gtin", "missing_markings", "added_codes", "matched_products"):
            self.assertEqual(counts[key], 1, key)

    def test_repeated_xml_gtin_fails_before_any_append(self):
        root = document(("00042", "00042"))
        before = ET.tostring(root)
        with self.assertRaisesRegex(ValueError, "allocation is ambiguous"):
            excelxml.inject_markings(root, markings(), reference())
        self.assertEqual(ET.tostring(root), before)

    def test_existing_code_on_other_product_rejected_without_mutation(self):
        root = document(("00042", "unmapped"))
        detail = root.findall(".//ДопСведТов")[1]
        ET.SubElement(ET.SubElement(detail, "НомСредИдентТов"), "КИЗ").text = "SYNTHETIC-CODE"
        before = ET.tostring(root)
        with self.assertRaisesRegex(ValueError, "another XML product"):
            excelxml.inject_markings(root, markings(), reference())
        self.assertEqual(ET.tostring(root), before)

    def test_invalid_gtin_missing_code_and_xml_control_rejected(self):
        for bad_gtin, bad_code in [("1.0", "code"), ("1" * 15, "code"),
                                   (GTIN, None), (GTIN, "synthetic\x1dcode")]:
            with self.subTest(gtin=bad_gtin, code=repr(bad_code)):
                with self.assertRaises(ValueError):
                    excelxml.inject_markings(document(), markings([(bad_gtin, bad_code)]), reference())

    def test_namespaced_or_unsupported_document_rejected(self):
        with self.assertRaisesRegex(ValueError, "non-namespaced"):
            excelxml.inject_markings(ET.fromstring('<Файл xmlns="urn:synthetic"><СведТов/></Файл>'),
                                     markings(), reference())

    def test_carriage_return_is_rejected_instead_of_silently_normalized(self):
        with self.assertRaisesRegex(ValueError, "carriage return"):
            excelxml.inject_markings(document(), markings([(GTIN, "synthetic\rcode")]), reference())

    def test_import_is_inert(self):
        spec = importlib.util.spec_from_file_location("import_check", Path(excelxml.__file__))
        module = importlib.util.module_from_spec(spec)
        original_import = builtins.__import__
        def import_without_gui(name, *args, **kwargs):
            if name == "tkinter" or name.startswith("tkinter."):
                raise AssertionError("GUI import forbidden")
            return original_import(name, *args, **kwargs)
        with patch("builtins.__import__", side_effect=import_without_gui), \
             patch("pandas.read_excel", side_effect=AssertionError("XLSX I/O forbidden")), \
             patch("xml.etree.ElementTree.parse", side_effect=AssertionError("XML I/O forbidden")):
            spec.loader.exec_module(module)


class SyntheticFileTests(unittest.TestCase):
    def setUp(self):
        # Keep openpyxl's own temporary files inside this synthetic test directory.
        work = Path(__file__).resolve().parents[1] / "work"
        work.mkdir(exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(prefix="synthetic-", dir=work)
        self.folder = Path(self.temp.name)
        self.previous_tempdir = tempfile.tempdir
        tempfile.tempdir = self.temp.name
        self.addCleanup(self.restore_temp)
        self.markings = self.folder / "markings.xlsx"
        self.reference = self.folder / "reference.xlsx"
        self.xml = self.folder / "input.xml"
        self.out = self.folder / "output.xml"
        markings([(GTIN, 'SYNTHETIC&<code>"\'')]).to_excel(self.markings, index=False)
        reference().to_excel(self.reference, index=False)
        ET.ElementTree(document()).write(self.xml, encoding="windows-1251", xml_declaration=True)

    def restore_temp(self):
        tempfile.tempdir = self.previous_tempdir
        self.temp.cleanup()

    def paths(self, out=None):
        return [str(p) for p in (self.markings, self.reference, self.xml, out or self.out)]

    def test_real_xlsx_xml_main_roundtrip_inputs_unchanged_and_no_dialogs(self):
        originals = {p: hashlib.sha256(p.read_bytes()).hexdigest()
                     for p in (self.markings, self.reference, self.xml)}
        args = []
        for flag, path in zip(("--markings", "--reference", "--xml", "--out"), self.paths()):
            args.extend((flag, path))
        captured = io.StringIO()
        with patch.object(excelxml, "select_file", side_effect=AssertionError("GUI forbidden")), \
             contextlib.redirect_stdout(captured):
            self.assertEqual(excelxml.main(args), 0)
        self.assertEqual(ET.parse(self.out).find(".//КИЗ").text, 'SYNTHETIC&<code>"\'')
        self.assertIn(b"windows-1251", self.out.read_bytes())
        self.assertNotIn("SYNTHETIC", captured.getvalue())
        for path, digest in originals.items():
            self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), digest)

    def test_existing_output_requires_explicit_overwrite(self):
        self.out.write_bytes(b"SYNTHETIC PREEXISTING")
        with self.assertRaises(FileExistsError):
            excelxml.process_files(*self.paths())
        self.assertEqual(self.out.read_bytes(), b"SYNTHETIC PREEXISTING")
        self.assertEqual(excelxml.process_files(*self.paths(), overwrite=True)["added_codes"], 1)

    def test_literal_na_text_survives_xlsx_import(self):
        markings([(GTIN, "NA")]).to_excel(self.markings, index=False)
        excelxml.process_files(*self.paths())
        self.assertEqual(ET.parse(self.out).find(".//КИЗ").text, "NA")

    def test_cannot_overwrite_an_input_even_explicitly(self):
        before = self.xml.read_bytes()
        with self.assertRaisesRegex(ValueError, "separate"):
            excelxml.process_files(*self.paths(self.xml), overwrite=True)
        self.assertEqual(self.xml.read_bytes(), before)

    def test_cancel_does_not_read_any_files(self):
        with patch.object(excelxml, "select_file", return_value=""), \
             patch.object(excelxml, "process_files", side_effect=AssertionError("No file I/O")), \
             contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(excelxml.main([]), 2)


if __name__ == "__main__":
    unittest.main()
