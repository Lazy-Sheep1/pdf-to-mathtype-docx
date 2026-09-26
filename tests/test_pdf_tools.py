import importlib.util
from pathlib import Path
import tempfile
import unittest

import fitz

ROOT = Path(__file__).resolve().parents[1]


def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


inventory = load("pdf_inventory")
integrity = load("text_integrity")


class PDFToolsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.pdf = self.root / "synthetic.pdf"
        with fitz.open() as doc:
            page = doc.new_page(width=400, height=300)
            page.insert_text((30, 50), "A well-known method [23].")
            doc.new_page(width=400, height=300)
            doc.save(self.pdf)

    def tearDown(self):
        self.tmp.cleanup()

    def manifest(self, text):
        return {"entries": [{"id": "p1", "text": text,
                             "regions": [{"page": 1, "bbox": [20, 20, 390, 100]}]}]}

    def test_inventory_text_and_blank_page(self):
        report = inventory.inventory(self.pdf, self.root / "inventory", dpi=36)
        self.assertEqual(report["page_count"], 2)
        self.assertEqual(len(report["source_sha256"]), 64)
        self.assertFalse(report["pages"][0]["needs_ocr_review"])
        self.assertTrue(report["pages"][1]["needs_ocr_review"])
        self.assertTrue((self.root / "inventory" / "page-001.png").exists())
        with self.assertRaises(FileExistsError):
            inventory.inventory(self.pdf, self.root / "inventory", dpi=0)

    def test_complete_region(self):
        self.assertTrue(integrity.audit(self.pdf, self.manifest("A well-known method [23]."))["pass"])

    def test_punctuation_number_and_true_hyphen_are_not_erased(self):
        for bad in ("A well-known method [23]", "A well-known method [2].", "A wellknown method [23]."):
            self.assertFalse(integrity.audit(self.pdf, self.manifest(bad))["pass"])

    def test_normalization_preserves_math_alphabet_and_digits(self):
        self.assertEqual(integrity.normalize("of\ufb01ce \u00ad[12]."), "office[12].")
        self.assertNotEqual(integrity.normalize("\U0001d465"), integrity.normalize("x"))
        self.assertNotEqual(integrity.normalize("well-known"), integrity.normalize("wellknown"))

    def test_explicit_discretionary_break(self):
        self.assertEqual(integrity.remove_approved_breaks("para-\ngraph", ["para-\ngraph"]), "paragraph")
        with self.assertRaises(ValueError):
            integrity.remove_approved_breaks("paragraph", ["para-\ngraph"])

    def test_invalid_manifest_is_not_a_pass(self):
        with self.assertRaises(ValueError):
            integrity.audit(self.pdf, {"entries": []})
        data = self.manifest("text")
        data["entries"][0]["regions"][0]["bbox"] = [0, 0, 450, 100]
        with self.assertRaises(ValueError):
            integrity.audit(self.pdf, data)


if __name__ == "__main__":
    unittest.main()
