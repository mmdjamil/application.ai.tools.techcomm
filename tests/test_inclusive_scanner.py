import unittest

from scanners.inclusive_scanner import (
    apply_all_accepted_replacements,
    apply_replacement_to_text,
    pick_default_replacement,
    scan_for_inclusive_language,
)


class InclusiveScannerTests(unittest.TestCase):
    def test_pick_default_replacement_uses_first_option(self):
        self.assertEqual(pick_default_replacement("primary / initiator"), "primary")
        self.assertEqual(pick_default_replacement("alpha, beta"), "alpha")
        self.assertEqual(pick_default_replacement("denylist"), "denylist")

    def test_apply_replacement_to_text_preserves_case(self):
        text = "Blacklist BLACKLIST blacklist"
        replaced = apply_replacement_to_text(text, "blacklist", "denylist")
        self.assertEqual(replaced, "Denylist DENYLIST denylist")

    def test_apply_all_accepted_replacements_handles_multiple_terms(self):
        text = "The master controls the slave."
        accepted = [
            ("master", "primary / initiator"),
            ("slave", "secondary / target"),
        ]
        replaced = apply_all_accepted_replacements(text, accepted)
        self.assertEqual(replaced, "The primary controls the secondary.")

    def test_scan_returns_term_and_sentences(self):
        parsed_lines = [
            {"page": 1, "line": 1, "text": "Blacklist this item and whitelist that item."},
            {"page": 1, "line": 2, "text": "No issue here."},
        ]
        results = scan_for_inclusive_language(parsed_lines)
        self.assertEqual(results["total_non_inclusive_count"], 2)
        first = results["findings"][0]
        self.assertIn("term", first)
        self.assertIn("original_sentence", first)
        self.assertIn("suggested_sentence", first)
        self.assertEqual(first["original_sentence"], "Blacklist this item and whitelist that item.")
        self.assertIn("Denylist", first["suggested_sentence"])


if __name__ == "__main__":
    unittest.main()
