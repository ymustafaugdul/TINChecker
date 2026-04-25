from __future__ import annotations

import unittest

from tin_checker_core import load_countries, validate_bulk_entries, validate_many, validate_tin
from tin_checker_web import create_app


class TinCheckerCoreTests(unittest.TestCase):
    def test_all_listed_countries_validate_without_runtime_errors(self) -> None:
        countries = load_countries("en")
        self.assertGreater(len(countries), 100)

        for country in countries:
            with self.subTest(country=country["code"]):
                result = validate_tin(country["code"], "1")
                self.assertEqual(result["countryCode"], country["code"])
                self.assertIn("isValid", result)

    def test_common_input_shapes_do_not_hit_porting_errors(self) -> None:
        samples = [
            "1",
            "123456789",
            "1234567890",
            "12345678901",
            "1234567890123",
            "AB123456C",
            "A123456789",
            "12-3456789",
        ]

        for country in load_countries("en"):
            for sample in samples:
                with self.subTest(country=country["code"], sample=sample):
                    result = validate_tin(country["code"], sample)
                    self.assertFalse(str(result.get("errorMsg", "")).startswith("Validation rule failed safely"))

    def test_missing_input_returns_controlled_error(self) -> None:
        result = validate_tin("DE", "")

        self.assertFalse(result["isValid"])
        self.assertEqual(result["errorMsg"], "TIN is required.")

    def test_turkish_and_english_localization_are_available(self) -> None:
        tr_countries = load_countries("tr")
        en_countries = load_countries("en")

        self.assertEqual(next(item for item in tr_countries if item["code"] == "GB")["label"], "Birleşik Krallık")
        self.assertEqual(next(item for item in en_countries if item["code"] == "GB")["label"], "United Kingdom")
        self.assertEqual(validate_tin("DE", "", "tr")["errorMsg"], "TIN zorunludur.")
        self.assertEqual(validate_tin("DE", "", "en")["errorMsg"], "TIN is required.")

    def test_bulk_summary_counts_results(self) -> None:
        payload = validate_many("US", ["123456789", "", "12-3456789"])

        self.assertEqual(payload["summary"]["total"], 2)
        self.assertEqual(payload["summary"]["valid"] + payload["summary"]["invalid"], 2)

    def test_bulk_entries_support_multiple_countries(self) -> None:
        payload = validate_bulk_entries([
            {"countryCode": "DE", "values": ["123456789", "12345678901"]},
            {"countryCode": "US", "values": ["123456789"]},
        ], "en")

        self.assertEqual(payload["summary"]["total"], 3)
        self.assertEqual(len(payload["groups"]), 2)
        self.assertEqual(payload["groups"][0]["countryCode"], "DE")

    def test_web_app_serves_interface_and_validation_api(self) -> None:
        client = create_app().test_client()

        page = client.get("/")
        self.assertEqual(page.status_code, 200)
        html = page.data.decode("utf-8")
        self.assertIn("Yabancı VKN Kontrolcüsü", html)
        self.assertIn("Made by", html)
        self.assertIn("Mustafa ÜĞDÜL", html)
        self.assertIn("2026", html)
        self.assertNotIn("Local validation screen that runs without Excel", html)
        response = client.post("/api/validate/single", json={"lang": "en", "countryCode": "DE", "tin": "123456789"})

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json["countryCode"], "DE")


if __name__ == "__main__":
    unittest.main()
