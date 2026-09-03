import importlib.util
import os
import unittest


SCRIPT_PATH = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    "google-business-profiles-cloud-seo-scraper.py",
)

spec = importlib.util.spec_from_file_location("gbp_scraper", SCRIPT_PATH)
scraper = importlib.util.module_from_spec(spec)
spec.loader.exec_module(scraper)


class ScraperUtilityTests(unittest.TestCase):
    def test_env_flag_true_values(self):
        os.environ["TEST_FLAG"] = "sim"
        self.assertTrue(scraper.env_flag("TEST_FLAG", False))

    def test_env_flag_false_values(self):
        os.environ["TEST_FLAG"] = "não"
        self.assertFalse(scraper.env_flag("TEST_FLAG", True))

    def test_classify_own_website(self):
        self.assertEqual(
            scraper.classify_website("https://example.com"),
            ("Site próprio", True),
        )

    def test_classify_social_website(self):
        self.assertEqual(
            scraper.classify_website("https://www.instagram.com/example"),
            ("Rede social", False),
        )

    def test_classify_delivery_website(self):
        self.assertEqual(
            scraper.classify_website("https://www.ifood.com/example"),
            ("App delivery", False),
        )

    def test_whatsapp_link_from_brazilian_phone(self):
        self.assertEqual(
            scraper.normalize_phone_to_whatsapp_link("(24) 99999-9999"),
            "https://wa.me/5524999999999",
        )

    def test_domain_root_removes_www(self):
        self.assertEqual(
            scraper.get_domain_root("https://www.example.com/path"),
            "example.com",
        )

    def test_sheet_title_is_sanitized_and_limited(self):
        title = scraper.sanitize_sheet_title("A/B:C*D?E[FG] " + "x" * 40)
        self.assertNotRegex(title, r"[\[\]\*/:\?]")
        self.assertLessEqual(len(title), 31)

    def test_extract_city_state_from_search_fallback(self):
        self.assertEqual(
            scraper.extract_city_state(None, "Petrópolis, RJ"),
            ("Petrópolis", "RJ"),
        )


if __name__ == "__main__":
    unittest.main()
