import json
import re
import unittest
from pathlib import Path
from urllib.parse import urlparse

from gmpai.catalog import CatalogError, load_catalog

OFFICIAL_HOSTS = {
    "health.ec.europa.eu",
    "www.ema.europa.eu",
    "eur-lex.europa.eu",
    "www.fda.gov",
}


class CatalogTests(unittest.TestCase):
    def setUp(self):
        self.catalog = load_catalog()

    def test_catalog_loads_and_has_both_authorities(self):
        self.assertGreaterEqual(len(self.catalog), 10)
        self.assertEqual(self.catalog.authorities, ["EU", "FDA"])

    def test_every_url_points_at_an_official_regulator_host(self):
        for doc in self.catalog:
            for url in filter(None, (doc.landing_page, doc.direct_url)):
                host = urlparse(url).netloc
                self.assertIn(host, OFFICIAL_HOSTS, f"{doc.id} uses non-official host {host}")
                self.assertTrue(url.startswith("https://"), f"{doc.id} must use HTTPS")

    def test_filenames_are_unique_and_pdf(self):
        names = [d.filename for d in self.catalog]
        self.assertEqual(len(names), len(set(names)))
        for name in names:
            self.assertTrue(name.endswith(".pdf"), name)
            self.assertNotIn("/", name)

    def test_discover_patterns_compile(self):
        for doc in self.catalog:
            if doc.discover_pattern:
                re.compile(doc.discover_pattern)

    def test_filter_by_authority_and_id(self):
        eu = self.catalog.filter(authority="eu")
        self.assertTrue(eu)
        self.assertTrue(all(d.authority == "EU" for d in eu))

        one = self.catalog.filter(ids=["eu-gmp-annex22-ai"])
        self.assertEqual(len(one), 1)
        self.assertEqual(one[0].id, "eu-gmp-annex22-ai")

    def test_unknown_id_is_rejected(self):
        with self.assertRaises(CatalogError):
            self.catalog.filter(ids=["does-not-exist"])

    def test_annex22_and_fda_ai_guidance_are_present(self):
        ids = {d.id for d in self.catalog}
        self.assertIn("eu-gmp-annex22-ai", ids)
        self.assertIn("fda-ai-regulatory-decision-making", ids)

    def test_document_without_url_or_pattern_is_rejected(self):
        tmp = Path(self.id() + ".json")
        tmp.write_text(
            json.dumps(
                {
                    "documents": [
                        {
                            "id": "broken",
                            "title": "t",
                            "authority": "EU",
                            "category": "gmp-ai",
                            "status": "draft",
                            "filename": "t.pdf",
                            "landing_page": "https://health.ec.europa.eu/",
                        }
                    ]
                }
            ),
            encoding="utf-8",
        )
        self.addCleanup(tmp.unlink)
        with self.assertRaises(CatalogError):
            load_catalog(tmp)


if __name__ == "__main__":
    unittest.main()
