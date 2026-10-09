import base64
import html
import json
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from offer_automation import (
    _retailer_display_name,
    _search_terms,
    _fit_image_dimensions,
    categorize_offer,
    _filter_offers,
    Offer,
    Publication,
    CoopCoverParser,
    JsonLdParser,
    _overlaps_week,
    parse_publications,
    parse_search_results,
    select_week_publication,
)


class OfferParsingTests(unittest.TestCase):
    def test_parses_offer_from_json_ld(self):
        html = '''<script type="application/ld+json">
        {"@type":"SearchResultsPage","mainEntity":{"@type":"ItemList",
        "itemListElement":[{"item":{"@type":"Offer","name":"Head & Shoulders Shampoo",
        "url":"https://etilbudsavis.dk/Netto?publication=week42&offer=1","price":20,
        "image":"https://images.example/offer.webp","seller":{"@type":"Organization",
        "name":"Netto"},"validThrough":"2099-01-01T00:00:00+0000"}}]}}
        </script>'''
        offers = parse_search_results(html)
        self.assertEqual(len(offers), 1)
        self.assertEqual(offers[0].retailer, "Netto")
        self.assertEqual(offers[0].image_url, "https://images.example/offer.webp")
        self.assertEqual(offers[0].publication_id, "week42")

    def test_parses_full_embedded_publication_offer_list(self):
        key = base64.b64encode(json.dumps([
            "offers",
            {"sources": ["publication"], "pagination": {"limit": 30, "offset": 0}},
        ]).encode()).decode()
        payload = {"data": [{
            "publicId": "offer42",
            "name": "Ansigtscreme",
            "image": "https://images.example/small.webp",
            "imageLarge": "https://images.example/large.webp",
            "price": 49,
            "description": "Produktbeskrivelse",
            "validUntil": "2026-10-16T21:59:59+0000",
            "publicationPublicId": "week42",
            "business": {"name": "Netto", "slugs": ["Netto"]},
        }]}
        page = f'<app-data data-key="{key}">{html.escape(json.dumps(payload))}</app-data>'
        offers = parse_search_results(page)
        self.assertEqual(len(offers), 1)
        self.assertEqual(offers[0].publication_id, "week42")
        self.assertEqual(offers[0].image_url, "https://images.example/large.webp")
        self.assertIn("offer=offer42", offers[0].url)
        self.assertEqual(offers[0].description, "Produktbeskrivelse")

    def test_classifies_longest_matching_category_term(self):
        categories = {
            "Personal care": {"keywords": ["deo", "deodorant"], "brands": []},
            "Home care": {"keywords": ["vask"], "brands": []},
        }
        self.assertEqual(categorize_offer("Nivea deodorant", categories), "Personal care")
        self.assertIsNone(categorize_offer("Ukendt produkt", categories))

    def test_category_rules_include_household_consumables_and_tissues(self):
        config_path = Path(__file__).resolve().parents[1] / "config" / "categories.json"
        config = json.loads(config_path.read_text(encoding="utf-8"))
        categories = config["categories"]
        exclusions = config["excluded_terms"]
        for product in (
            "Alufolie", "Bagepapir", "Fryseposer", "Affaldsposer", "Alubakker",
            "Maskinopvask",
        ):
            self.assertEqual(categorize_offer(product, categories, exclusions), "Home care")
        self.assertEqual(
            categorize_offer("Kleenex lommetørklæder", categories, exclusions),
            "Personal care",
        )
        self.assertEqual(
            categorize_offer("Tena duo pakke", categories, exclusions), "Personal care"
        )
        self.assertEqual(
            categorize_offer("Batiste tørshampoo", categories, exclusions), "Personal care"
        )
        for product in (
            "CIEN ansigtsrens", "Nivea ansigtsrens", "Cosrx ansigtsrens", "Rimmel mascara",
        ):
            self.assertEqual(
                categorize_offer(product, categories, exclusions), "Personal care"
            )
        for product in ("Godt papir køkkenrulle", "Domestos WC rens", "Bref WC-blok"):
            self.assertEqual(categorize_offer(product, categories, exclusions), "Home care")
        self.assertEqual(
            categorize_offer(
                "Kæmpekøb papir", categories, exclusions, "toiletpapir og køkkenrulle"
            ),
            "Home care",
        )

    def test_diapers_and_oral_care_are_excluded_from_both_categories(self):
        config_path = Path(__file__).resolve().parents[1] / "config" / "categories.json"
        config = json.loads(config_path.read_text(encoding="utf-8"))
        for product in (
            "Libero bleer", "Colgate tandpasta", "Oral-B tandbørste", "Mundskyl",
            "Eltandbørste", "Sensodyne tandpasta", "Parodontax", "Jordan eltandbørste",
            "Libresse hygiejnebind", "Always tamponer", "Barbercreme", "Skrabere",
            "Gillette skraber",
        ):
            self.assertIsNone(
                categorize_offer(product, config["categories"], config["excluded_terms"])
            )

    def test_matches_danish_retailer_alias(self):
        retailers = [{"name": "Føtex", "aliases": ["fotex"]}]
        self.assertEqual(_retailer_display_name("føtex", retailers), "Føtex")

    def test_searches_configured_keywords_and_brands_once(self):
        config = {
            "search_terms": ["shampoo", "Nivea ansigtsrens"],
            "categories": {
                "Personal care": {"keywords": ["shampoo", "deo"], "brands": ["Nivea"]},
            },
        }
        self.assertEqual(
            _search_terms(config), ["shampoo", "Nivea ansigtsrens", "deo", "Nivea"]
        )

    def test_fits_image_in_inches_without_shrinking_it(self):
        width, height = _fit_image_dimensions(300, 199, 2.85, 2.48)
        self.assertAlmostEqual(width, 2.85)
        self.assertAlmostEqual(height, 1.8905, places=3)

    def test_parses_publication_front_and_dates(self):
        key = base64.b64encode(b'["fronts",{"businessSlugs":["Netto"]}]').decode()
        payload = [{"publications": [{
            "id": "week42",
            "label": "Netto uge 42",
            "view": "https://images.example/front.webp?w=700&sig=abc",
            "validFrom": "2026-10-09T22:00:00.000Z",
            "validUntil": "2026-10-16T22:00:00.000Z",
            "publish": "2026-10-07T05:00:00.000Z",
        }]}]
        page = f'<app-data data-key="{key}">{html.escape(json.dumps(payload))}</app-data>'
        publications = parse_publications(page)
        self.assertEqual(len(publications), 1)
        self.assertEqual(publications[0].id, "week42")
        self.assertIn("front.webp", publications[0].cover_image)

    def test_parses_coop_publication_id_and_viewer_key(self):
        parser = JsonLdParser()
        parser.feed(
            '<script id="sgn-sdk" data-app-key="public-key"></script>'
            '<div class="incito-wrapper" data-publication-id="coop-42"></div>'
        )
        self.assertEqual(parser.coop_api_key, "public-key")
        self.assertEqual(parser.incito_publication_id, "coop-42")

    def test_parses_coop_cover_carousel(self):
        parser = CoopCoverParser()
        parser.feed(
            '<div class="catalog-images-slider coop-slider"><div class="catalog-image">'
            '<a href="https://kvickly.coop.dk/avis/">'
            '<img src="https://images.example/kvickly-cover.webp"></a>'
            '</div></div>'
        )
        self.assertEqual(
            parser.covers,
            [("https://kvickly.coop.dk/avis/", "https://images.example/kvickly-cover.webp")],
        )

    def test_coop_catalogs_and_offers_must_overlap_target_week(self):
        self.assertTrue(
            _overlaps_week("2026-10-08T22:00:00+0000", "2026-10-14T21:59:59+0000", 42)
        )
        self.assertFalse(
            _overlaps_week("2026-10-03T22:00:00+0000", "2026-10-11T21:59:59+0000", 42)
        )

    def test_selects_publication_for_requested_week(self):
        issues = [
            Publication("old", "Netto uge 41", "old.jpg", "", "", "2026-09-30"),
            Publication("special", "Halloween", "halloween.jpg", "", "", "2026-10-05"),
            Publication("new", "Netto uge 42", "new.jpg", "", "", "2026-10-07"),
        ]
        self.assertEqual(select_week_publication(issues, 42).id, "new")

    def test_filters_out_offers_from_other_publications(self):
        offers = [
            Offer("Omo vaskemiddel", "old", "image", "Netto", "20", "", publication_id="week41"),
            Offer("Omo vaskemiddel", "new", "image", "Netto", "20", "", publication_id="week42"),
        ]
        config = {
            "retailers": [{"name": "Netto", "aliases": ["netto"]}],
            "categories": {"Home care": {"keywords": ["vaskemiddel"], "brands": ["Omo"]}},
        }
        publication = Publication("week42", "Netto uge 42", "front.jpg", "", "", "")
        selected = _filter_offers(offers, config, {"Netto": publication})
        self.assertEqual([offer.url for offer in selected], ["new"])

    def test_collects_unrecognized_week_offers_but_not_excluded_products(self):
        offers = [
            Offer("Ukendt hudplejeprodukt", "review", "image", "Netto", "20", "",
                  publication_id="week42"),
            Offer("Colgate tandpasta", "excluded", "image", "Netto", "20", "",
                  publication_id="week42"),
        ]
        config = {
            "retailers": [{"name": "Netto", "aliases": ["netto"]}],
            "search_terms": ["personlig pleje"],
            "categories": {"Personal care": {"keywords": ["shampoo"], "brands": []}},
            "excluded_terms": ["colgate", "tandpasta"],
        }
        publication = Publication("week42", "Netto uge 42", "front.jpg", "", "", "")
        candidates = []
        selected = _filter_offers(
            offers, config, {"Netto": publication}, uncategorized=candidates
        )
        self.assertEqual(selected, [])
        self.assertEqual([offer.url for offer in candidates], ["review"])

    def test_coop_review_candidates_must_match_configured_search_terms(self):
        offers = [
            Offer("Ukendt fødevare", "food", "image", "Netto", "20", "",
                  publication_id="week42", description="rugbrød"),
            Offer("Ukendt plejeprodukt", "care", "image", "Netto", "20", "",
                  publication_id="week42", description="personlig pleje"),
        ]
        config = {
            "retailers": [{"name": "Netto", "aliases": ["netto"]}],
            "search_terms": ["personlig pleje"],
            "categories": {"Personal care": {"keywords": ["shampoo"], "brands": []}},
        }
        publication = Publication(
            "week42", "Netto uge 42", "front.jpg", "", "", "", source="coop"
        )
        candidates = []
        _filter_offers(offers, config, {"Netto": publication}, uncategorized=candidates)
        self.assertEqual([offer.url for offer in candidates], ["care"])


if __name__ == "__main__":
    unittest.main()
