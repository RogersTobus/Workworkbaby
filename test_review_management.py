import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import dm_assistant


class ReviewManagementTests(unittest.TestCase):
    def test_gaea_defaults_follow_the_source_deck_order(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "review_templates.json"
            with patch.object(dm_assistant, "REVIEW_TEMPLATES_PATH", path):
                titles = [
                    item["title"]
                    for item in dm_assistant.get_review_templates()["brands"]["gaia"]
                ]

        self.assertEqual(
            [
                "맛 / 풍미 만족",
                "품질 / 신선도 만족",
                "재구매 / 선물 만족",
                "가격 관련 의견",
                "향이 강하다 / 쓴맛 관련",
                "포장 / 누유 / 뚜껑 관련",
                "빠른 공통 답변",
            ],
            titles,
        )

    def test_template_text_can_be_freely_edited_and_saved(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "review_templates.json"
            with patch.object(dm_assistant, "REVIEW_TEMPLATES_PATH", path):
                result = dm_assistant.update_review_template(
                    {
                        "action": "save",
                        "brand": "gaia",
                        "id": "gaia-taste",
                        "title": "맛 / 풍미 만족",
                        "keywords": "고소하다, 향이 깔끔하다",
                        "message": "첫 문장\n\n사용자가 자유롭게 고친 문장",
                    }
                )
                saved = next(
                    item
                    for item in result["brands"]["gaia"]
                    if item["id"] == "gaia-taste"
                )

                self.assertEqual("첫 문장\n\n사용자가 자유롭게 고친 문장", saved["message"])
                self.assertEqual(["고소하다", "향이 깔끔하다"], saved["keywords"])

    def test_an_intentionally_empty_brand_list_is_not_replaced_by_defaults(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "review_templates.json"
            path.write_text('{"alp": [], "gaia": []}', encoding="utf-8")
            with patch.object(dm_assistant, "REVIEW_TEMPLATES_PATH", path):
                brands = dm_assistant.get_review_templates()["brands"]

        self.assertEqual([], brands["alp"])
        self.assertEqual([], brands["gaia"])


if __name__ == "__main__":
    unittest.main()
