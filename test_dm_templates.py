import unittest

from dm_templates import DM_MESSAGE_TEMPLATES


class DmTemplateTests(unittest.TestCase):
    def test_each_template_has_one_instagram_id_placeholder(self):
        for brand, template in DM_MESSAGE_TEMPLATES.items():
            with self.subTest(brand=brand):
                self.assertEqual(template.count("{instagram_id}"), 1)

    def test_alp_template_contains_required_identity_and_ending(self):
        template = DM_MESSAGE_TEMPLATES["alp"]
        self.assertIn("알프뉴트리션 공식 수입·유통사", template)
        self.assertTrue(template.endswith("주식회사 투버스\n박영준 드림"))

    def test_gaia_template_contains_required_identity_and_ending(self):
        template = DM_MESSAGE_TEMPLATES["gaia"]
        self.assertIn("GAEA(가이아) 공식 수입·유통사", template)
        self.assertTrue(template.endswith("주식회사 투버스\n박영준 드림"))


if __name__ == "__main__":
    unittest.main()
