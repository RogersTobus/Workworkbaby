import unittest

from openpyxl import Workbook

from brand_connect_proposal import (
    content_url_score,
    count_unique_creators,
    normalize_campaign_status,
    normalize_proposal_date,
)
from brand_connect_sheet import (
    find_favorite_date_column,
    first_header_columns,
    is_favorite_candidate,
    is_legacy_favorite_value,
    row_has_proposal_status,
    set_favorite_date_if_blank,
    set_proposal_date_if_blank,
    should_add_missing_campaign_creator,
)


class BrandConnectProposalDateTest(unittest.TestCase):
    def setUp(self):
        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self.sheet.append(["제안날짜", "크리에이터", "오일", "병절임", "파우치"])

    def tearDown(self):
        self.workbook.close()

    def test_automation_never_sets_blank_date(self):
        self.sheet.append(["", "creator", "수락", "", ""])

        changed = set_proposal_date_if_blank(
            self.sheet,
            row=2,
            date_column=1,
            date_text="2026.07.30",
        )

        self.assertFalse(changed)
        self.assertEqual(self.sheet["A2"].value, "")

    def test_preserves_existing_date(self):
        self.sheet.append(["2026.07.29", "creator", "거절", "", ""])

        changed = set_proposal_date_if_blank(
            self.sheet,
            row=2,
            date_column=1,
            date_text="2026.07.30",
        )

        self.assertFalse(changed)
        self.assertEqual(self.sheet["A2"].value, "2026.07.29")

    def test_favorite_sets_date_only_when_blank(self):
        self.sheet.append(["", "creator", "", "", ""])

        changed = set_favorite_date_if_blank(
            self.sheet,
            row=2,
            date_column=1,
            date_text="2026.07.31",
        )

        self.assertTrue(changed)
        self.assertEqual(self.sheet["A2"].value, "2026.07.31")

    def test_favorite_preserves_existing_date(self):
        self.sheet.append(["2026.07.30", "creator", "", "", ""])

        changed = set_favorite_date_if_blank(
            self.sheet,
            row=2,
            date_column=1,
            date_text="2026.07.31",
        )

        self.assertFalse(changed)
        self.assertEqual(self.sheet["A2"].value, "2026.07.30")

    def test_favorite_candidate_requires_blank_date_and_product(self):
        self.assertTrue(is_favorite_candidate("", ""))
        self.assertFalse(is_favorite_candidate("2026.07.31", ""))
        self.assertFalse(is_favorite_candidate("", "대기"))
        self.assertFalse(is_favorite_candidate("2026.07.31", "대기"))

    def test_alp_uses_product_specific_proposal_date_columns(self):
        headers = {
            "제안날짜(이뮨)": 6,
            "알프이뮨": 7,
            "제안날짜(아이언드롭)": 8,
            "알프아이언드롭": 9,
        }

        self.assertEqual(
            find_favorite_date_column(headers, "alp", "immun"),
            6,
        )
        self.assertEqual(
            find_favorite_date_column(headers, "alp", "iron_drop"),
            8,
        )

    def test_gaia_keeps_common_proposal_date_column(self):
        self.assertEqual(
            find_favorite_date_column({"제안날짜": 1}, "gaia", "oil"),
            1,
        )

    def test_gaia_uses_product_specific_proposal_date_columns(self):
        headers = {
            "제안날짜(오일)": 6,
            "제안날짜(병절임)": 8,
            "제안날짜(파우치)": 10,
        }

        self.assertEqual(
            find_favorite_date_column(headers, "gaia", "oil"),
            6,
        )
        self.assertEqual(
            find_favorite_date_column(headers, "gaia", "pickles"),
            8,
        )
        self.assertEqual(
            find_favorite_date_column(headers, "gaia", "pouch"),
            10,
        )

    def test_recognizes_only_proposal_status_values(self):
        self.sheet.append(["", "creator", "대기", "", ""])
        self.sheet.append(["", "creator2", "찜", "", ""])

        self.assertTrue(row_has_proposal_status(self.sheet, 2, [3, 4, 5]))
        self.assertFalse(row_has_proposal_status(self.sheet, 3, [3, 4, 5]))

    def test_legacy_favorite_is_not_a_proposal_status(self):
        self.assertTrue(is_legacy_favorite_value("찜"))
        self.assertFalse(is_legacy_favorite_value("대기"))

    def test_all_campaign_statuses_add_a_missing_creator(self):
        for status in ("대기", "수락", "거절"):
            with self.subTest(status=status):
                self.assertTrue(should_add_missing_campaign_creator(status, []))
        self.assertFalse(should_add_missing_campaign_creator("찜", []))
        self.assertFalse(should_add_missing_campaign_creator("수락", [12, 13]))

    def test_duplicate_summary_header_does_not_replace_product_column(self):
        self.sheet["C1"] = "기존 상태"
        self.sheet["G1"] = "오일"
        self.sheet["O1"] = "오일"
        self.assertEqual(first_header_columns(self.sheet)["오일"], 7)

    def test_creator_count_removes_campaign_duplicates(self):
        self.assertEqual(
            count_unique_creators(
                [
                    {"creator": "같은 사람"},
                    {"creator": " 같은   사람 "},
                    {"creator": "다른 사람"},
                ]
            ),
            2,
        )

    def test_progress_status_is_accepted(self):
        self.assertEqual(normalize_campaign_status("진행 중"), "수락")
        self.assertEqual(normalize_campaign_status("진행중"), "수락")
        self.assertEqual(normalize_campaign_status("진행 완료"), "수락")
        self.assertEqual(normalize_campaign_status("진행완료"), "수락")
        self.assertEqual(normalize_campaign_status("정산/확정 필요"), "수락")

    def test_proposal_date_is_normalized(self):
        self.assertEqual(normalize_proposal_date("2026-07-30"), "2026.07.30")
        self.assertEqual(normalize_proposal_date("2026.07.30"), "2026.07.30")

    def test_invalid_proposal_date_is_rejected(self):
        with self.assertRaises(ValueError):
            normalize_proposal_date("2026.07.40")

    def test_submitted_creator_content_beats_visible_help_links(self):
        self.assertGreater(
            content_url_score("https://www.instagram.com/p/DbJ_N5WEkU8/"),
            content_url_score("https://blog.naver.com/brandconnect/223327624176"),
        )

    def test_creator_blog_is_content_but_brandconnect_help_blog_is_not(self):
        self.assertGreater(
            content_url_score("https://blog.naver.com/example_creator/223123456789"),
            0,
        )
        self.assertEqual(
            content_url_score("https://blog.naver.com/brandconnect/223327624176"),
            0,
        )



if __name__ == "__main__":
    unittest.main()
