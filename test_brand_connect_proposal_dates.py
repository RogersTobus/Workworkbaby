import unittest

from openpyxl import Workbook

from brand_connect_proposal import normalize_campaign_status
from brand_connect_sheet import (
    row_has_proposal_status,
    set_favorite_date_if_blank,
    set_proposal_date_if_blank,
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

    def test_recognizes_only_proposal_status_values(self):
        self.sheet.append(["", "creator", "대기", "", ""])
        self.sheet.append(["", "creator2", "찜", "", ""])

        self.assertTrue(row_has_proposal_status(self.sheet, 2, [3, 4, 5]))
        self.assertFalse(row_has_proposal_status(self.sheet, 3, [3, 4, 5]))

    def test_progress_status_is_accepted(self):
        self.assertEqual(normalize_campaign_status("진행 중"), "수락")
        self.assertEqual(normalize_campaign_status("진행중"), "수락")


if __name__ == "__main__":
    unittest.main()
