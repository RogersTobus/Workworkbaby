import unittest
from datetime import date

from work_log import (
    build_work_log_draft,
    classify_brands,
    classify_partner,
    latest_or_upcoming_friday,
)


class WorkLogDraftTests(unittest.TestCase):
    def test_weekend_uses_latest_friday(self):
        self.assertEqual(
            latest_or_upcoming_friday(date(2026, 8, 1)),
            date(2026, 7, 31),
        )

    def test_weekday_uses_that_weeks_friday(self):
        self.assertEqual(
            latest_or_upcoming_friday(date(2026, 8, 3)),
            date(2026, 8, 7),
        )

    def test_classifies_partner_and_brands(self):
        text = "베네피아 가이아 병절임 행사 제안"
        self.assertEqual(classify_partner(text), "베네피아")
        self.assertEqual(classify_brands(text), ["gaia"])

    def test_draft_carries_unfinished_work_to_next_week(self):
        draft = build_work_log_draft(
            [
                {
                    "id": "1",
                    "date": "2026-07-30",
                    "text": "알프 롯데아이몰 행사 제안",
                    "kind": "task",
                    "status": "todo",
                },
                {
                    "id": "2",
                    "date": "2026-07-29",
                    "text": "가이아 베네피아 상품 등록",
                    "kind": "task",
                    "status": "done",
                },
            ],
            date(2026, 7, 31),
        )
        alp = draft["sales"]["alp"][0]
        self.assertIn("알프 롯데아이몰 행사 제안", alp["next_week"])
        self.assertEqual(draft["summary"]["carry_over"], 1)

    def test_draft_uses_existing_sheet_tone_without_status_labels(self):
        draft = build_work_log_draft(
            [
                {
                    "id": "1",
                    "date": "2026-07-27",
                    "text": "G마켓 가격 수정",
                    "kind": "task",
                    "status": "done",
                },
                {
                    "id": "2",
                    "date": "2026-07-28",
                    "text": "G마켓 행사 전 확인",
                    "kind": "task",
                    "status": "done",
                },
            ],
            date(2026, 7, 31),
        )
        text = next(
            item["this_week"]
            for item in draft["partners"]
            if item["partner"] == "G마켓/옥션"
        )
        self.assertEqual(text, "1) G마켓 가격 수정\n2) G마켓 행사 전 확인")
        self.assertNotIn("[완료]", text)

    def test_event_is_written_as_progress_in_sheet_tone(self):
        draft = build_work_log_draft(
            [
                {
                    "id": "event-1",
                    "date": "2026-07-27",
                    "end_date": "2026-07-28",
                    "text": "SSG_가이아_쓱특가",
                    "kind": "event",
                }
            ],
            date(2026, 7, 31),
        )
        text = next(
            item["this_week"]
            for item in draft["partners"]
            if item["partner"] == "SSG"
        )
        self.assertEqual(text, "SSG 가이아 쓱특가 진행")


if __name__ == "__main__":
    unittest.main()
