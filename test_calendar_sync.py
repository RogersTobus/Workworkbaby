import unittest
from datetime import date

from calendar_sync import _parse_sheet_values


class CalendarSyncParsingTest(unittest.TestCase):
    def test_ignores_past_month_sections(self):
        rows = [
            ["2026.08"],
            ["현재 행사_26.08.10~26.08.12"],
            ["2026.07"],
            ["삭제했지만 과거 칸에 남은 행사_26.07.28~26.08.04"],
        ]

        events = _parse_sheet_values(rows, reference_date=date(2026, 8, 5))

        self.assertEqual([item["text"] for item in events], ["현재 행사"])

    def test_keeps_event_spanning_into_future_month_section(self):
        rows = [
            ["2026.09"],
            ["월 경계 행사_26.08.28~26.09.03"],
            ["2026.08"],
            ["이번 달 행사_26.08.10~26.08.12"],
        ]

        events = _parse_sheet_values(rows, reference_date=date(2026, 8, 5))

        self.assertEqual(
            [item["text"] for item in events],
            ["이번 달 행사", "월 경계 행사"],
        )

    def test_ignores_date_not_overlapping_its_month_section(self):
        rows = [
            ["2026.08"],
            ["잘못 복사된 행사_26.07.10~26.07.18"],
        ]

        events = _parse_sheet_values(rows, reference_date=date(2026, 8, 5))

        self.assertEqual(events, [])

    def test_prefers_most_repeated_overlapping_period(self):
        rows = [
            ["베네피아_가이아_화끈딜_26.07.28~26.08.04"],
            ["베네피아_가이아_화끈딜_26.07.28~26.08.04"],
            ["베네피아_가이아_화끈딜_26.07.28~26.08.04"],
            ["베네피아_가이아_화끈딜_26.07.24~26.08.11"],
            ["베네피아_가이아_화끈딜_26.07.24~26.08.11"],
        ]

        events = _parse_sheet_values(rows)

        self.assertEqual(len(events), 1)
        self.assertEqual(events[0]["date"], "2026-07-28")
        self.assertEqual(events[0]["end_date"], "2026-08-04")

    def test_keeps_separate_non_overlapping_occurrences(self):
        rows = [
            ["SSG_가이아_오반장_26.07.26"],
            ["SSG_가이아_오반장_26.07.28"],
        ]

        events = _parse_sheet_values(rows)

        self.assertEqual(
            [(item["date"], item["end_date"]) for item in events],
            [
                ("2026-07-26", "2026-07-26"),
                ("2026-07-28", "2026-07-28"),
            ],
        )

    def test_keeps_event_type_rules(self):
        rows = [
            ["현대백화점_가이아_팝업행사_26.08.28~26.09.03"],
            ["하루약사_알프_공구 1차_26.08.24~26.08.30"],
        ]

        events = _parse_sheet_values(rows)

        self.assertEqual(
            [item["event_type"] for item in events],
            ["group_buy", "popup"],
        )


if __name__ == "__main__":
    unittest.main()
