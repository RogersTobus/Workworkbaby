import unittest

from calendar_sync import _parse_sheet_values


class CalendarSyncParsingTest(unittest.TestCase):
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
