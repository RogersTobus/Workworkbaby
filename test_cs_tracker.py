import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import dm_assistant


class CSTrackerTests(unittest.TestCase):
    def test_create_update_and_delete_case(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cs_cases.json"
            with patch.object(dm_assistant, "CS_CASES_PATH", path):
                created = dm_assistant.update_cs_case(
                    {
                        "action": "save",
                        "received_date": "2026-08-05",
                        "brand": "가이아",
                        "partner": "베네피아",
                        "order_number": "ORDER-1",
                        "tracking_number": "TRACK-1",
                        "customer_name": "홍길동",
                        "contact": "010-1234-5678",
                        "status": "접수",
                    }
                )["case"]
                self.assertTrue(created["id"])
                self.assertEqual("접수", created["status"])
                self.assertEqual("2026-08-05", created["received_date"])

                updated = dm_assistant.update_cs_case(
                    {
                        **created,
                        "action": "save",
                        "status": "완료",
                    }
                )["case"]
                self.assertEqual(created["id"], updated["id"])
                self.assertEqual("완료", updated["status"])
                self.assertEqual(1, len(dm_assistant.get_cs_cases()["cases"]))

                dm_assistant.update_cs_case(
                    {"action": "delete", "id": created["id"]}
                )
                self.assertEqual([], dm_assistant.get_cs_cases()["cases"])

    def test_required_fields_are_checked(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cs_cases.json"
            with patch.object(dm_assistant, "CS_CASES_PATH", path):
                with self.assertRaisesRegex(ValueError, "브랜드"):
                    dm_assistant.update_cs_case(
                        {"action": "save", "status": "접수"}
                    )


if __name__ == "__main__":
    unittest.main()
