import unittest
from unittest.mock import patch

import dm_assistant


class HomeDashboardTests(unittest.TestCase):
    @patch.object(dm_assistant.BRAND_CONNECT_PROPOSAL_MANAGER, "snapshot")
    @patch.object(dm_assistant.BRAND_CONNECT_FAVORITE_MANAGER, "snapshot")
    @patch.object(dm_assistant.BRAND_CONNECT_MANAGER, "snapshot")
    @patch.object(dm_assistant.PRICE_MANAGER, "snapshot")
    @patch.object(dm_assistant, "get_dashboard")
    @patch.object(dm_assistant, "get_cs_cases")
    @patch.object(dm_assistant, "get_calendar_checklist_routines")
    @patch.object(dm_assistant, "get_calendar_today")
    def test_compacts_live_work_data(
        self,
        calendar_today,
        checklist_routines,
        cs_cases,
        dm_dashboard,
        price_snapshot,
        crawl_snapshot,
        favorite_snapshot,
        proposal_snapshot,
    ):
        calendar_today.return_value = {
            "date": "2026-08-05",
            "tasks": [
                {"id": "low", "text": "낮은 우선순위", "priority": 1, "completed": False},
                {"id": "high", "text": "높은 우선순위", "priority": 3, "completed": False},
                {"id": "done", "text": "완료 업무", "priority": 2, "completed": True},
            ],
            "nearby_events": [
                {"source_id": "event-1", "text": "진행 행사", "event_start": "2026-08-04", "event_end": "2026-08-06", "event_label": "8/4~8/6", "event_type": "popup", "event_type_label": "팝업행사"},
                {"source_id": "event-2", "text": "어제 행사", "event_start": "2026-08-03", "event_end": "2026-08-04", "event_label": "8/3~8/4", "event_type": "special", "event_type_label": "특가전"},
            ],
        }
        checklist_routines.return_value = {"routines": [{"id": "r1", "text": "루틴", "completed": False}]}
        cs_cases.return_value = {"cases": [{"id": "c1", "status": "접수"}, {"id": "c2", "status": "완료"}]}
        dm_dashboard.side_effect = lambda brand, force_sync=False: {
            "brand_name": brand.upper(), "goal": 30, "weekly_sent": 12,
            "remaining_to_goal": 18, "pending_count": 4,
            "targets": [{"username": "should-not-leak"}],
        }
        price_snapshot.return_value = {"status": "idle", "message": "대기 중", "processed": 49, "total": 49}
        crawl_snapshot.return_value = {"status": "idle", "message": "대기 중"}
        favorite_snapshot.return_value = {"status": "running", "message": "처리 중"}
        proposal_snapshot.return_value = {"status": "error", "message": "로그인 필요"}

        result = dm_assistant.get_home_dashboard()

        self.assertEqual(result["tasks"]["remaining"], 2)
        self.assertEqual(result["tasks"]["completed"], 1)
        self.assertEqual(result["tasks"]["items"][0]["id"], "high")
        self.assertEqual([item["id"] for item in result["events"]], ["event-1"])
        self.assertEqual(result["routines"]["remaining"], 1)
        self.assertEqual(result["cs"]["open"], 1)
        self.assertEqual(len(result["dm"]), 2)
        self.assertNotIn("targets", result["dm"][0])
        self.assertEqual(result["systems"]["automations"]["favorite"]["status"], "running")


if __name__ == "__main__":
    unittest.main()
