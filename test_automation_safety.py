import unittest
from unittest.mock import patch

import dm_assistant


class AutomationSafetyTests(unittest.TestCase):
    def test_drive_preflight_fails_closed(self):
        with patch.object(
            dm_assistant,
            "sync_dm_workbook",
            return_value={
                "status": "login_required",
                "message": "Google Drive 로그인이 필요합니다.",
            },
        ):
            with self.assertRaisesRegex(
                dm_assistant.GoogleDriveLoginRequired,
                "로그인이 필요",
            ):
                dm_assistant.require_dm_workbook_sync()

    def test_browser_automation_start_is_blocked_by_other_job(self):
        states = {
            "dm": {"status": "idle"},
            "prices": {"status": "crawling"},
            "crawl": {"status": "idle"},
            "favorite": {"status": "idle"},
            "proposal": {"status": "idle"},
        }
        callback_called = False

        def callback():
            nonlocal callback_called
            callback_called = True

        with patch.object(
            dm_assistant, "browser_automation_snapshots", return_value=states
        ):
            with self.assertRaisesRegex(ValueError, "네이버 가격 최신화"):
                dm_assistant.start_exclusive_browser_automation("dm", callback)
        self.assertFalse(callback_called)

    def test_browser_automation_starts_when_slot_is_free(self):
        states = {
            key: {"status": "idle"}
            for key in ("dm", "prices", "crawl", "favorite", "proposal")
        }
        with patch.object(
            dm_assistant, "browser_automation_snapshots", return_value=states
        ):
            result = dm_assistant.start_exclusive_browser_automation(
                "dm", lambda: {"status": "queued"}
            )
        self.assertEqual(result["status"], "queued")

    def test_drive_browser_automation_checks_drive_before_start(self):
        callback_called = False

        def callback():
            nonlocal callback_called
            callback_called = True
            return {"status": "queued"}

        with patch.object(
            dm_assistant,
            "require_dm_workbook_sync",
            side_effect=dm_assistant.GoogleDriveLoginRequired("login required"),
        ), patch.object(
            dm_assistant,
            "browser_automation_snapshots",
            return_value={
                key: {"status": "idle"}
                for key in ("dm", "prices", "crawl", "favorite", "proposal")
            },
        ):
            with self.assertRaises(dm_assistant.GoogleDriveLoginRequired):
                dm_assistant.start_drive_browser_automation("crawl", callback)
        self.assertFalse(callback_called)


if __name__ == "__main__":
    unittest.main()
