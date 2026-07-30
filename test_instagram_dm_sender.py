import tempfile
import time
import unittest
from pathlib import Path

from instagram_dm_sender import InstagramDMSenderManager


class FakeInstagramDMSender(InstagramDMSenderManager):
    def _run(self, brand: str) -> None:
        self._set(
            status="completed",
            message="done",
            completed=self.snapshot()["requested"],
        )


def dashboard(remaining: int) -> dict:
    return {
        "remaining_to_goal": remaining,
        "message_ready": True,
        "targets": [
            {
                "row": 2,
                "instagram_id": "sample",
                "profile_url": "https://www.instagram.com/sample/",
                "message": "hello",
            }
        ],
        "goal": 20,
        "weekly_sent": 20 - remaining,
    }


class InstagramDMSenderManagerTests(unittest.TestCase):
    def make_manager(self, remaining: int):
        temporary = tempfile.TemporaryDirectory()
        app_dir = Path(temporary.name)
        image = app_dir / "reference.png"
        image.write_bytes(b"png")
        manager = FakeInstagramDMSender(
            app_dir,
            lambda brand, force: dashboard(remaining),
            lambda brand, row, action: {"ok": True},
            {"gaia": {"expected_account": "@gaeagreece_kor"}},
            {"gaia": image},
        )
        return temporary, manager

    def test_start_rejects_already_completed_goal(self):
        temporary, manager = self.make_manager(0)
        self.addCleanup(temporary.cleanup)
        with self.assertRaisesRegex(ValueError, "이미 달성"):
            manager.start("gaia")

    def test_start_uses_exact_remaining_goal(self):
        temporary, manager = self.make_manager(7)
        self.addCleanup(temporary.cleanup)
        manager.start("gaia")
        deadline = time.time() + 1
        while manager.snapshot()["status"] != "completed":
            self.assertLess(time.time(), deadline)
            time.sleep(0.01)
        state = manager.snapshot()
        self.assertEqual(state["requested"], 7)
        self.assertEqual(state["completed"], 7)
        self.assertEqual(state["expected_account"], "@gaeagreece_kor")


if __name__ == "__main__":
    unittest.main()
