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

    def test_instagram_redirect_does_not_fail_navigation(self):
        class RedirectingPage:
            waited = False
            timeout = 0

            def goto(self, *_args, **_kwargs):
                raise RuntimeError(
                    "Page.goto: Navigation is interrupted by another navigation"
                )

            def wait_for_load_state(self, *_args, **_kwargs):
                self.waited = True

            def wait_for_timeout(self, timeout):
                self.timeout = timeout

        page = RedirectingPage()
        InstagramDMSenderManager._goto_instagram(
            page,
            "https://www.instagram.com/",
        )
        self.assertTrue(page.waited)
        self.assertEqual(1_500, page.timeout)

    def test_onetap_username_is_not_treated_as_logged_in_account(self):
        class Body:
            def inner_text(self):
                return "Continue as alpnutrition_official_kr"

            def count(self):
                return 0

        class OneTapPage:
            url = "https://www.instagram.com/accounts/onetap/?lsrc=ci"

            def locator(self, _selector):
                return Body()

        self.assertFalse(
            InstagramDMSenderManager._account_visible(
                OneTapPage(),
                "@alpnutrition_official_kr",
            )
        )

    def test_navigation_call_log_is_hidden_from_status(self):
        error = RuntimeError(
            "Page.goto interrupted by another navigation Call log: details"
        )
        message = InstagramDMSenderManager._brief_error(error)
        self.assertNotIn("Call log", message)


if __name__ == "__main__":
    unittest.main()
