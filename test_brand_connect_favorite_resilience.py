import unittest

from brand_connect_favorite import BrandConnectFavoriteManager


class _DisabledButton:
    def __init__(self):
        self.clicked = False

    def is_enabled(self):
        return False

    def click(self, **_kwargs):
        self.clicked = True


class _FavoriteLocator:
    def __init__(self, button):
        self.first = button

    def count(self):
        return 1


class _Card:
    def __init__(self, button):
        self.button = button

    def locator(self, _selector):
        return _FavoriteLocator(self.button)


class _CardLocator:
    def __init__(self, button):
        self.first = _Card(button)

    def filter(self, **_kwargs):
        return self

    def count(self):
        return 1


class _Page:
    def __init__(self, button):
        self.button = button

    def locator(self, _selector):
        return _CardLocator(self.button)


class FavoriteAutomationResilienceTests(unittest.TestCase):
    def test_disabled_favorite_button_is_skipped_without_waiting_for_click(self):
        button = _DisabledButton()
        with self.assertRaisesRegex(LookupError, "비활성화"):
            BrandConnectFavoriteManager._open_favorite_popup(
                _Page(button),
                "테스트 크리에이터",
            )
        self.assertFalse(button.clicked)

    def test_browser_call_log_is_hidden_from_user_message(self):
        error = RuntimeError(
            "Locator.click: Timeout 5000ms exceeded. Call log: internal details"
        )
        message = BrandConnectFavoriteManager._brief_error(error)
        self.assertEqual("화면 응답이 없어 다음 후보로 넘어갑니다.", message)
        self.assertNotIn("Call log", message)


if __name__ == "__main__":
    unittest.main()
