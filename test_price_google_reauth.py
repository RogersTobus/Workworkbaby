import json
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

from price_updater import PriceUpdateManager


class _RefreshError(Exception):
    pass


class _ExpiredCredentials:
    expired = True
    refresh_token = "expired-refresh-token"
    valid = False

    def refresh(self, _request):
        raise _RefreshError("invalid_grant")


class _FreshCredentials:
    expired = False
    refresh_token = "fresh-refresh-token"
    valid = True

    def to_json(self):
        return json.dumps({"refresh_token": self.refresh_token})


class GoogleReauthorizationTests(unittest.TestCase):
    def test_revoked_refresh_token_falls_back_to_interactive_login(self):
        login_calls = []
        fresh = _FreshCredentials()

        class Credentials:
            @staticmethod
            def from_authorized_user_file(_path, _scopes):
                return _ExpiredCredentials()

        class InstalledAppFlow:
            @staticmethod
            def from_client_secrets_file(_path, _scopes):
                return types.SimpleNamespace(
                    run_local_server=lambda **kwargs: (
                        login_calls.append(kwargs) or fresh
                    )
                )

        modules = {
            "google": types.ModuleType("google"),
            "google.auth": types.ModuleType("google.auth"),
            "google.auth.exceptions": types.SimpleNamespace(
                RefreshError=_RefreshError
            ),
            "google.auth.transport": types.ModuleType("google.auth.transport"),
            "google.auth.transport.requests": types.SimpleNamespace(
                Request=lambda: object()
            ),
            "google.oauth2": types.ModuleType("google.oauth2"),
            "google.oauth2.credentials": types.SimpleNamespace(
                Credentials=Credentials
            ),
            "google_auth_oauthlib": types.ModuleType("google_auth_oauthlib"),
            "google_auth_oauthlib.flow": types.SimpleNamespace(
                InstalledAppFlow=InstalledAppFlow
            ),
            "googleapiclient": types.ModuleType("googleapiclient"),
            "googleapiclient.discovery": types.SimpleNamespace(
                build=lambda *_args, **_kwargs: "sheets-service"
            ),
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            app_dir = Path(temp_dir)
            (app_dir / "google_credentials.json").write_text(
                "{}", encoding="utf-8"
            )
            token_path = app_dir / "app_data" / "google_token.json"
            token_path.parent.mkdir(parents=True)
            token_path.write_text("{}", encoding="utf-8")
            updater = PriceUpdateManager(
                app_dir,
                {
                    "google_credentials_path": "google_credentials.json",
                    "google_token_path": "app_data/google_token.json",
                },
            )

            with patch.dict(sys.modules, modules):
                service = updater._google_service()

            self.assertEqual("sheets-service", service)
            self.assertEqual([{"port": 0, "open_browser": True}], login_calls)
            self.assertEqual(
                "fresh-refresh-token",
                json.loads(token_path.read_text(encoding="utf-8"))["refresh_token"],
            )


if __name__ == "__main__":
    unittest.main()
