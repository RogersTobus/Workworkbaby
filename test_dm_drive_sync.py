import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import dm_assistant


class FakeResponse(io.BytesIO):
    def __init__(self, body=b"", headers=None):
        super().__init__(body)
        self.headers = headers or {}

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        self.close()


class DriveSyncFallbackTests(unittest.TestCase):
    def test_download_uses_shared_url_when_drive_login_expired(self):
        with tempfile.TemporaryDirectory() as directory:
            token_path = Path(directory) / "drive_token.json"
            token_path.write_text("{}", encoding="utf-8")
            destination = Path(directory) / "download.xlsx"
            responses = [
                FakeResponse(headers={"Last-Modified": "remote-time"}),
                FakeResponse(b"xlsx-bytes", headers={"Content-Length": "10"}),
            ]
            config = {
                "drive_file_id": "file-id",
                "download_url": "https://example.invalid/workbook.xlsx",
            }
            with (
                patch.object(dm_assistant, "DM_DRIVE_TOKEN_PATH", token_path),
                patch.dict(dm_assistant.CONFIG, {"dm_sync": config}),
                patch.object(
                    dm_assistant,
                    "load_dm_drive_service",
                    side_effect=dm_assistant.GoogleDriveLoginRequired("expired"),
                ),
                patch.object(
                    dm_assistant.urllib.request,
                    "urlopen",
                    side_effect=responses,
                ),
            ):
                modified = dm_assistant.download_dm_workbook(
                    destination,
                    config["download_url"],
                )

            self.assertEqual(modified, "remote-time")
            self.assertEqual(destination.read_bytes(), b"xlsx-bytes")

    def test_upload_reports_login_required(self):
        with tempfile.TemporaryDirectory() as directory:
            token_path = Path(directory) / "drive_token.json"
            token_path.write_text("{}", encoding="utf-8")
            workbook_path = Path(directory) / "workbook.xlsx"
            workbook_path.write_bytes(b"xlsx")
            with (
                patch.object(dm_assistant, "DM_DRIVE_TOKEN_PATH", token_path),
                patch.dict(
                    dm_assistant.CONFIG,
                    {
                        "dm_sync": {"drive_file_id": "file-id"},
                        "_workbook_path": workbook_path,
                    },
                ),
                patch.object(
                    dm_assistant,
                    "load_dm_drive_service",
                    side_effect=dm_assistant.GoogleDriveLoginRequired("reconnect"),
                ),
            ):
                result = dm_assistant.upload_dm_workbook_to_drive()

            self.assertEqual(result["status"], "login_required")
            self.assertEqual(result["message"], "reconnect")


if __name__ == "__main__":
    unittest.main()
