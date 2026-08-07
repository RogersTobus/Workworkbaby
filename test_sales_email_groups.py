import json
import subprocess
import tempfile
import unittest
import sys
import types
from pathlib import Path
from unittest.mock import patch

simulation_manager = types.ModuleType("simulation_manager")


class SimulationManager:
    def __init__(self, *args, **kwargs):
        pass


simulation_manager.SimulationManager = SimulationManager
sys.modules.setdefault("simulation_manager", simulation_manager)

import dm_assistant


class SalesEmailGroupTests(unittest.TestCase):
    def test_contact_group_is_saved_and_returned(self):
        with tempfile.TemporaryDirectory() as directory:
            data_path = Path(directory) / "sales_email_data.json"
            with patch.object(dm_assistant, "SALES_EMAIL_DATA_PATH", data_path):
                result = dm_assistant.update_sales_email_data(
                    {
                        "action": "save_contact",
                        "name": "김담당",
                        "company": "테스트몰",
                        "email": "buyer@example.com",
                        "group": "8월 제안",
                        "memo": "",
                    }
                )
                self.assertEqual(result["contact"]["group"], "8월 제안")
                public = dm_assistant.public_sales_email_data()
                self.assertEqual(public["contacts"][0]["group"], "8월 제안")

    def test_outlook_draft_is_launched_without_waiting(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            data_path = root / "sales_email_data.json"
            attachment_path = root / "attachments"
            script_path = root / "open_draft.ps1"
            script_path.write_text("param([string]$DraftPath)", encoding="utf-8")
            data_path.write_text(
                json.dumps(
                    {
                        "contacts": [
                            {
                                "id": "contact-1",
                                "name": "Test",
                                "email": "test@example.com",
                            }
                        ],
                        "templates": [],
                        "drafts": {},
                        "history": [],
                    }
                ),
                encoding="utf-8",
            )
            with patch.object(dm_assistant, "SALES_EMAIL_DATA_PATH", data_path), patch.object(
                dm_assistant, "SALES_EMAIL_ATTACHMENTS_DIR", attachment_path
            ), patch.object(dm_assistant, "APP_DIR", root), patch.object(
                dm_assistant, "OUTLOOK_POPUP_SCRIPT_PATH", script_path
            ), patch.object(dm_assistant.subprocess, "Popen") as popen:
                result = dm_assistant.create_outlook_draft(
                    {
                        "contact_id": "contact-1",
                        "subject": "Subject",
                        "full_body": "Body",
                        "full_html": "<p>Body</p>",
                    }
                )

            self.assertEqual(result["status"], "opening")
            popen.assert_called_once()
            kwargs = popen.call_args.kwargs
            self.assertEqual(kwargs["stdout"], subprocess.DEVNULL)
            self.assertNotIn("timeout", kwargs)


if __name__ == "__main__":
    unittest.main()
