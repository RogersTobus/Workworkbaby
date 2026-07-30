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


if __name__ == "__main__":
    unittest.main()
