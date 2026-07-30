import tempfile
import unittest
import sys
import types
from datetime import date
from pathlib import Path
from unittest.mock import patch

simulation_manager = types.ModuleType("simulation_manager")


class SimulationManager:
    def __init__(self, *args, **kwargs):
        pass


simulation_manager.SimulationManager = SimulationManager
sys.modules.setdefault("simulation_manager", simulation_manager)

import dm_assistant


class ChecklistRoutineTests(unittest.TestCase):
    def test_daily_period_changes_next_day(self):
        item = {"recurrence": "daily"}
        self.assertEqual(
            dm_assistant.checklist_routine_period_key(item, date(2026, 7, 30)),
            "2026-07-30",
        )
        self.assertEqual(
            dm_assistant.checklist_routine_period_key(item, date(2026, 7, 31)),
            "2026-07-31",
        )

    def test_weekly_period_uses_monday_and_changes_next_week(self):
        item = {"recurrence": "weekly"}
        self.assertEqual(
            dm_assistant.checklist_routine_period_key(item, date(2026, 7, 30)),
            "2026-07-27",
        )
        self.assertEqual(
            dm_assistant.checklist_routine_period_key(item, date(2026, 8, 3)),
            "2026-08-03",
        )

    def test_add_toggle_and_delete_checklist_routine(self):
        with tempfile.TemporaryDirectory() as directory:
            task_path = Path(directory) / "calendar_tasks.json"
            with patch.object(dm_assistant, "CALENDAR_TASKS_PATH", task_path):
                added = dm_assistant.update_calendar_checklist_routine(
                    {
                        "action": "add",
                        "text": "오전 메일 확인",
                        "recurrence": "daily",
                    }
                )["routine"]
                routine_id = added["id"]
                toggled = dm_assistant.update_calendar_checklist_routine(
                    {"action": "toggle", "id": routine_id, "completed": True}
                )["routine"]
                self.assertTrue(toggled["completed"])
                listed = dm_assistant.get_calendar_checklist_routines()["routines"]
                self.assertEqual(len(listed), 1)
                self.assertTrue(listed[0]["completed"])
                dm_assistant.update_calendar_checklist_routine(
                    {"action": "delete", "id": routine_id}
                )
                self.assertEqual(
                    dm_assistant.get_calendar_checklist_routines()["routines"], []
                )


if __name__ == "__main__":
    unittest.main()
