import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

from brand_connect_crawler import BrandConnectCrawlerManager, normalize_creator


class _Page:
    url = "https://brandconnect.naver.com/creator/search"


class _Context:
    def __init__(self):
        self.pages = [_Page()]

    def on(self, *_args):
        return None


class _Browser:
    def __init__(self):
        self.contexts = [_Context()]

    def close(self):
        return None


class _Playwright:
    def __init__(self):
        self.chromium = types.SimpleNamespace(
            connect_over_cdp=lambda _url: _Browser()
        )

    def stop(self):
        return None


class _SyncPlaywright:
    def start(self):
        return _Playwright()


class _Manager(BrandConnectCrawlerManager):
    def __init__(self, persist_callback, crawl_batches):
        super().__init__(
            Path(tempfile.gettempdir()),
            prepare_callback=lambda _brand: set(),
            persist_callback=persist_callback,
        )
        self.crawl_batches = list(crawl_batches)
        self.crawl_calls = 0

    def _open_debug_chrome(self):
        return None

    def _wait_for_login(self, _page):
        return True

    def _close_detail(self, _page):
        return None

    def _crawl_platform(
        self,
        _page,
        _brand,
        _platform,
        target,
        known,
        collected,
        collected_offset=0,
    ):
        self.crawl_calls += 1
        batch = self.crawl_batches.pop(0) if self.crawl_batches else []
        for creator in batch:
            if len(collected) >= target:
                break
            key = normalize_creator(creator)
            if key in known:
                continue
            known.add(key)
            collected.append({"creator": creator})
        self._set(collected=collected_offset + len(collected))
        return 0


class BrandConnectCrawlerTargetTests(unittest.TestCase):
    def _playwright_modules(self):
        package = types.ModuleType("playwright")
        sync_api = types.ModuleType("playwright.sync_api")
        sync_api.sync_playwright = lambda: _SyncPlaywright()
        return patch.dict(
            sys.modules,
            {"playwright": package, "playwright.sync_api": sync_api},
        )

    def test_keeps_crawling_until_sheet_added_count_reaches_request(self):
        save_calls = []

        def persist(_brand, candidates):
            save_calls.append([item["creator"] for item in candidates])
            added = 2 if len(save_calls) == 1 else len(candidates)
            return {
                "added": added,
                "duplicates": len(candidates) - added,
                "drive": {"status": "completed"},
            }

        manager = _Manager(persist, [["a", "b", "c"], ["d"]])
        with self._playwright_modules():
            manager._run("gaia", ["blog"], 3)

        state = manager.snapshot()
        self.assertEqual("completed", state["status"])
        self.assertEqual(3, state["added"])
        self.assertEqual(2, manager.crawl_calls)
        self.assertEqual([["a", "b", "c"], ["d"]], save_calls)

    def test_reports_warning_when_no_more_new_creators_exist(self):
        manager = _Manager(lambda _brand, _items: {}, [[]])
        with self._playwright_modules():
            manager._run("gaia", ["blog"], 100)

        state = manager.snapshot()
        self.assertEqual("completed_with_warning", state["status"])
        self.assertEqual(0, state["added"])


if __name__ == "__main__":
    unittest.main()
