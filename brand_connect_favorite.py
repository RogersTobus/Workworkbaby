from __future__ import annotations

import os
import shutil
import subprocess
import threading
import time
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from brand_connect_crawler import BRAND_CONNECT_URL


PLATFORM_LABELS = {
    "blog": "블로그",
    "instagram": "인스타그램",
}
BRAND_LABELS = {
    "alp": "알프",
    "gaia": "가이아",
}
PRODUCT_LABELS = {
    "alp": {
        "immun": "알프이뮨",
        "iron_drop": "알프아이언드롭",
    },
    "gaia": {
        "oil": "오일",
        "pickles": "병절임",
        "pouch": "파우치",
    },
}
RUNNING_STATUSES = {"queued", "preparing", "running", "saving", "login_required"}


class BrandConnectFavoriteManager:
    def __init__(
        self,
        app_dir: Path,
        prepare_callback: Callable[[str, str, str, int], list[dict[str, Any]]],
        persist_callback: Callable[[str, str, list[dict[str, Any]]], dict[str, Any]],
    ):
        self.app_dir = Path(app_dir)
        self.prepare_callback = prepare_callback
        self.persist_callback = persist_callback
        self.profile_dir = self.app_dir / "app_data" / "brand_connect_chrome"
        self.debug_port = 9231
        self.lock = threading.Lock()
        self.stop_event = threading.Event()
        self.resume_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.state: dict[str, Any] = {
            "status": "idle",
            "message": "찜 자동화 대기 중",
            "brand": "",
            "platform": "",
            "product": "",
            "product_label": "",
            "requested": 0,
            "available": 0,
            "processed": 0,
            "failed": 0,
            "marked": 0,
            "recorded": 0,
            "dated": 0,
            "date_preserved": 0,
            "group_name": "",
            "current_creator": "",
            "started_at": None,
            "finished_at": None,
        }

    def _set(self, **values: Any) -> None:
        with self.lock:
            self.state.update(values)

    def snapshot(self) -> dict[str, Any]:
        with self.lock:
            return dict(self.state)

    def is_running(self) -> bool:
        return self.snapshot().get("status") in RUNNING_STATUSES

    def start(
        self,
        brand: str,
        platform: str,
        product: str,
        count: int,
    ) -> dict[str, Any]:
        if brand not in BRAND_LABELS:
            raise ValueError("지원하지 않는 브랜드입니다.")
        if platform not in PLATFORM_LABELS:
            raise ValueError("블로그 또는 인스타그램을 선택해주세요.")
        product_label = PRODUCT_LABELS.get(brand, {}).get(product)
        if not product_label:
            raise ValueError("선택한 브랜드에 맞는 상품을 선택해주세요.")
        count = int(count)
        if count not in {10, 20, 30, 40, 50}:
            raise ValueError("찜 인원은 10명 단위로 10명부터 50명까지 선택해주세요.")
        with self.lock:
            if self.worker and self.worker.is_alive():
                raise ValueError("이미 찜 자동화가 진행 중입니다.")
            self.stop_event.clear()
            self.resume_event.clear()
            group_name = (
                f"{BRAND_LABELS[brand]} {PLATFORM_LABELS[platform]} "
                f"({datetime.now():%y.%m.%d})"
            )
            self.state.update(
                status="queued",
                message="찜 후보를 준비 중입니다.",
                brand=brand,
                platform=platform,
                product=product,
                product_label=product_label,
                requested=count,
                available=0,
                processed=0,
                failed=0,
                marked=0,
                recorded=0,
                dated=0,
                date_preserved=0,
                group_name=group_name,
                current_creator="",
                started_at=datetime.now().isoformat(timespec="seconds"),
                finished_at=None,
            )
            self.worker = threading.Thread(
                target=self._run,
                args=(brand, platform, product, product_label, count, group_name),
                daemon=True,
                name="brand-connect-favorite",
            )
            self.worker.start()
        return self.snapshot()

    def resume_after_login(self) -> dict[str, Any]:
        if self.snapshot().get("status") != "login_required":
            raise ValueError("현재 로그인 대기 상태가 아닙니다.")
        self.resume_event.set()
        self._set(status="running", message="로그인 확인 후 찜 자동화를 계속합니다.")
        return self.snapshot()

    def stop(self) -> dict[str, Any]:
        self.stop_event.set()
        self.resume_event.set()
        self._set(message="중단 요청을 처리 중입니다.")
        return self.snapshot()

    def _chrome_path(self) -> Path:
        candidates = [
            Path(os.environ.get("PROGRAMFILES", ""))
            / "Google/Chrome/Application/chrome.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", ""))
            / "Google/Chrome/Application/chrome.exe",
            Path(os.environ.get("LOCALAPPDATA", ""))
            / "Google/Chrome/Application/chrome.exe",
        ]
        for path in candidates:
            if path.is_file():
                return path
        discovered = shutil.which("chrome") or shutil.which("chrome.exe")
        if discovered:
            return Path(discovered)
        raise RuntimeError("Google Chrome을 찾을 수 없습니다.")

    def _debug_ready(self) -> bool:
        try:
            with urllib.request.urlopen(
                f"http://127.0.0.1:{self.debug_port}/json/version",
                timeout=1,
            ) as response:
                return response.status == 200
        except Exception:
            return False

    def _open_debug_chrome(self) -> None:
        if self._debug_ready():
            return
        self.profile_dir.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(
            [
                str(self._chrome_path()),
                f"--remote-debugging-port={self.debug_port}",
                f"--user-data-dir={self.profile_dir}",
                "--no-first-run",
                "--no-default-browser-check",
                BRAND_CONNECT_URL,
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        deadline = time.time() + 20
        while time.time() < deadline:
            if self._debug_ready():
                return
            if self.stop_event.wait(0.25):
                return
        raise RuntimeError("Chrome 디버깅 창에 연결하지 못했습니다.")

    @staticmethod
    def _page_is_ready(page: Any) -> bool:
        try:
            return page.get_by_text("크리에이터 찾기", exact=True).count() > 0
        except Exception:
            return False

    def _wait_for_login(self, page: Any) -> bool:
        if self._page_is_ready(page):
            return True
        self._set(
            status="login_required",
            message=(
                "열린 Chrome에서 네이버 브랜드커넥트 로그인을 완료한 뒤 "
                "'로그인 완료·계속'을 눌러주세요."
            ),
        )
        self.resume_event.clear()
        self.resume_event.wait()
        if self.stop_event.is_set():
            return False
        page.goto(BRAND_CONNECT_URL, wait_until="domcontentloaded", timeout=60_000)
        page.wait_for_timeout(1500)
        return self._page_is_ready(page)

    @staticmethod
    def _close_popup(page: Any) -> None:
        close = page.get_by_role("button", name="팝업 닫기", exact=True)
        if close.count():
            close.first.click(timeout=5000)
            page.wait_for_timeout(250)

    @staticmethod
    def _select_platform(page: Any, platform: str) -> None:
        label = PLATFORM_LABELS[platform]
        tab = page.get_by_text(label, exact=True)
        if not tab.count():
            raise RuntimeError(f"{label} 탭을 찾지 못했습니다.")
        tab.first.click(timeout=10_000)
        page.wait_for_timeout(700)

    @staticmethod
    def _search_creator(page: Any, creator: str) -> None:
        search = page.get_by_placeholder(
            "브랜드 커넥트 스페이스명 혹은 채널명을 입력해 주세요."
        )
        search.fill("")
        search.fill(creator)
        page.wait_for_timeout(900)
        suggestion = page.locator(
            '[class*="CreatorSearchInfluencerItem_item"]'
        ).filter(has_text=creator)
        if not suggestion.count():
            raise LookupError("검색 결과에서 크리에이터를 찾지 못했습니다.")
        suggestion.first.click(timeout=5000)
        page.wait_for_timeout(850)

    @staticmethod
    def _open_favorite_popup(page: Any, creator: str) -> None:
        card = page.locator('li[class*="CreatorItem_root"]').filter(
            has_text=creator
        )
        if not card.count():
            raise RuntimeError("검색 결과 카드에서 크리에이터를 확인하지 못했습니다.")
        favorite = card.first.locator(
            'button[aria-label="찜하기"],button[aria-label="찜하기 해제"]'
        )
        if not favorite.count():
            raise RuntimeError("크리에이터 찜 버튼을 찾지 못했습니다.")
        favorite.first.click(timeout=5000)
        page.get_by_text("크리에이터 찜", exact=True).wait_for(
            state="visible",
            timeout=5000,
        )

    @staticmethod
    def _ensure_group(page: Any, group_name: str) -> None:
        title = page.get_by_text("크리에이터 찜", exact=True)
        if not title.count():
            raise RuntimeError("크리에이터 찜 창을 열지 못했습니다.")
        group = page.get_by_text(group_name, exact=True)
        if not group.count():
            page.get_by_role(
                "button",
                name="그룹 추가/편집",
                exact=True,
            ).click(timeout=5000)
            group_input = page.get_by_placeholder("그룹명을 입력해 주세요.")
            group_input.fill(group_name)
            page.get_by_role("button", name="추가하기", exact=True).click(
                timeout=5000
            )
            page.get_by_text(group_name, exact=True).wait_for(
                state="visible",
                timeout=5000,
            )
            page.get_by_role(
                "button",
                name="추가/편집 완료",
                exact=True,
            ).click(timeout=5000)
            page.wait_for_timeout(350)
        group = page.get_by_text(group_name, exact=True).first
        group_item = group.locator("xpath=ancestor::li[1]")
        checkbox = group_item.locator('input[type="checkbox"]')
        if checkbox.count() != 1:
            raise RuntimeError("찜 그룹 선택 원을 찾지 못했습니다.")
        if not checkbox.is_checked():
            checkbox.check(force=True)
        page.get_by_role("button", name="저장하기", exact=True).click(
            timeout=5000
        )
        title.wait_for(state="hidden", timeout=5000)
        page.wait_for_timeout(300)

    def _favorite_creator(
        self,
        page: Any,
        creator: str,
        group_name: str,
    ) -> None:
        self._close_popup(page)
        self._search_creator(page, creator)
        self._open_favorite_popup(page, creator)
        self._ensure_group(page, group_name)

    def _run(
        self,
        brand: str,
        platform: str,
        product: str,
        product_label: str,
        count: int,
        group_name: str,
    ) -> None:
        playwright = None
        browser = None
        completed: list[dict[str, Any]] = []
        terminal_error: Exception | None = None
        try:
            self._set(
                status="preparing",
                message="최신 시트에서 아직 선택되지 않은 후보를 확인합니다.",
            )
            candidates = self.prepare_callback(brand, platform, product, count)
            self._set(available=len(candidates))
            if not candidates:
                raise RuntimeError(
                    f"{PLATFORM_LABELS[platform]} · {product_label}의 빈칸 후보가 없습니다."
                )
            self._open_debug_chrome()
            from playwright.sync_api import sync_playwright

            playwright = sync_playwright().start()
            browser = playwright.chromium.connect_over_cdp(
                f"http://127.0.0.1:{self.debug_port}"
            )
            context = browser.contexts[0]
            page = next(
                (item for item in context.pages if "brandconnect.naver.com" in item.url),
                None,
            )
            if page is None:
                page = context.new_page()
                page.goto(BRAND_CONNECT_URL, wait_until="domcontentloaded", timeout=60_000)
            for other_page in list(context.pages):
                if other_page != page:
                    try:
                        other_page.close()
                    except Exception:
                        pass

            def close_unwanted_page(opened_page: Any) -> None:
                try:
                    if opened_page != page:
                        opened_page.close()
                except Exception:
                    pass

            context.on("page", close_unwanted_page)
            if not self._wait_for_login(page):
                raise RuntimeError("네이버 브랜드커넥트 로그인을 확인하지 못했습니다.")
            self._close_popup(page)
            self._select_platform(page, platform)
            self._set(
                status="running",
                message=f"{group_name} 그룹에 크리에이터를 추가합니다.",
            )
            for candidate in candidates:
                if self.stop_event.is_set():
                    break
                creator = str(candidate.get("creator", "")).strip()
                self._set(current_creator=creator)
                try:
                    self._favorite_creator(page, creator, group_name)
                    completed.append(candidate)
                    self._set(
                        processed=len(completed),
                        message=f"{creator} 찜 그룹 추가 완료",
                    )
                except Exception as exc:
                    self._close_popup(page)
                    self._set(
                        failed=self.snapshot()["failed"] + 1,
                        message=f"{creator} 건너뜀: {exc}",
                    )
            self._set(
                status="saving",
                message="완료된 명단을 앱 내부 찜 이력에 저장합니다.",
                current_creator="",
            )
        except Exception as exc:
            terminal_error = exc
        finally:
            try:
                self._close_popup(
                    next(
                        item
                        for item in browser.contexts[0].pages
                        if "brandconnect.naver.com" in item.url
                    )
                ) if browser is not None else None
            except Exception:
                pass
            try:
                if browser is not None:
                    browser.close()
            except Exception:
                pass
            try:
                if playwright is not None:
                    playwright.stop()
            except Exception:
                pass

        result: dict[str, Any] = {
            "marked": 0,
            "recorded": 0,
            "dated": 0,
            "date_preserved": 0,
            "drive": {"status": "skipped"},
        }
        if completed:
            try:
                result = self.persist_callback(brand, product, completed)
                self._set(
                    marked=int(result.get("marked", 0) or 0),
                    recorded=int(result.get("recorded", 0) or 0),
                    dated=int(result.get("dated", 0) or 0),
                    date_preserved=int(result.get("date_preserved", 0) or 0),
                )
            except Exception as exc:
                terminal_error = terminal_error or exc

        finished_at = datetime.now().isoformat(timespec="seconds")
        drive = dict(result.get("drive") or {})
        drive_ok = drive.get("status") in {"completed", "skipped"}
        if self.stop_event.is_set():
            self._set(
                status="stopped" if drive_ok else "completed_with_warning",
                message=(
                    f"중단했습니다. 완료된 {len(completed)}명은 시트에 반영했습니다."
                ),
                current_creator="",
                finished_at=finished_at,
            )
        elif terminal_error is not None:
            self._set(
                status="failed",
                message=f"찜 자동화 실패: {terminal_error}",
                current_creator="",
                finished_at=finished_at,
            )
        else:
            requested_message = (
                f"요청 {count}명 중 {len(completed)}명을 {group_name} 그룹에 추가했습니다."
            )
            requested_message += (
                " Google Sheet에는 찜 표시를 남기지 않았습니다."
            )
            if not drive_ok:
                requested_message += f" {drive.get('message', 'Drive 최신화 확인이 필요합니다.')}"
            self._set(
                status="completed" if drive_ok else "completed_with_warning",
                message=requested_message,
                current_creator="",
                finished_at=finished_at,
            )
