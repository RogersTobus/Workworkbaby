from __future__ import annotations

import math
import os
import re
import shutil
import subprocess
import threading
import time
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Callable


BRAND_CONNECT_URL = (
    "https://brandconnect.naver.com/852209558043936/creator/search"
)


def normalize_creator(value: object) -> str:
    return re.sub(r"\s+", "", str(value or "")).casefold()


def compact_number(value: str) -> int:
    text = str(value or "").replace(",", "").replace(" ", "")
    match = re.search(r"(\d+(?:\.\d+)?)\s*(만|천)?", text)
    if not match:
        return 0
    number = float(match.group(1))
    multiplier = {"만": 10000, "천": 1000}.get(match.group(2), 1)
    return int(number * multiplier)


class BrandConnectCrawlerManager:
    def __init__(
        self,
        app_dir: Path,
        prepare_callback: Callable[[str], set[str]],
        persist_callback: Callable[[str, list[dict[str, Any]]], dict[str, Any]],
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
            "message": "크롤링 대기 중",
            "brand": "",
            "platforms": [],
            "requested": 0,
            "collected": 0,
            "added": 0,
            "duplicates": 0,
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

    def start(self, brand: str, platforms: list[str], count: int) -> dict[str, Any]:
        if brand not in {"alp", "gaia"}:
            raise ValueError("지원하지 않는 브랜드입니다.")
        clean_platforms = [
            item for item in ("blog", "instagram") if item in set(platforms)
        ]
        if not clean_platforms:
            raise ValueError("블로그 또는 인스타그램을 선택해주세요.")
        count = max(1, min(int(count), 200))
        with self.lock:
            if self.worker and self.worker.is_alive():
                raise ValueError("이미 브랜드 커넥팅 크롤링이 진행 중입니다.")
            self.stop_event.clear()
            self.resume_event.clear()
            self.state.update(
                status="queued",
                message="크롤링 준비 중",
                brand=brand,
                platforms=clean_platforms,
                requested=count,
                collected=0,
                added=0,
                duplicates=0,
                current_creator="",
                started_at=datetime.now().isoformat(timespec="seconds"),
                finished_at=None,
            )
            self.worker = threading.Thread(
                target=self._run,
                args=(brand, clean_platforms, count),
                daemon=True,
                name="brand-connect-crawler",
            )
            self.worker.start()
        return self.snapshot()

    def resume_after_login(self) -> dict[str, Any]:
        if self.snapshot().get("status") != "login_required":
            raise ValueError("현재 로그인 대기 상태가 아닙니다.")
        self.resume_event.set()
        self._set(status="running", message="로그인 확인 후 크롤링을 계속합니다.")
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
                "열린 Chrome에서 네이버 브랜드 커넥트 로그인을 완료한 뒤 "
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
    def _click_text(page: Any, text: str) -> bool:
        locator = page.get_by_text(text, exact=True)
        if locator.count() < 1:
            return False
        locator.first.click(timeout=10_000)
        page.wait_for_timeout(350)
        return True

    def _select_filters(self, page: Any, brand: str, platform: str) -> None:
        platform_text = "블로그" if platform == "blog" else "인스타그램"
        if not self._click_text(page, platform_text):
            raise RuntimeError(f"{platform_text} 탭을 찾지 못했습니다.")
        page.wait_for_timeout(900)

        reset_button = page.get_by_role(
            "button",
            name="검색 조건 초기화",
            exact=True,
        )
        if reset_button.count():
            reset_button.first.click()
            page.wait_for_timeout(600)

        activity = page.get_by_text(re.compile(r"^활동 주제"))
        if activity.count():
            activity.first.click()
            page.wait_for_timeout(250)
            topics = ["푸드"] if brand == "gaia" else ["생활건강", "운동/레저"]
            for topic in topics:
                option = page.get_by_text(topic, exact=True)
                if option.count():
                    option.first.click()
                    page.wait_for_timeout(180)
            page.keyboard.press("Escape")

        experience = page.get_by_text(re.compile(r"^제휴 경험"))
        if experience.count():
            experience.first.click()
            page.wait_for_timeout(250)
            option = page.get_by_text("콘텐츠 제휴 경험", exact=True)
            if option.count():
                option.first.click()
            page.keyboard.press("Escape")

        sort_text = (
            "일평균 방문 많은 순"
            if platform == "blog"
            else "팔로워 많은 순"
        )
        sort_option = page.get_by_text(sort_text, exact=True)
        if sort_option.count():
            sort_option.first.click()
        page.wait_for_timeout(1200)

    @staticmethod
    def _open_detail_from_button(button: Any) -> str:
        card = button.locator("xpath=ancestor::li[1]")
        card_text = str(card.inner_text(timeout=3000) or "")
        detail_button = card.get_by_role(
            "button",
            name="크리에이터 상세 보기",
            exact=True,
        )
        if detail_button.count() != 1:
            raise RuntimeError("크리에이터 상세 보기 버튼을 찾지 못했습니다.")
        detail_button.click(timeout=5000)
        return card_text

    @staticmethod
    def _detail_text(page: Any) -> str:
        page.wait_for_timeout(700)
        marker = page.get_by_text("스페이스 연동 채널", exact=True)
        if marker.count() < 1:
            return ""
        return marker.first.evaluate(
            """element => {
              let node = element;
              let best = '';
              while (node && node !== document.body) {
                const text = (node.innerText || '').trim();
                if (text.includes('스페이스 연동 채널') && text.length > best.length && text.length < 9000) best = text;
                if (text.includes('콘텐츠 제휴 진행 이력')) return text;
                node = node.parentElement;
              }
              return best;
            }"""
        )

    @staticmethod
    def _close_detail(page: Any) -> None:
        close_button = page.get_by_role("button", name="팝업 닫기", exact=True)
        if close_button.count():
            close_button.first.click(timeout=5000)
            page.wait_for_timeout(250)

    @staticmethod
    def _parse_candidate(
        detail_text: str,
        card_text: str,
        platform: str,
    ) -> dict[str, Any] | None:
        lines = [line.strip() for line in detail_text.splitlines() if line.strip()]
        if not lines:
            return None
        ignored = {
            "스페이스 연동 채널",
            "제휴 진행 이력",
            "콘텐츠 제휴",
            "공동구매",
            "제안하기",
        }
        creator = next(
            (
                line
                for line in lines
                if line not in ignored
                and not line.startswith(("blog", "팔로워", "이웃", "구독"))
                and "|" not in line
            ),
            "",
        )
        card_lines = [line.strip() for line in card_text.splitlines() if line.strip()]
        topics = ", ".join(
            topic
            for topic in ("생활건강", "운동/레저", "푸드")
            if topic in card_lines
        )
        history_match = re.search(
            r"콘텐츠 제휴 진행 이력\s*총\s*([\d,]+)\s*건",
            detail_text,
        )
        history_count = int(history_match.group(1).replace(",", "")) if history_match else 0
        if platform == "blog":
            audience_match = re.search(
                r"일평균 방문(?:\s*수)?\s*([\d,.만천]+)",
                card_text,
            ) or re.search(r"이웃\s*([\d,.만천]+)", detail_text)
        else:
            audience_match = re.search(
                r"팔로워(?:\s*수)?\s*([\d,.만천]+)",
                card_text + "\n" + detail_text,
            )
        audience_count = compact_number(audience_match.group(1)) if audience_match else 0
        if not creator:
            return None
        return {
            "creator": creator,
            "topics": topics,
            "platform": "블로그" if platform == "blog" else "인스타그램",
            "history_count": history_count,
            "audience_count": audience_count,
        }

    def _crawl_platform(
        self,
        page: Any,
        brand: str,
        platform: str,
        target: int,
        known: set[str],
        collected: list[dict[str, Any]],
    ) -> int:
        self._select_filters(page, brand, platform)
        duplicates = 0
        stalled_rounds = 0
        detail_failures = 0
        seen_cards: set[str] = set()
        self._set(
            message=(
                f"{'블로그' if platform == 'blog' else '인스타그램'} "
                "크리에이터 상세 정보를 확인합니다."
            )
        )
        while len(collected) < target and not self.stop_event.is_set():
            buttons = page.get_by_role("button", name="제안하기")
            count = buttons.count()
            progress = False
            for index in range(count):
                if len(collected) >= target or self.stop_event.is_set():
                    break
                button = buttons.nth(index)
                try:
                    card_text = str(
                        button.locator("xpath=ancestor::li[1]").inner_text(timeout=3000)
                        or ""
                    )
                    card_key = f"{platform}:{index}:{card_text[:120]}"
                    if card_key in seen_cards:
                        continue
                    seen_cards.add(card_key)
                    card_text = self._open_detail_from_button(button)
                    detail_text = self._detail_text(page)
                    if not detail_text:
                        detail_failures += 1
                        if detail_failures >= 5:
                            raise RuntimeError(
                                "브랜드 커넥트 상세 팝업을 열지 못했습니다."
                            )
                        continue
                    detail_failures = 0
                    candidate = self._parse_candidate(detail_text, card_text, platform)
                    self._close_detail(page)
                    if not candidate:
                        continue
                    creator_key = normalize_creator(candidate["creator"])
                    self._set(current_creator=candidate["creator"])
                    if creator_key in known:
                        duplicates += 1
                        continue
                    known.add(creator_key)
                    collected.append(candidate)
                    self._set(
                        collected=len(collected),
                        duplicates=self.snapshot()["duplicates"] + duplicates,
                        message=f"{candidate['creator']} 정보 수집 완료",
                    )
                    duplicates = 0
                    progress = True
                except Exception:
                    try:
                        self._close_detail(page)
                    except Exception:
                        pass
                    if detail_failures >= 5:
                        raise RuntimeError(
                            "브랜드 커넥트 상세 팝업을 열지 못해 안전하게 중단했습니다."
                        )
                    continue
            if progress:
                stalled_rounds = 0
            else:
                stalled_rounds += 1
            if stalled_rounds >= 4:
                break
            page.mouse.wheel(0, 1600)
            page.wait_for_timeout(900)
        return duplicates

    def _run(self, brand: str, platforms: list[str], count: int) -> None:
        playwright = None
        browser = None
        try:
            self._set(status="preparing", message="기존 시트 명단과 최신 원본을 확인합니다.")
            existing = self.prepare_callback(brand)
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
                raise RuntimeError("네이버 브랜드 커넥트 로그인을 확인하지 못했습니다.")
            self._close_detail(page)

            self._set(status="running", message="크리에이터 검색 조건을 설정합니다.")
            collected: list[dict[str, Any]] = []
            known = set(existing)
            remaining_platforms = len(platforms)
            for platform in platforms:
                if self.stop_event.is_set() or len(collected) >= count:
                    break
                remaining = count - len(collected)
                platform_goal = len(collected) + math.ceil(
                    remaining / remaining_platforms
                )
                leftovers = self._crawl_platform(
                    page,
                    brand,
                    platform,
                    platform_goal,
                    known,
                    collected,
                )
                if leftovers:
                    self._set(duplicates=self.snapshot()["duplicates"] + leftovers)
                remaining_platforms -= 1

            if self.stop_event.is_set():
                self._set(
                    status="stopped",
                    message="크롤링을 중단했습니다. 시트에는 추가하지 않았습니다.",
                    finished_at=datetime.now().isoformat(timespec="seconds"),
                )
                return
            self._set(status="saving", message="중복을 제외한 새 명단을 시트에 저장합니다.")
            result = self.persist_callback(brand, collected)
            drive = dict(result.get("drive") or {})
            drive_ok = drive.get("status") in {"completed", "skipped"}
            if drive_ok:
                message = (
                    f"새 크리에이터 {result.get('added', 0)}명을 "
                    f"{result.get('sheet_name', '')} 시트에 추가했습니다."
                )
            else:
                message = (
                    f"로컬에 {result.get('added', 0)}명을 저장했습니다. "
                    f"{drive.get('message', 'Drive 최신화 확인이 필요합니다.')}"
                )
            self._set(
                status="completed" if drive_ok else "completed_with_warning",
                message=message,
                added=result.get("added", 0),
                duplicates=self.snapshot()["duplicates"]
                + result.get("duplicates", 0),
                finished_at=datetime.now().isoformat(timespec="seconds"),
                current_creator="",
            )
        except Exception as exc:
            self._set(
                status="failed",
                message=f"크롤링 실패: {exc}",
                finished_at=datetime.now().isoformat(timespec="seconds"),
            )
        finally:
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
