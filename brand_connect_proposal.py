from __future__ import annotations

import os
import re
import shutil
import subprocess
import threading
import time
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Callable


BRAND_CONNECT_ORIGIN = "https://brandconnect.naver.com"
BRAND_LABELS = {"alp": "알프", "gaia": "가이아"}
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
ACCEPTED_CAMPAIGN_STATUS_KEYWORDS = (
    "정산/확정필요",
    "진행완료",
    "진행중",
    "정산",
    "확정",
    "수락",
    "선정",
)


def normalize_proposal_date(value: object) -> str:
    text = str(value or "").strip().replace("-", ".")
    try:
        parsed = datetime.strptime(text, "%Y.%m.%d")
    except ValueError as exc:
        raise ValueError("제안날짜는 yyyy.mm.dd 형식으로 입력해주세요.") from exc
    return parsed.strftime("%Y.%m.%d")


def normalize_campaign_status(value: object) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    compact = text.replace(" ", "")
    if "거절" in compact:
        return "거절"
    if "수락대기" in compact or compact == "대기":
        return "대기"
    if any(keyword in compact for keyword in ACCEPTED_CAMPAIGN_STATUS_KEYWORDS):
        return "수락"
    return ""


def count_unique_creators(results: list[dict[str, Any]]) -> int:
    return len(
        {
            re.sub(r"\s+", " ", str(item.get("creator", ""))).strip().casefold()
            for item in results
            if str(item.get("creator", "")).strip()
        }
    )


def validate_campaign_url(value: object) -> str:
    url = str(value or "").strip()
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme != "https" or parsed.netloc != "brandconnect.naver.com":
        raise ValueError("네이버 브랜드커넥트 캠페인 링크를 입력해주세요.")
    if not re.fullmatch(r"/\d+/campaign/\d+/progress/?", parsed.path):
        raise ValueError("캠페인의 진행 현황(progress) 링크를 입력해주세요.")
    return urllib.parse.urlunparse(
        (parsed.scheme, parsed.netloc, parsed.path.rstrip("/"), "", "", "")
    )


def content_url_score(value: object) -> int:
    """Return a preference score for a submitted creator-content URL."""
    url = str(value or "").strip()
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme not in {"http", "https"}:
        return 0
    host = parsed.netloc.lower().split(":", 1)[0]
    path = parsed.path.lower()
    if host in {"instagram.com", "www.instagram.com"}:
        return 100 if re.search(r"/(?:p|reel|tv)/", path) else 80
    if host in {"blog.naver.com", "m.blog.naver.com"}:
        if path.strip("/").split("/", 1)[0] == "brandconnect":
            return 0
        return 95
    if host in {"youtube.com", "www.youtube.com", "youtu.be"}:
        return 90
    if host.endswith("naver.com") or host.endswith("navercorp.com"):
        return 0
    if host == "brandconnect.naver.com":
        return 0
    return 10


class BrandConnectProposalManager:
    def __init__(
        self,
        app_dir: Path,
        sync_callback: Callable[[], dict[str, Any]],
        persist_callback: Callable[
            [str, list[dict[str, Any]], list[dict[str, Any]]],
            dict[str, Any],
        ],
    ):
        self.app_dir = Path(app_dir)
        self.sync_callback = sync_callback
        self.persist_callback = persist_callback
        self.profile_dir = self.app_dir / "app_data" / "brand_connect_chrome"
        self.debug_port = 9231
        self.lock = threading.Lock()
        self.stop_event = threading.Event()
        self.resume_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.state: dict[str, Any] = {
            "status": "idle",
            "message": "제안 결과 확인 대기 중",
            "brand": "",
            "campaign_count": 0,
            "processed_campaigns": 0,
            "pages_checked": 0,
            "creators_seen": 0,
            "matched": 0,
            "updated": 0,
            "added": 0,
            "dated": 0,
            "reconciled": 0,
            "summary_rows": 0,
            "links": 0,
            "unmatched": 0,
            "failed": 0,
            "current_campaign": "",
            "errors": [],
            "started_at": None,
            "finished_at": None,
        }

    def _set(self, **values: Any) -> None:
        with self.lock:
            self.state.update(values)

    def snapshot(self) -> dict[str, Any]:
        with self.lock:
            result = dict(self.state)
            result["errors"] = list(result.get("errors") or [])
            return result

    def is_running(self) -> bool:
        return self.snapshot().get("status") in RUNNING_STATUSES

    def start(self, brand: str, campaigns: list[dict[str, Any]]) -> dict[str, Any]:
        if brand not in BRAND_LABELS:
            raise ValueError("지원하지 않는 브랜드입니다.")
        cleaned: list[dict[str, str]] = []
        seen_urls: set[tuple[str, str]] = set()
        for item in campaigns:
            url = validate_campaign_url(item.get("url"))
            product = str(item.get("product", "")).strip()
            if product not in PRODUCT_LABELS[brand]:
                raise ValueError("선택한 브랜드에 맞는 상품을 선택해주세요.")
            raw_proposal_date = str(item.get("proposal_date", "")).strip()
            proposal_date = (
                normalize_proposal_date(raw_proposal_date)
                if raw_proposal_date
                else ""
            )
            key = (url, product)
            if key in seen_urls:
                continue
            seen_urls.add(key)
            cleaned.append(
                {
                    "url": url,
                    "product": product,
                    "product_label": PRODUCT_LABELS[brand][product],
                    "proposal_date": proposal_date,
                }
            )
        if not cleaned:
            raise ValueError("확인할 캠페인 링크를 한 개 이상 입력해주세요.")
        if len(cleaned) > 20:
            raise ValueError("한 번에 최대 20개의 캠페인을 확인할 수 있습니다.")
        with self.lock:
            if self.worker and self.worker.is_alive():
                raise ValueError("이미 제안 결과 확인이 진행 중입니다.")
            self.stop_event.clear()
            self.resume_event.clear()
            self.state.update(
                status="queued",
                message="캠페인 확인을 준비 중입니다.",
                brand=brand,
                campaign_count=len(cleaned),
                processed_campaigns=0,
                pages_checked=0,
                creators_seen=0,
                matched=0,
                updated=0,
                added=0,
                dated=0,
                reconciled=0,
                summary_rows=0,
                links=0,
                unmatched=0,
                failed=0,
                current_campaign="",
                errors=[],
                started_at=time.strftime("%Y-%m-%dT%H:%M:%S"),
                finished_at=None,
            )
            self.worker = threading.Thread(
                target=self._run,
                args=(brand, cleaned),
                daemon=True,
                name="brand-connect-proposal",
            )
            self.worker.start()
        return self.snapshot()

    def resume_after_login(self) -> dict[str, Any]:
        if self.snapshot().get("status") != "login_required":
            raise ValueError("현재 로그인 대기 상태가 아닙니다.")
        self.resume_event.set()
        self._set(status="running", message="로그인 확인 후 제안 결과 확인을 계속합니다.")
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

    def _open_debug_chrome(self, initial_url: str) -> None:
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
                initial_url,
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
    def _login_required(page: Any) -> bool:
        return "nid.naver.com" in str(page.url) or "nidlogin" in str(page.url)

    def _navigate(self, page: Any, url: str) -> None:
        """Navigate while tolerating Naver's immediate login redirect."""
        recoverable_error: Exception | None = None
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60_000)
        except Exception as exc:
            message = str(exc)
            recoverable_navigation = (
                "interrupted by another navigation" in message
                or "ERR_ABORTED" in message
                or "Timeout" in message
            )
            if not recoverable_navigation:
                raise
            recoverable_error = exc

        # A login redirect can continue after goto() returns (or interrupts).
        # Wait until either the login page or the campaign table is observable.
        for _ in range(40):
            if self.stop_event.is_set():
                return
            if self._login_required(page):
                return
            try:
                if page.locator("table tbody tr").count() > 0:
                    return
            except Exception:
                pass
            page.wait_for_timeout(500)
        if recoverable_error is not None:
            raise recoverable_error

    def _wait_for_login(self, page: Any, url: str) -> bool:
        if not self._login_required(page):
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
        self._navigate(page, url)
        return not self._login_required(page)

    @staticmethod
    def _close_overlay(page: Any) -> None:
        for selector in (
            'button[aria-label*="닫기"]',
            'button[aria-label*="팝업"]',
        ):
            locator = page.locator(selector)
            for index in range(locator.count()):
                item = locator.nth(index)
                try:
                    if item.is_visible():
                        item.click(timeout=2500)
                        page.wait_for_timeout(200)
                        return
                except Exception:
                    continue
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass

    @staticmethod
    def _submitted_content_link(page: Any, row: Any) -> str:
        submitted_cells = row.locator("td").filter(has_text="제출")
        if submitted_cells.count() < 1:
            return ""
        cell = submitted_cells.first
        cell_text = re.sub(r"\s+", "", cell.inner_text(timeout=2500))
        if re.search(r"(?:0/\d+건|0건)제출", cell_text):
            return ""
        button = cell.get_by_role("button", name="조회/수정", exact=True)
        if button.count() < 1:
            button = cell.get_by_role("button", name="조회", exact=True)
        if button.count() < 1:
            return ""
        pages_before = list(page.context.pages)
        button.first.click(timeout=5000)
        page.wait_for_timeout(700)
        candidates: list[tuple[int, str]] = []

        def add_visible_links(scope: Any) -> None:
            try:
                anchors = scope.locator('a[href]').all()
            except Exception:
                return
            for anchor in anchors:
                try:
                    if not anchor.is_visible():
                        continue
                    href = str(anchor.get_attribute("href") or "").strip()
                    score = content_url_score(href)
                    if score and all(existing != href for _, existing in candidates):
                        candidates.append((score, href))
                except Exception:
                    continue

        overlay_selectors = (
            '[role="dialog"]:visible',
            '[class*="Modal"]:visible',
            '[class*="modal"]:visible',
            '[class*="Layer"]:visible',
            '[class*="layer"]:visible',
        )
        for selector in overlay_selectors:
            try:
                overlay = page.locator(selector)
                if overlay.count() < 1:
                    continue
                add_visible_links(overlay.last)
            except Exception:
                continue
        if not candidates:
            for current_page in page.context.pages:
                try:
                    add_visible_links(current_page)
                except Exception:
                    continue
        for current_page in list(page.context.pages):
            if current_page not in pages_before and current_page != page:
                try:
                    current_page.close()
                except Exception:
                    pass
        BrandConnectProposalManager._close_overlay(page)
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1] if candidates else ""

    def _read_current_page(
        self,
        page: Any,
        campaign: dict[str, str],
    ) -> list[dict[str, Any]]:
        rows = page.locator("table tbody tr")
        if rows.count() < 1:
            rows = page.locator("table tr")
        results: list[dict[str, Any]] = []
        for index in range(rows.count()):
            if self.stop_event.is_set():
                break
            row = rows.nth(index)
            cells = row.locator("td")
            if cells.count() < 2:
                continue
            status_text = cells.nth(0).inner_text(timeout=3000).strip()
            creator = cells.nth(1).inner_text(timeout=3000).strip()
            creator = re.sub(r"\s+", " ", creator).strip()
            status = normalize_campaign_status(status_text)
            if not creator or not status:
                continue
            content_url = ""
            try:
                content_url = self._submitted_content_link(page, row)
            except Exception:
                self._close_overlay(page)
            results.append(
                {
                    "creator": creator,
                    "status": status,
                    "raw_status": status_text,
                    "content_url": content_url,
                    "product": campaign["product"],
                    "product_label": campaign["product_label"],
                    "campaign_url": campaign["url"],
                    "proposal_date": campaign["proposal_date"],
                }
            )
        return results

    @staticmethod
    def _next_page(page: Any) -> Any | None:
        selectors = (
            'button[aria-label*="다음"]',
            'a[aria-label*="다음"]',
            'button:has-text("다음")',
            'a:has-text("다음")',
        )
        for selector in selectors:
            locator = page.locator(selector)
            for index in range(locator.count()):
                item = locator.nth(index)
                try:
                    if not item.is_visible():
                        continue
                    if item.get_attribute("disabled") is not None:
                        continue
                    if str(item.get_attribute("aria-disabled") or "").lower() == "true":
                        continue
                    return item
                except Exception:
                    continue
        return None

    def _read_campaign(
        self,
        page: Any,
        campaign: dict[str, str],
    ) -> tuple[list[dict[str, Any]], int]:
        self._navigate(page, campaign["url"])
        if not self._wait_for_login(page, campaign["url"]):
            raise RuntimeError("네이버 브랜드커넥트 로그인을 확인하지 못했습니다.")
        page.locator("table").first.wait_for(state="visible", timeout=15_000)
        all_results: list[dict[str, Any]] = []
        seen_pages: set[str] = set()
        pages_checked = 0
        for _ in range(100):
            if self.stop_event.is_set():
                break
            marker = page.url + "|" + str(
                page.locator("table tbody tr").first.inner_text(timeout=2500)
                if page.locator("table tbody tr").count()
                else ""
            )
            if marker in seen_pages:
                break
            seen_pages.add(marker)
            pages_checked += 1
            all_results.extend(self._read_current_page(page, campaign))
            next_button = self._next_page(page)
            if next_button is None:
                break
            previous_marker = marker
            next_button.click(timeout=5000)
            for _ in range(20):
                page.wait_for_timeout(300)
                current_marker = page.url + "|" + str(
                    page.locator("table tbody tr").first.inner_text(timeout=2500)
                    if page.locator("table tbody tr").count()
                    else ""
                )
                if current_marker != previous_marker:
                    break
        return all_results, pages_checked

    def _run(self, brand: str, campaigns: list[dict[str, str]]) -> None:
        playwright = None
        browser = None
        results: list[dict[str, Any]] = []
        successful_campaigns: list[dict[str, Any]] = []
        pages_checked = 0
        errors: list[str] = []
        terminal_error: Exception | None = None
        try:
            self._set(status="preparing", message="실행 전 최신 Google Drive 원본을 읽고 있습니다.")
            sync = self.sync_callback()
            if str(sync.get("status", "")) in {"failed", "login_required"}:
                raise RuntimeError(sync.get("message") or "Google Drive 최신화에 실패했습니다.")
            # Let Playwright perform the first campaign navigation after it
            # attaches; launching Chrome directly on the same URL can race
            # with page.goto() and produce an interrupted-navigation error.
            self._open_debug_chrome("about:blank")
            from playwright.sync_api import sync_playwright

            playwright = sync_playwright().start()
            browser = playwright.chromium.connect_over_cdp(
                f"http://127.0.0.1:{self.debug_port}"
            )
            context = browser.contexts[0]
            page = next(
                (
                    item
                    for item in context.pages
                    if "brandconnect.naver.com" in item.url
                    or "nid.naver.com" in item.url
                ),
                None,
            )
            if page is None:
                page = context.new_page()
            self._set(status="running", message="캠페인 제안 결과를 확인하고 있습니다.")
            for index, campaign in enumerate(campaigns, start=1):
                if self.stop_event.is_set():
                    break
                self._set(
                    current_campaign=campaign["url"],
                    message=(
                        f"{campaign['product_label']} 캠페인 "
                        f"{index}/{len(campaigns)} 확인 중"
                    ),
                )
                try:
                    campaign_results, campaign_pages = self._read_campaign(
                        page,
                        campaign,
                    )
                    results.extend(campaign_results)
                    successful_campaigns.append(dict(campaign))
                    pages_checked += campaign_pages
                    self._set(
                        processed_campaigns=index,
                        creators_seen=count_unique_creators(results),
                        pages_checked=pages_checked,
                    )
                except Exception as exc:
                    message = f"{campaign['product_label']}: {exc}"
                    errors.append(message)
                    self._set(
                        failed=len(errors),
                        errors=errors[-20:],
                        processed_campaigns=index,
                    )
            if not results and not self.stop_event.is_set():
                raise RuntimeError("캠페인에서 제안 크리에이터 명단을 읽지 못했습니다.")
            self._set(
                status="saving",
                message="크리에이터를 시트와 매칭해 상태와 콘텐츠 링크를 저장합니다.",
                current_campaign="",
            )
            saved = (
                self.persist_callback(
                    brand,
                    results,
                    successful_campaigns if not errors else [],
                )
                if results
                else {}
            )
            unmatched_names = list(saved.get("unmatched_names") or [])
            if unmatched_names:
                errors.append(
                    "시트 미매칭: " + ", ".join(str(name) for name in unmatched_names[:10])
                )
            self._set(
                matched=int(saved.get("matched", 0)),
                updated=int(saved.get("updated", 0)),
                added=int(saved.get("added", 0)),
                dated=int(saved.get("dated", 0)),
                reconciled=int(saved.get("reconciled", 0)),
                summary_rows=int(saved.get("summary_rows", 0)),
                links=int(saved.get("links", 0)),
                unmatched=int(saved.get("unmatched", 0)),
            )
            drive = dict(saved.get("drive") or {})
            if str(drive.get("status", "")) not in {"completed", "skipped"}:
                errors.append(
                    str(drive.get("message") or "Google Drive 업로드 확인이 필요합니다.")
                )
        except Exception as exc:
            terminal_error = exc
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
        finished_at = time.strftime("%Y-%m-%dT%H:%M:%S")
        snapshot = self.snapshot()
        if terminal_error is not None:
            self._set(
                status="failed",
                message=f"제안 결과 확인 실패: {terminal_error}",
                failed=max(1, int(snapshot.get("failed", 0))),
                errors=(errors + [str(terminal_error)])[-20:],
                current_campaign="",
                finished_at=finished_at,
            )
        elif self.stop_event.is_set():
            self._set(
                status="stopped",
                message="제안 결과 확인을 중단했습니다.",
                errors=errors[-20:],
                current_campaign="",
                finished_at=finished_at,
            )
        else:
            updated = int(self.snapshot().get("updated", 0))
            added = int(self.snapshot().get("added", 0))
            self._set(
                status="completed_with_warning" if errors else "completed",
                message=(
                    f"제안 결과 확인 완료 · 시트 {updated}건 반영"
                    + (f" · 시트 신규 {added}명 추가" if added else "")
                ),
                errors=errors[-20:],
                current_campaign="",
                finished_at=finished_at,
            )
