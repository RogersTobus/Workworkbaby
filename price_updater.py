from __future__ import annotations

import json
import platform
import re
import threading
import time
from pathlib import Path
from typing import Any
from urllib.parse import quote


PRICE_PATTERN = re.compile(r"([\d,]+)\s*원")
GOOGLE_SCOPE = ["https://www.googleapis.com/auth/spreadsheets"]
MAX_SECURITY_RETRIES = 2


class NaverSecurityCheckError(RuntimeError):
    pass


class PriceUpdateManager:
    def __init__(self, app_dir: Path, config: dict[str, Any]) -> None:
        self.app_dir = app_dir
        self.config = config
        self.lock = threading.Lock()
        self.resume_event = threading.Event()
        self.stop_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.state: dict[str, Any] = {}
        self._reset_state()
        self._load_last_run()

    def _reset_state(self) -> None:
        self.state = {
            "status": "idle",
            "message": "가격 최신화를 시작할 수 있습니다.",
            "processed": 0,
            "total": 0,
            "updated": 0,
            "errors": 0,
            "current_row": None,
            "current_product": "",
            "results": [],
            "started_at": None,
            "finished_at": None,
            "last_updated_at": None,
        }

    def _history_path(self) -> Path:
        return self.app_dir / "app_data" / "price_last_run.json"

    def _checkpoint_path(self) -> Path:
        return self.app_dir / "app_data" / "price_resume.json"

    def _load_checkpoint(self) -> list[dict[str, Any]]:
        path = self._checkpoint_path()
        if not path.exists():
            return []
        try:
            saved = json.loads(path.read_text(encoding="utf-8"))
            return saved.get("results", [])
        except (OSError, ValueError, TypeError):
            return []

    def _save_checkpoint(self) -> None:
        path = self._checkpoint_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(".tmp")
        temporary.write_text(
            json.dumps(self.snapshot(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        temporary.replace(path)

    def _clear_checkpoint(self) -> None:
        try:
            self._checkpoint_path().unlink(missing_ok=True)
        except OSError:
            pass

    def _load_last_run(self) -> None:
        path = self._history_path()
        if not path.exists():
            return
        try:
            saved = json.loads(path.read_text(encoding="utf-8"))
            results = saved.get("results", [])
            self.state.update(
                {
                    "message": "마지막 가격 확인 결과를 불러왔습니다. 다시 최신화할 수 있습니다.",
                    "processed": len(results),
                    "total": saved.get("total", len(results)),
                    "updated": saved.get("updated", 0),
                    "errors": saved.get("errors", 0),
                    "results": results,
                    "finished_at": saved.get("finished_at"),
                    "last_updated_at": saved.get(
                        "last_updated_at", saved.get("finished_at")
                    ),
                }
            )
        except (OSError, ValueError, TypeError):
            pass

    def _save_last_run(self) -> None:
        path = self._history_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(".tmp")
        temporary.write_text(
            json.dumps(self.snapshot(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        temporary.replace(path)

    def snapshot(self) -> dict[str, Any]:
        with self.lock:
            return json.loads(json.dumps(self.state, ensure_ascii=False))

    def _set(self, **changes: Any) -> None:
        with self.lock:
            self.state.update(changes)

    def start(self) -> dict[str, Any]:
        with self.lock:
            if self.worker and self.worker.is_alive():
                return json.loads(json.dumps(self.state, ensure_ascii=False))
            previous_updated_at = self.state.get("last_updated_at")
            self._reset_state()
            self.state.update(
                {
                    "status": "starting",
                    "message": "Google Sheet 연결을 확인하고 있습니다.",
                    "started_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "last_updated_at": previous_updated_at,
                }
            )
            self.resume_event.clear()
            self.stop_event.clear()
            self.worker = threading.Thread(target=self._run, daemon=True)
            self.worker.start()
            return json.loads(json.dumps(self.state, ensure_ascii=False))

    def resume_after_login(self) -> dict[str, Any]:
        status = self.snapshot().get("status")
        self.resume_event.set()
        if status == "naver_security_required":
            self._set(message="가격 확인용 브라우저를 새로 열고 같은 상품부터 재시도합니다.")
        else:
            self._set(message="네이버 로그인 상태를 다시 확인합니다.")
        return self.snapshot()

    def stop(self) -> dict[str, Any]:
        self.stop_event.set()
        self.resume_event.set()
        self._set(message="중지 요청을 처리하고 있습니다.")
        return self.snapshot()

    def _google_paths(self) -> tuple[Path, Path]:
        credential_name = self.config.get(
            "google_credentials_path", "google_credentials.json"
        )
        token_name = self.config.get(
            "google_token_path", "app_data/google_token.json"
        )
        return self.app_dir / credential_name, self.app_dir / token_name

    def _google_service(self):
        try:
            from google.auth.transport.requests import Request
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise RuntimeError(
                "Google 연결 구성요소가 없습니다. install_windows.bat을 다시 실행해주세요."
            ) from exc

        credentials_path, token_path = self._google_paths()
        if not credentials_path.exists():
            raise RuntimeError(
                "google_credentials.json이 없습니다. Google Cloud의 데스크톱 OAuth "
                "클라이언트 파일을 앱 폴더에 넣어주세요."
            )

        token_path.parent.mkdir(parents=True, exist_ok=True)
        credentials = None
        if token_path.exists():
            credentials = Credentials.from_authorized_user_file(
                str(token_path), GOOGLE_SCOPE
            )
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        if not credentials or not credentials.valid:
            self._set(
                status="google_login_required",
                message="열린 Google 로그인 화면에서 시트 접근을 허용해주세요.",
            )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(credentials_path), GOOGLE_SCOPE
            )
            credentials = flow.run_local_server(port=0, open_browser=True)
        token_path.write_text(credentials.to_json(), encoding="utf-8")
        return build("sheets", "v4", credentials=credentials, cache_discovery=False)

    def _read_targets(self, service) -> list[dict[str, Any]]:
        spreadsheet_id = self.config["spreadsheet_id"]
        sheet_name = self.config["sheet_name"].replace("'", "''")
        start_row = int(self.config["start_row"])
        end_row = int(self.config["end_row"])
        ranges = [
            f"'{sheet_name}'!D{start_row}:D{end_row}",
            f"'{sheet_name}'!X{start_row}:Y{end_row}",
            f"'{sheet_name}'!AE{start_row}:AE{end_row}",
        ]
        response = (
            service.spreadsheets()
            .values()
            .batchGet(
                spreadsheetId=spreadsheet_id,
                ranges=ranges,
                valueRenderOption="UNFORMATTED_VALUE",
            )
            .execute()
        )
        value_ranges = response.get("valueRanges", [])
        names = value_ranges[0].get("values", []) if len(value_ranges) > 0 else []
        prices = value_ranges[1].get("values", []) if len(value_ranges) > 1 else []
        links = value_ranges[2].get("values", []) if len(value_ranges) > 2 else []

        targets = []
        for offset in range(end_row - start_row + 1):
            product = names[offset][0] if offset < len(names) and names[offset] else ""
            old_prices = prices[offset] if offset < len(prices) else []
            url = links[offset][0] if offset < len(links) and links[offset] else ""
            if not str(url).strip():
                continue
            targets.append(
                {
                    "row": start_row + offset,
                    "product": str(product),
                    "url": str(url).strip(),
                    "old_instant": old_prices[0] if len(old_prices) > 0 else None,
                    "old_best": old_prices[1] if len(old_prices) > 1 else None,
                }
            )
        return targets

    def _launch_browser(self, playwright):
        profile = self.app_dir / "app_data" / "naver_browser_profile"
        profile.mkdir(parents=True, exist_ok=True)
        channels: list[str | None]
        system = platform.system()
        if system == "Windows":
            # 가격 수집은 로그인 상태를 유지하는 전용 Chrome 프로필에서만 진행한다.
            channels = ["chrome"]
        elif system == "Darwin":
            channels = ["chrome", "msedge", None]
        else:
            channels = ["chrome", None]

        last_error = None
        for channel in channels:
            try:
                kwargs: dict[str, Any] = {
                    "user_data_dir": str(profile),
                    "headless": False,
                    "viewport": {"width": 1280, "height": 900},
                }
                if channel:
                    kwargs["channel"] = channel
                return playwright.chromium.launch_persistent_context(**kwargs)
            except Exception as exc:
                last_error = exc
        raise RuntimeError(
            "가격 확인용 브라우저를 실행하지 못했습니다. "
            "install_windows.bat을 다시 실행해주세요."
        ) from last_error

    @staticmethod
    def _is_logged_out(page) -> bool:
        try:
            return (
                "nidlogin.login" in page.url
                or page.get_by_role("link", name="로그인", exact=True).count() > 0
            )
        except Exception:
            return True

    @staticmethod
    def _has_naver_login(browser_context) -> bool:
        """화면 문구가 아니라 네이버 로그인 세션 자체가 있는지 확인한다."""
        try:
            cookies = browser_context.cookies(
                ["https://www.naver.com", "https://shopping.naver.com"]
            )
        except Exception:
            return False
        names = {
            cookie.get("name")
            for cookie in cookies
            if str(cookie.get("domain", "")).endswith("naver.com")
        }
        return "NID_AUT" in names and "NID_SES" in names

    @staticmethod
    def _parse_price(text: str) -> int:
        match = PRICE_PATTERN.search(text)
        if not match:
            raise ValueError(f"가격을 찾을 수 없습니다: {text[:120]}")
        return int(match.group(1).replace(",", ""))

    @staticmethod
    def _is_security_check(page) -> bool:
        markers = (
            "보안 확인",
            "보안확인",
            "자동입력 방지",
            "자동 입력 방지",
            "비정상적인 접근",
            "접근이 제한",
            "접속이 제한",
            "captcha",
        )
        try:
            title = page.title()
        except Exception:
            title = ""
        try:
            url = page.url
        except Exception:
            url = ""
        try:
            body = page.locator("body").inner_text(timeout=3_000)
        except Exception:
            body = ""
        text = f"{title}\n{url}\n{body}".lower()
        return any(marker.lower() in text for marker in markers)

    @staticmethod
    def _browser_was_closed(error: Exception) -> bool:
        text = str(error).lower()
        return any(
            marker in text
            for marker in (
                "target page, context or browser has been closed",
                "browser has been closed",
                "page has been closed",
                "target closed",
            )
        )

    def _extract_prices(self, page, url: str) -> tuple[int, int]:
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
        if self._is_security_check(page):
            raise NaverSecurityCheckError("네이버 보안확인이 감지되었습니다.")
        try:
            # 같은 접근성 문구가 두 번 있어도 첫 가격 영역이 나타나면 계속한다.
            page.get_by_text("상품 가격", exact=True).first.wait_for(timeout=30_000)
        except Exception:
            if self._is_security_check(page):
                raise NaverSecurityCheckError("네이버 보안확인이 감지되었습니다.")
            raise
        price_labels = page.get_by_text("상품 가격", exact=True)
        price_label_count = price_labels.count()
        if price_label_count < 1:
            raise ValueError("상품 가격을 찾을 수 없습니다.")
        price_label = price_labels.first
        instant_text = price_label.locator("xpath=..").inner_text()
        instant = self._parse_price(instant_text)

        best = instant
        # 네이버의 새 화면은 로그인 할인가도 '상품 가격'으로 한 번 더 표시한다.
        if price_label_count > 1:
            best_text = price_labels.nth(1).locator("xpath=..").inner_text()
            best = self._parse_price(best_text)
        else:
            # 이전 화면 형식도 계속 지원한다.
            for label_text in ("최대할인가", "최대혜택가"):
                best_labels = page.get_by_text(label_text, exact=False)
                if best_labels.count() < 1:
                    continue
                price_container = best_labels.first.locator(
                    "xpath=ancestor::*[contains(normalize-space(.), '상품 가격')][1]"
                )
                best = self._parse_price(price_container.inner_text())
                break
        return instant, best

    def _wait_for_naver_login(
        self, browser_context, page, return_url: str
    ) -> bool:
        if self._has_naver_login(browser_context) and not self._is_logged_out(page):
            return True
        login_url = (
            "https://nid.naver.com/nidlogin.login?mode=form&url="
            + quote(return_url, safe="")
        )
        page.goto(login_url, wait_until="domcontentloaded", timeout=60_000)
        self._set(
            status="naver_login_required",
            message=(
                "가격 확인용 브라우저에서 네이버 로그인을 완료한 뒤 "
                "컨트롤 타워의 '로그인 완료·계속'을 눌러주세요."
            ),
        )
        self.resume_event.clear()
        self.resume_event.wait()
        if self.stop_event.is_set():
            return False

        # 로그인 완료 직후 네이버가 원래 상품으로 자동 이동한다. 그 이동과
        # page.goto가 겹치면 ERR_ABORTED가 나므로 세션과 이동이 안정될 때까지 기다린다.
        deadline = time.time() + 30
        while time.time() < deadline:
            if self._has_naver_login(browser_context):
                try:
                    page.wait_for_load_state("domcontentloaded", timeout=3_000)
                except Exception:
                    pass
                if "nidlogin.login" not in page.url:
                    break
            if self.stop_event.wait(0.5):
                return False

        if not self._has_naver_login(browser_context):
            return False
        if page.url.rstrip("/") != return_url.rstrip("/"):
            page.goto(return_url, wait_until="domcontentloaded", timeout=60_000)
        return (
            self._has_naver_login(browser_context)
            and not self._is_logged_out(page)
        )

    def _write_results(self, service, results: list[dict[str, Any]]) -> None:
        spreadsheet_id = self.config["spreadsheet_id"]
        sheet_name = self.config["sheet_name"].replace("'", "''")
        data = [
            {
                "range": f"'{sheet_name}'!X{item['row']}:Y{item['row']}",
                "majorDimension": "ROWS",
                "values": [[item["instant"], item["best"]]],
            }
            for item in results
            if item["status"] == "ready"
        ]
        if not data:
            return
        (
            service.spreadsheets()
            .values()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"valueInputOption": "RAW", "data": data},
            )
            .execute()
        )

    def _run(self) -> None:
        browser_context = None
        try:
            service = self._google_service()
            targets = self._read_targets(service)
            self._set(
                status="opening_browser",
                message="가격 확인용 브라우저를 열고 있습니다.",
                total=len(targets),
            )
            try:
                from playwright.sync_api import sync_playwright
            except ImportError as exc:
                raise RuntimeError(
                    "브라우저 자동화 구성요소가 없습니다. "
                    "install_windows.bat을 다시 실행해주세요."
                ) from exc

            checkpoint_results = self._load_checkpoint()
            target_rows = {target["row"] for target in targets}
            checkpoint_results = [
                item
                for item in checkpoint_results
                if item.get("row") in target_rows
            ]
            completed_rows = {item["row"] for item in checkpoint_results}
            ready_results: list[dict[str, Any]] = [
                item
                for item in checkpoint_results
                if item.get("status") == "ready"
            ]
            if checkpoint_results:
                self._set(
                    results=checkpoint_results,
                    processed=len(checkpoint_results),
                    updated=sum(
                        1
                        for item in checkpoint_results
                        if item.get("status") == "ready" and item.get("changed")
                    ),
                    errors=sum(
                        1
                        for item in checkpoint_results
                        if item.get("status") == "error"
                    ),
                    message=(
                        f"저장된 {len(checkpoint_results)}개 결과 다음부터 이어갑니다."
                    ),
                )
            with sync_playwright() as playwright:
                browser_context = self._launch_browser(playwright)
                page = browser_context.pages[0] if browser_context.pages else browser_context.new_page()
                pending_targets = [
                    target for target in targets if target["row"] not in completed_rows
                ]
                if pending_targets:
                    page.goto(
                        pending_targets[0]["url"],
                        wait_until="domcontentloaded",
                        timeout=60_000,
                    )
                    if not self._wait_for_naver_login(
                        browser_context, page, pending_targets[0]["url"]
                    ):
                        if not self.stop_event.is_set():
                            raise RuntimeError(
                                "네이버 로그인이 확인되지 않았습니다. 다시 시도해주세요."
                            )
                        return

                self._set(status="crawling", message="네이버 가격을 확인하고 있습니다.")
                for position, target in enumerate(targets, start=1):
                    if target["row"] in completed_rows:
                        continue
                    if self.stop_event.is_set():
                        self._set(status="stopped", message="사용자가 작업을 중지했습니다.")
                        return
                    self._set(
                        current_row=target["row"],
                        current_product=target["product"],
                        message=f"{position}/{len(targets)} 상품 가격 확인 중",
                    )
                    result = dict(target)
                    security_retries = 0
                    while True:
                        try:
                            if not self._wait_for_naver_login(
                                browser_context, page, target["url"]
                            ):
                                raise RuntimeError(
                                    "네이버 로그인이 확인되지 않았습니다. 다시 로그인해주세요."
                                )
                            instant, best = self._extract_prices(page, target["url"])
                            result.update(
                                {
                                    "instant": instant,
                                    "best": best,
                                    "status": "ready",
                                    "changed": (
                                        target["old_instant"] != instant
                                        or target["old_best"] != best
                                    ),
                                }
                            )
                            ready_results.append(result)
                            break
                        except Exception as exc:
                            security_check = isinstance(
                                exc, NaverSecurityCheckError
                            ) or self._browser_was_closed(exc)
                            if not security_check:
                                result.update(
                                    {
                                        "instant": None,
                                        "best": None,
                                        "status": "error",
                                        "error": str(exc),
                                        "changed": False,
                                    }
                                )
                                break

                            security_retries += 1
                            if browser_context is not None:
                                try:
                                    browser_context.close()
                                except Exception:
                                    pass
                                browser_context = None

                            if security_retries >= MAX_SECURITY_RETRIES:
                                result.update(
                                    {
                                        "instant": None,
                                        "best": None,
                                        "status": "error",
                                        "error": (
                                            "같은 링크에서 네이버 보안확인이 "
                                            f"{MAX_SECURITY_RETRIES}회 감지되어 다음 상품으로 이동했습니다."
                                        ),
                                        "changed": False,
                                    }
                                )
                                break

                            retry_delay = min(3 * security_retries, 15)
                            self._set(
                                status="naver_security_restarting",
                                message=(
                                    "네이버 보안확인이 감지되어 가격 확인 창을 닫았습니다. "
                                    f"{retry_delay}초 후 새 창으로 현재 상품을 자동 재시도합니다. "
                                    f"({security_retries}/{MAX_SECURITY_RETRIES})"
                                ),
                            )
                            if self.stop_event.wait(retry_delay):
                                self._set(
                                    status="stopped",
                                    message="사용자가 작업을 중지했습니다.",
                                )
                                return

                            self._set(
                                status="opening_browser",
                                message="가격 확인용 브라우저를 새로 열고 있습니다.",
                            )
                            browser_context = self._launch_browser(playwright)
                            page = (
                                browser_context.pages[0]
                                if browser_context.pages
                                else browser_context.new_page()
                            )
                            self._set(
                                status="crawling",
                                message=(
                                    f"{position}/{len(targets)} 상품을 자동으로 다시 확인합니다."
                                ),
                            )
                    result["checked_at"] = time.strftime("%Y.%m.%d %H:%M:%S")
                    with self.lock:
                        self.state["results"].append(result)
                        self.state["processed"] = position
                        self.state["updated"] = sum(
                            1
                            for item in self.state["results"]
                            if item.get("status") == "ready"
                            and item.get("changed")
                        )
                        self.state["errors"] = sum(
                            1
                            for item in self.state["results"]
                            if item.get("status") == "error"
                        )
                    self._save_checkpoint()
                    if self.stop_event.wait(0.7):
                        self._set(status="stopped", message="사용자가 작업을 중지했습니다.")
                        return

                self._set(status="writing", message="확인된 가격을 Google Sheet에 반영합니다.")
                self._write_results(service, ready_results)
                completed_at = time.strftime("%Y.%m.%d %H:%M:%S")
                self._set(
                    status="completed",
                    message="가격 최신화가 완료되었습니다.",
                    finished_at=completed_at,
                    last_updated_at=completed_at,
                    current_row=None,
                    current_product="",
                )
                self._save_last_run()
                self._clear_checkpoint()
        except Exception as exc:
            self._set(
                status="error",
                message=str(exc),
                finished_at=time.strftime("%Y-%m-%d %H:%M:%S"),
            )
        finally:
            if browser_context is not None:
                try:
                    browser_context.close()
                except Exception:
                    pass
