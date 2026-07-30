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


RUNNING_STATUSES = {
    "queued",
    "running",
    "stopping",
    "login_required",
    "account_required",
}


class InstagramDMSenderManager:
    """Send the remaining weekly DM goal in a background Chrome session."""

    def __init__(
        self,
        app_dir: Path,
        dashboard_callback: Callable[[str, bool], dict[str, Any]],
        record_callback: Callable[[str, int, str], dict[str, Any]],
        brands: dict[str, dict[str, Any]],
        reference_images: dict[str, Path],
    ):
        self.app_dir = Path(app_dir)
        self.dashboard_callback = dashboard_callback
        self.record_callback = record_callback
        self.brands = brands
        self.reference_images = reference_images
        self.profile_dir = self.app_dir / "app_data" / "instagram_dm_chrome"
        self.debug_port = 9232
        self.lock = threading.Lock()
        self.stop_event = threading.Event()
        self.resume_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.state: dict[str, Any] = self._idle_state()

    @staticmethod
    def _idle_state() -> dict[str, Any]:
        return {
            "status": "idle",
            "message": "자동 발송 대기 중",
            "brand": "",
            "expected_account": "",
            "goal": 0,
            "weekly_sent_at_start": 0,
            "requested": 0,
            "completed": 0,
            "failed": 0,
            "current_target": "",
            "last_error": "",
            "started_at": None,
            "finished_at": None,
        }

    def _set(self, **values: Any) -> None:
        with self.lock:
            self.state.update(values)

    def snapshot(self) -> dict[str, Any]:
        with self.lock:
            state = dict(self.state)
            state["running"] = bool(
                self.worker and self.worker.is_alive()
            )
            return state

    def is_running(self) -> bool:
        return self.snapshot().get("status") in RUNNING_STATUSES

    def start(self, brand: str) -> dict[str, Any]:
        if brand not in self.brands:
            raise ValueError("지원하지 않는 브랜드입니다.")
        with self.lock:
            if self.worker and self.worker.is_alive():
                raise ValueError("이미 목표 자동 발송이 진행 중입니다.")

        dashboard = self.dashboard_callback(brand, True)
        remaining = int(dashboard.get("remaining_to_goal", 0) or 0)
        if remaining <= 0:
            raise ValueError("이번 주 목표를 이미 달성했습니다.")
        if not dashboard.get("message_ready"):
            raise ValueError("DM 문구를 먼저 저장해주세요.")
        image_path = self.reference_images.get(brand)
        if image_path is None or not image_path.is_file():
            raise ValueError("공동구매 구성안 이미지를 찾을 수 없습니다.")
        if not dashboard.get("targets"):
            raise ValueError("발송할 대기 대상이 없습니다.")

        self.stop_event.clear()
        self.resume_event.clear()
        expected_account = str(
            self.brands[brand].get("expected_account", "")
        ).strip()
        with self.lock:
            self.state = {
                "status": "queued",
                "message": "목표 자동 발송을 준비하고 있습니다.",
                "brand": brand,
                "expected_account": expected_account,
                "goal": int(dashboard.get("goal", 0) or 0),
                "weekly_sent_at_start": int(
                    dashboard.get("weekly_sent", 0) or 0
                ),
                "requested": remaining,
                "completed": 0,
                "failed": 0,
                "current_target": "",
                "last_error": "",
                "started_at": datetime.now().isoformat(timespec="seconds"),
                "finished_at": None,
            }
            self.worker = threading.Thread(
                target=self._run,
                args=(brand,),
                daemon=True,
                name="instagram-dm-goal-sender",
            )
            self.worker.start()
        return self.snapshot()

    def stop(self) -> dict[str, Any]:
        if not self.is_running():
            return self.snapshot()
        self.stop_event.set()
        self.resume_event.set()
        self._set(
            status="stopping",
            message="현재 단계가 끝나는 즉시 중단합니다.",
        )
        return self.snapshot()

    def resume_after_login(self) -> dict[str, Any]:
        status = self.snapshot().get("status")
        if status not in {"login_required", "account_required"}:
            raise ValueError("현재 로그인 또는 계정 확인 대기 상태가 아닙니다.")
        self.resume_event.set()
        self._set(
            status="running",
            message="Instagram 로그인과 발송 계정을 다시 확인합니다.",
        )
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
                "https://www.instagram.com/",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        deadline = time.time() + 20
        while time.time() < deadline:
            if self._debug_ready():
                return
            if self.stop_event.wait(0.25):
                return
        raise RuntimeError("Instagram 자동화용 Chrome에 연결하지 못했습니다.")

    @staticmethod
    def _account_visible(page: Any, expected_account: str) -> bool:
        username = expected_account.lstrip("@")
        try:
            if page.locator(f'a[href="/{username}/"]').count() > 0:
                return True
            return username.casefold() in page.locator("body").inner_text().casefold()
        except Exception:
            return False

    def _wait_for_account(self, page: Any, expected_account: str) -> bool:
        while not self.stop_event.is_set():
            if self._account_visible(page, expected_account):
                return True
            body_text = ""
            try:
                body_text = page.locator("body").inner_text().casefold()
            except Exception:
                pass
            if "로그인" in body_text or "log in" in body_text:
                status = "login_required"
                instruction = "열린 Chrome에서 Instagram 로그인을 완료해주세요."
            else:
                status = "account_required"
                instruction = (
                    f"열린 Chrome에서 더 보기 → 계정 전환을 눌러 "
                    f"{expected_account} 계정으로 바꿔주세요."
                )
            self.resume_event.clear()
            self._set(status=status, message=instruction)
            self.resume_event.wait()
            if self.stop_event.is_set():
                return False
            page.goto(
                "https://www.instagram.com/",
                wait_until="domcontentloaded",
                timeout=60_000,
            )
            page.wait_for_timeout(1_500)
        return False

    @staticmethod
    def _click_unique(locator: Any, description: str) -> None:
        count = locator.count()
        if count != 1:
            raise RuntimeError(f"{description}을(를) 확인할 수 없습니다.")
        locator.click(timeout=15_000)

    def _send_target(
        self,
        page: Any,
        target: dict[str, Any],
        message: str,
        image_path: Path,
    ) -> None:
        instagram_id = str(target["instagram_id"]).strip().lstrip("@")
        page.goto(
            str(target.get("profile_url") or "")
            or f"https://www.instagram.com/{instagram_id}/",
            wait_until="domcontentloaded",
            timeout=60_000,
        )
        page.wait_for_timeout(1_300)

        profile_text = page.locator("body").inner_text()
        if instagram_id.casefold() not in profile_text.casefold():
            raise RuntimeError("대상 프로필을 확인할 수 없습니다.")

        composer = page.get_by_role("textbox")
        if composer.count() != 1:
            self._click_unique(
                page.get_by_role(
                    "button", name="메시지 보내기", exact=True
                ),
                "메시지 보내기 버튼",
            )
            for _ in range(12):
                if self.stop_event.wait(0.5):
                    return
                composer = page.get_by_role("textbox")
                if composer.count() == 1:
                    break
        if composer.count() != 1:
            raise RuntimeError("DM 입력칸을 확인할 수 없습니다.")

        composer.fill(message)
        self._click_unique(
            page.get_by_role("button", name="보내기", exact=True),
            "문구 보내기 버튼",
        )
        page.wait_for_timeout(800)
        first_line = message.splitlines()[0] if message.splitlines() else message
        if first_line not in page.locator("body").inner_text():
            raise RuntimeError("DM 문구 발송을 확인하지 못했습니다.")

        image_button = page.get_by_role(
            "button", name="사진 또는 동영상 추가", exact=True
        )
        if image_button.count() != 1:
            raise RuntimeError("사진 추가 버튼을 확인할 수 없습니다.")
        with page.expect_file_chooser(timeout=10_000) as chooser_info:
            image_button.click(timeout=15_000)
        chooser_info.value.set_files(str(image_path))

        image_send = None
        for _ in range(12):
            if self.stop_event.wait(0.5):
                return
            image_send = page.get_by_role(
                "button", name="보내기", exact=True
            )
            if image_send.count() == 1:
                break
        if image_send is None or image_send.count() != 1:
            raise RuntimeError("이미지 보내기 버튼을 확인할 수 없습니다.")
        image_send.click(timeout=15_000)
        page.wait_for_timeout(1_000)
        if page.locator(
            'button[aria-label^="첨부 파일 삭제"]'
        ).count() > 0:
            raise RuntimeError("구성안 이미지 발송을 확인하지 못했습니다.")

    def _run(self, brand: str) -> None:
        playwright = None
        browser = None
        attempted_rows: set[int] = set()
        try:
            self._set(status="running", message="Instagram 발송 계정을 확인합니다.")
            self._open_debug_chrome()
            if self.stop_event.is_set():
                raise InterruptedError

            from playwright.sync_api import sync_playwright

            playwright = sync_playwright().start()
            browser = playwright.chromium.connect_over_cdp(
                f"http://127.0.0.1:{self.debug_port}"
            )
            context = browser.contexts[0]
            pages = context.pages
            page = pages[0] if pages else context.new_page()
            page.goto(
                "https://www.instagram.com/",
                wait_until="domcontentloaded",
                timeout=60_000,
            )
            page.wait_for_timeout(1_500)

            expected_account = str(
                self.brands[brand].get("expected_account", "")
            ).strip()
            if not self._wait_for_account(page, expected_account):
                raise InterruptedError

            image_path = self.reference_images[brand]
            while not self.stop_event.is_set():
                dashboard = self.dashboard_callback(brand, False)
                remaining = int(
                    dashboard.get("remaining_to_goal", 0) or 0
                )
                if remaining <= 0:
                    self._set(
                        status="completed",
                        message="이번 주 목표까지 자동 발송을 완료했습니다.",
                        current_target="",
                        finished_at=datetime.now().isoformat(
                            timespec="seconds"
                        ),
                    )
                    return
                candidates = [
                    target
                    for target in dashboard.get("targets", [])
                    if int(target.get("row", 0) or 0) not in attempted_rows
                ]
                if not candidates:
                    self._set(
                        status="completed_with_errors",
                        message=(
                            "남은 대상을 모두 확인했지만 일부 실패로 "
                            "목표를 전부 채우지 못했습니다."
                        ),
                        current_target="",
                        finished_at=datetime.now().isoformat(
                            timespec="seconds"
                        ),
                    )
                    return

                target = candidates[0]
                row = int(target["row"])
                instagram_id = str(target["instagram_id"])
                attempted_rows.add(row)
                self._set(
                    status="running",
                    current_target=f"@{instagram_id}",
                    message=f"@{instagram_id} 발송 중",
                    last_error="",
                )
                try:
                    if not self._account_visible(page, expected_account):
                        if not self._wait_for_account(
                            page, expected_account
                        ):
                            raise InterruptedError
                    message = str(target.get("message", "")).strip()
                    if not message:
                        raise RuntimeError("DM 문구가 비어 있습니다.")
                    self._send_target(page, target, message, image_path)
                    if self.stop_event.is_set():
                        raise InterruptedError
                    record = self.record_callback(brand, row, "sent")
                    drive_status = str(
                        (record.get("drive_sync") or {}).get("status", "")
                    )
                    if drive_status not in {"completed", "disabled"}:
                        raise RuntimeError(
                            (record.get("drive_sync") or {}).get(
                                "message", "Google Drive 최신화 실패"
                            )
                        )
                    completed = int(
                        self.snapshot().get("completed", 0) or 0
                    ) + 1
                    self._set(
                        completed=completed,
                        current_target="",
                        message=(
                            f"@{instagram_id} 완료 · "
                            f"{completed}/{self.snapshot()['requested']}명"
                        ),
                    )
                except InterruptedError:
                    raise
                except Exception as exc:
                    failed = int(self.snapshot().get("failed", 0) or 0) + 1
                    self._set(
                        failed=failed,
                        last_error=f"@{instagram_id}: {exc}",
                        message=(
                            f"@{instagram_id} 실패 · 다음 대상을 계속합니다."
                        ),
                    )
                    if self.stop_event.wait(1):
                        raise InterruptedError

            raise InterruptedError
        except InterruptedError:
            self._set(
                status="stopped",
                message="목표 자동 발송을 중단했습니다.",
                current_target="",
                finished_at=datetime.now().isoformat(timespec="seconds"),
            )
        except Exception as exc:
            self._set(
                status="error",
                message=f"목표 자동 발송 실패: {exc}",
                last_error=str(exc),
                current_target="",
                finished_at=datetime.now().isoformat(timespec="seconds"),
            )
        finally:
            if browser is not None:
                try:
                    browser.close()
                except Exception:
                    pass
            if playwright is not None:
                try:
                    playwright.stop()
                except Exception:
                    pass
