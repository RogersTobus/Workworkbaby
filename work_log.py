from __future__ import annotations

import re
from collections import defaultdict
from datetime import date, timedelta
from typing import Any


STATUS_LABELS = {
    "done": "완료",
    "doing": "진행 중",
    "review": "검토 중",
    "todo": "예정",
}

PARTNER_RULES = (
    ("현대H몰", ("현대h몰", "현대홈쇼핑", "현대톡마트")),
    ("현대그린푸드", ("현대그린푸드", "현대 그린푸드")),
    ("현대백화점", ("현대백화점", "현대 백화점", "현대아울렛", "현대 아울렛")),
    ("베네피아", ("베네피아",)),
    ("컬리", ("컬리",)),
    ("G마켓/옥션", ("g마켓", "지마켓", "옥션")),
    ("SSG", ("ssg", "쓱특가", "오반장")),
    ("롯데아이몰", ("롯데아이몰",)),
    ("롯데홈쇼핑", ("롯데홈쇼핑", "롯데 홈쇼핑")),
    ("CJ온스타일", ("cj온스타일", "cj 온스타일", "cj")),
    ("KT알파쇼핑", ("kt알파", "kt 알파", "kt 상품", "kt상품")),
    ("11번가", ("11번가",)),
    ("카카오 톡스토어", ("톡딜", "톡스토어", "카카오")),
    ("메이커스", ("메이커스",)),
    ("브랜드커넥트", ("브랜드커넥트", "브랜드 커넥트")),
    ("공동구매", ("공구", "공동구매", "벤더사")),
    ("SNS 콘텐츠", ("sns", "인스타그램", "페이스북")),
)

BRAND_RULES = {
    "gaia": ("가이아", "올리브", "병절임", "발사믹", "파우치"),
    "alp": ("알프", "이뮨", "아이언드롭", "아이언 드롭"),
    "coffee": ("이탈리안커피", "이탈리안 커피", "네스프레소", "돌체", "커피"),
}

BRAND_LABELS = {
    "gaia": "가이아",
    "alp": "알프",
    "coffee": "이탈리안커피",
    "general": "분류 필요",
}


def latest_or_upcoming_friday(target: date) -> date:
    if target.weekday() <= 4:
        return target + timedelta(days=4 - target.weekday())
    return target - timedelta(days=target.weekday() - 4)


def report_period(friday: date) -> tuple[date, date, date, date]:
    start = friday - timedelta(days=4)
    next_start = friday + timedelta(days=3)
    next_end = next_start + timedelta(days=4)
    return start, friday, next_start, next_end


def _date(value: object) -> date | None:
    try:
        return date.fromisoformat(str(value or ""))
    except ValueError:
        return None


def _routine_occurs(item: dict[str, Any], target: date) -> bool:
    start = _date(item.get("date"))
    if start is None or target < start:
        return False
    recurrence = str(item.get("recurrence", "weekly"))
    if recurrence == "daily":
        return True
    if recurrence == "weekdays":
        return target.weekday() < 5
    if recurrence == "weekly":
        return target.weekday() == start.weekday()
    if recurrence == "monthly":
        return target.day == start.day
    return False


def _status_for(item: dict[str, Any], target: date | None = None) -> str:
    if target is not None:
        status = dict(item.get("statuses") or {}).get(target.isoformat())
        if status in STATUS_LABELS:
            return status
        if dict(item.get("completions") or {}).get(target.isoformat()):
            return "done"
    status = str(item.get("status", "todo"))
    return status if status in STATUS_LABELS else "todo"


def collect_calendar_items(
    tasks: list[dict[str, Any]],
    start: date,
    end: date,
) -> list[dict[str, str]]:
    collected: list[dict[str, str]] = []
    for item in tasks:
        kind = str(item.get("kind", "task"))
        text = str(item.get("text", "")).strip()
        if not text or kind == "checklist_routine":
            continue
        if kind == "event":
            event_start = _date(item.get("date"))
            event_end = _date(item.get("end_date")) or event_start
            if event_start and event_end and event_start <= end and event_end >= start:
                collected.append(
                    {
                        "id": str(item.get("id", "")),
                        "date": max(event_start, start).isoformat(),
                        "text": text.replace("_", " · "),
                        "status": "doing" if event_end >= end else "done",
                        "kind": "event",
                    }
                )
            continue
        if kind == "routine":
            occurrence_dates = []
            cursor = start
            while cursor <= end:
                if _routine_occurs(item, cursor):
                    occurrence_dates.append(cursor)
                cursor += timedelta(days=1)
            if occurrence_dates:
                statuses = [_status_for(item, day) for day in occurrence_dates]
                status = "done" if statuses and all(value == "done" for value in statuses) else next(
                    (value for value in ("doing", "review", "todo") if value in statuses),
                    "todo",
                )
                collected.append(
                    {
                        "id": str(item.get("id", "")),
                        "date": occurrence_dates[0].isoformat(),
                        "text": text,
                        "status": status,
                        "kind": "routine",
                    }
                )
            continue
        item_date = _date(item.get("date"))
        if item_date and start <= item_date <= end:
            collected.append(
                {
                    "id": str(item.get("id", "")),
                    "date": item_date.isoformat(),
                    "text": text,
                    "status": _status_for(item),
                    "kind": kind if kind in {"task", "meeting"} else "task",
                }
            )
    collected.sort(key=lambda item: (item["date"], item["text"]))
    return collected


def classify_partner(text: str) -> str:
    normalized = text.casefold()
    for label, keywords in PARTNER_RULES:
        if any(keyword.casefold() in normalized for keyword in keywords):
            return label
    return "기타"


def classify_brands(text: str) -> list[str]:
    normalized = text.casefold()
    matched = [
        brand
        for brand, keywords in BRAND_RULES.items()
        if any(keyword.casefold() in normalized for keyword in keywords)
    ]
    return matched or ["general"]


def _sheet_style_text(item: dict[str, str]) -> str:
    text = re.sub(r"\s+", " ", item["text"]).strip()
    text = re.sub(r"^[-•]\s*", "", text)
    text = re.sub(r"^\[(?:완료|진행 중|검토 중|예정)\]\s*", "", text)
    if item.get("kind") == "event":
        text = re.sub(r"\s*·\s*", " ", text)
        if not text.endswith(("진행", "확정", "준비")):
            text = f"{text} 진행"
    return text


def _lines(items: list[dict[str, str]], next_week: bool = False) -> str:
    seen: set[str] = set()
    output: list[str] = []
    for item in items:
        text = _sheet_style_text(item)
        key = text.casefold()
        if not text or key in seen:
            continue
        seen.add(key)
        output.append(text)
    if len(output) <= 1:
        return "" if not output else output[0]
    return "\n".join(f"{index}) {text}" for index, text in enumerate(output, 1))


def build_work_log_draft(
    tasks: list[dict[str, Any]],
    friday: date,
) -> dict[str, Any]:
    start, end, next_start, next_end = report_period(friday)
    current = collect_calendar_items(tasks, start, end)
    upcoming = collect_calendar_items(tasks, next_start, next_end)
    carry_over = [item for item in current if item["status"] != "done" and item["kind"] != "event"]
    next_items = upcoming + carry_over

    partner_current: dict[str, list[dict[str, str]]] = defaultdict(list)
    partner_next: dict[str, list[dict[str, str]]] = defaultdict(list)
    for item in current:
        partner_current[classify_partner(item["text"])].append(item)
    for item in next_items:
        partner_next[classify_partner(item["text"])].append(item)
    partner_names = sorted(
        set(partner_current) | set(partner_next),
        key=lambda value: (value == "기타", value),
    )
    partners = [
        {
            "id": f"partner-{index}",
            "partner": partner,
            "this_week": _lines(partner_current[partner]),
            "next_week": _lines(partner_next[partner], next_week=True),
        }
        for index, partner in enumerate(partner_names, 1)
    ]

    sales_current: dict[str, dict[str, list[dict[str, str]]]] = defaultdict(
        lambda: defaultdict(list)
    )
    sales_next: dict[str, dict[str, list[dict[str, str]]]] = defaultdict(
        lambda: defaultdict(list)
    )
    for item in current:
        platform = classify_partner(item["text"])
        for brand in classify_brands(item["text"]):
            sales_current[brand][platform].append(item)
    for item in next_items:
        platform = classify_partner(item["text"])
        for brand in classify_brands(item["text"]):
            sales_next[brand][platform].append(item)

    sales = {}
    for brand in BRAND_LABELS:
        platforms = sorted(
            set(sales_current[brand]) | set(sales_next[brand]),
            key=lambda value: (value == "기타", value),
        )
        sales[brand] = [
            {
                "id": f"{brand}-{index}",
                "platform": platform,
                "this_week": _lines(sales_current[brand][platform]),
                "next_week": _lines(sales_next[brand][platform], next_week=True),
            }
            for index, platform in enumerate(platforms, 1)
        ]

    return {
        "week_friday": friday.isoformat(),
        "period_start": start.isoformat(),
        "period_end": end.isoformat(),
        "next_period_start": next_start.isoformat(),
        "next_period_end": next_end.isoformat(),
        "source_tasks": current,
        "next_source_tasks": upcoming,
        "partners": partners,
        "sales": sales,
        "brand_labels": BRAND_LABELS,
        "summary": {
            "this_week": len(current),
            "completed": sum(item["status"] == "done" for item in current),
            "next_week": len(upcoming),
            "carry_over": len(carry_over),
        },
    }
