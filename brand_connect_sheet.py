from __future__ import annotations


PROPOSAL_STATUS_VALUES = {"대기", "수락", "거절"}
LEGACY_FAVORITE_VALUE = "찜"

FAVORITE_DATE_HEADERS = {
    "alp": {
        "immun": ("제안날짜(이뮨)", "제안일자(이뮨)"),
        "iron_drop": ("제안날짜(아이언드롭)", "제안일자(아이언드롭)"),
    },
    "gaia": {
        "oil": ("제안날짜(오일)", "제안일자(오일)"),
        "pickles": ("제안날짜(병절임)", "제안일자(병절임)"),
        "pouch": ("제안날짜(파우치)", "제안일자(파우치)"),
    },
}


def first_header_columns(sheet) -> dict[str, int]:
    """중복 헤더가 있으면 왼쪽의 실제 데이터 열을 우선한다."""
    headers: dict[str, int] = {}
    for column in range(1, sheet.max_column + 1):
        header = str(sheet.cell(1, column).value or "").strip()
        if header:
            headers.setdefault(header, column)
    return headers


def find_favorite_date_column(
    headers: dict[str, int],
    brand_key: str,
    product: str,
) -> int | None:
    """선택 상품 전용 제안날짜 열을 우선 찾고 기존 공통 열도 지원한다."""
    candidates = FAVORITE_DATE_HEADERS.get(brand_key, {}).get(product, ())
    for header in (*candidates, "제안날짜", "제안일자"):
        column = headers.get(header)
        if column:
            return int(column)
    return None


def is_favorite_candidate(
    proposal_date: object,
    selected_product: object,
) -> bool:
    """제안날짜와 선택 상품 열이 모두 비어 있을 때만 찜 후보로 인정한다."""
    return not str(proposal_date or "").strip() and not str(
        selected_product or ""
    ).strip()


def set_proposal_date_if_blank(
    sheet,
    row: int,
    date_column: int,
    date_text: str,
) -> bool:
    """제안날짜는 사용자가 관리하므로 자동화에서 수정하지 않는다."""
    return False


def set_favorite_date_if_blank(
    sheet,
    row: int,
    date_column: int,
    date_text: str,
) -> bool:
    """찜 자동화가 성공한 행의 빈 제안날짜만 찜 완료일로 채운다."""
    cell = sheet.cell(row, date_column)
    if str(cell.value or "").strip():
        return False
    cell.value = date_text
    return True


def row_has_proposal_status(
    sheet,
    row: int,
    product_columns: list[int],
) -> bool:
    return any(
        str(sheet.cell(row, column).value or "").strip()
        in PROPOSAL_STATUS_VALUES
        for column in product_columns
    )


def is_legacy_favorite_value(value: object) -> bool:
    return str(value or "").strip() == LEGACY_FAVORITE_VALUE


def should_add_missing_campaign_creator(
    status: object,
    relaxed_matches: list[int],
) -> bool:
    """확인된 제안 상태이며 기존 유사 행도 없을 때만 새 명단으로 추가한다."""
    return (
        str(status or "").strip() in PROPOSAL_STATUS_VALUES
        and not relaxed_matches
    )
