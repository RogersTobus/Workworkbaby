from __future__ import annotations


PROPOSAL_STATUS_VALUES = {"대기", "수락", "거절"}

FAVORITE_DATE_HEADERS = {
    "alp": {
        "immun": ("제안날짜(이뮨)", "제안일자(이뮨)"),
        "iron_drop": ("제안날짜(아이언드롭)", "제안일자(아이언드롭)"),
    },
}


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
