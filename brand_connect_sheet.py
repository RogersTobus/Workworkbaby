from __future__ import annotations


PROPOSAL_STATUS_VALUES = {"대기", "수락", "거절"}


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
