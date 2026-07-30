from __future__ import annotations


PROPOSAL_STATUS_VALUES = {"대기", "수락", "거절"}


def set_proposal_date_if_blank(
    sheet,
    row: int,
    date_column: int,
    date_text: str,
) -> bool:
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
