from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


@dataclass(frozen=True)
class Block:
    r1: int
    c1: int
    r2: int
    c2: int
    text: str
    fill: Optional[str] = None
    font_color: Optional[str] = None
    bold: bool = False
    rotation: int = 0


def _apply_border(ws, r1: int, c1: int, r2: int, c2: int, border: Border) -> None:
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = border


def _apply_fill(ws, r1: int, c1: int, r2: int, c2: int, fill: Optional[PatternFill]) -> None:
    if fill is None:
        return
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).fill = fill


def _apply_alignment(ws, r1: int, c1: int, r2: int, c2: int, alignment: Alignment) -> None:
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).alignment = alignment


def _apply_font(ws, r1: int, c1: int, r2: int, c2: int, font: Font) -> None:
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).font = font


def _merge_and_set(ws, block: Block, border: Border) -> None:
    ws.merge_cells(start_row=block.r1, start_column=block.c1, end_row=block.r2, end_column=block.c2)
    cell = ws.cell(block.r1, block.c1)
    cell.value = block.text

    fill = PatternFill("solid", fgColor=block.fill) if block.fill else None
    _apply_fill(ws, block.r1, block.c1, block.r2, block.c2, fill)

    alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
        text_rotation=block.rotation,
    )
    _apply_alignment(ws, block.r1, block.c1, block.r2, block.c2, alignment)

    font = Font(
        name="Calibri",
        size=11,
        bold=block.bold,
        color=(block.font_color if block.font_color else None),
    )
    _apply_font(ws, block.r1, block.c1, block.r2, block.c2, font)

    _apply_border(ws, block.r1, block.c1, block.r2, block.c2, border)


def build_xlsx(out_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Flow"

    thin = Side(style="thin", color="404040")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Column setup: A is vertical section labels, B-E are swimlanes.
    col_widths = {
        1: 6,   # A
        2: 22,  # B
        3: 22,  # C
        4: 22,  # D
        5: 22,  # E
    }
    for c, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    # Row heights tuned to approximate the original layout.
    for r in range(1, 55):
        ws.row_dimensions[r].height = 18
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 22

    header_fill = "D9E1F2"
    lane_fill = "F2F2F2"
    section_fill = "EDEDED"

    blocks: list[Block] = []

    # Title
    blocks.append(
        Block(
            r1=1,
            c1=2,
            r2=1,
            c2=5,
            text="Toastmasters International 厦门秋季峰会财务收支流程及职责分工",
            fill=header_fill,
            bold=True,
        )
    )

    # Lane headers
    blocks.extend(
        [
            Block(2, 2, 2, 2, "售票收款人/出纳", fill=lane_fill, bold=True),
            Block(2, 3, 2, 3, "财务预算/会计", fill=lane_fill, bold=True),
            Block(2, 4, 2, 4, "报销人", fill=lane_fill, bold=True),
            Block(2, 5, 2, 5, "财务审核小组\n(主席团)", fill=lane_fill, bold=True),
        ]
    )

    # Section 1: 收款流程
    blocks.append(Block(3, 1, 17, 1, "收款流程", fill=section_fill, bold=True, rotation=90))

    # Note: Shapes/arrows in the original diagram are approximated as positioned blocks.
    # Lane 1 (售票收款人/出纳)
    blocks.extend(
        [
            Block(4, 2, 5, 2, "售票"),
            Block(7, 2, 8, 2, "整理收款明细"),
            Block(10, 2, 11, 2, "核对购买邮件"),
            Block(13, 2, 14, 2, "◇ 是否一致"),
            Block(16, 2, 17, 2, "财务确认"),
        ]
    )

    # Lane 2 (财务预算/会计)
    blocks.extend(
        [
            Block(4, 3, 5, 3, "结束"),
            Block(7, 3, 8, 3, "记录"),
            Block(12, 3, 12, 4, "否（核对票类别/有疑问邮件确认）", fill="FFF2CC"),
        ]
    )

    # Section 2: 支出申请流程（统筹组）
    blocks.append(Block(18, 1, 30, 1, "支出申请流程\n（统筹组）", fill=section_fill, bold=True, rotation=90))

    # Lane 3 (报销人)
    blocks.extend(
        [
            Block(19, 4, 20, 4, "支出申请"),
            Block(22, 4, 24, 4, "填写项目\n支出内容\n支出金额"),
            Block(26, 4, 27, 4, "补充开支"),
        ]
    )

    # Lane 2 (财务预算/会计)
    blocks.extend(
        [
            Block(22, 3, 23, 3, "支出审核"),
            Block(25, 3, 26, 3, "◇ 审核通过"),
            Block(28, 3, 29, 3, "登记\n支出项目明细"),
        ]
    )

    # Lane 4 (财务审核小组)
    blocks.extend(
        [
            Block(25, 5, 26, 5, "◇ 审核"),
            Block(28, 5, 29, 5, "支出审核"),
            Block(28, 4, 29, 4, "超出预算或\n不属于预算项目", fill="F8CBAD"),
            Block(25, 4, 25, 5, "不通过", fill="F8CBAD"),
        ]
    )

    # Section 3: 预算、核销和报销流程（统筹组）
    blocks.append(Block(31, 1, 50, 1, "预算、核销和报销流程\n（统筹组）", fill=section_fill, bold=True, rotation=90))

    # Lane 1
    blocks.extend(
        [
            Block(33, 2, 34, 2, "登记项目明细"),
            Block(36, 2, 37, 2, "支出款项"),
        ]
    )

    # Lane 2
    blocks.extend(
        [
            Block(33, 3, 34, 3, "核对已审核项目，\n更新状态"),
            Block(36, 3, 37, 3, "记账"),
            Block(39, 3, 40, 3, "每月底\n与出纳对账"),
            Block(42, 3, 43, 3, "每月底\n出具财务报告"),
        ]
    )

    # Lane 3
    blocks.extend(
        [
            Block(33, 4, 34, 4, "开立收款发票"),
            Block(36, 4, 37, 4, "补收发票"),
        ]
    )

    # Apply all blocks
    for b in blocks:
        _merge_and_set(ws, b, border)

    # Draw background grid borders for the full used area (helps the swimlane feel)
    used_r1, used_c1, used_r2, used_c2 = 1, 1, 50, 5
    _apply_border(ws, used_r1, used_c1, used_r2, used_c2, border)

    # Print setup
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4

    wb.save(out_path)


if __name__ == "__main__":
    build_xlsx("/workspace/toastmasters_finance_flow.xlsx")
