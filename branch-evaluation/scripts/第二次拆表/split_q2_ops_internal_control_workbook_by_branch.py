#!/usr/bin/env python3
"""
拆分“2026年二季度运营内控队伍履职评价基础数据V1”工作簿。

============================== 一、脚本目标 ==============================

把一份全辖 Excel 工作簿，按照“管辖行.xlsx”中的 18 个管辖行，
拆成 18 个独立 Excel 文件。

本脚本会尽量完整保留原 Excel 的：
字体、颜色、边框、列宽、行高、合并单元格、公式、隐藏行列、打印设置等。

同时，本次有几个特别要求：

1. 隐藏工作表不拆分，保持原样，也保持隐藏状态；
2. 评价方案、校验表不拆分，保持原样；
3. 可见数据表严格拆分：只保留当前管辖行数据，其他机构全部删除；
4. 多行表头完整保留，从第一行真正有管辖行归属的数据开始拆分；
5. 公式版工作簿保存时开启 Excel 打开后自动重算。

============================== 二、拆分逻辑 ==============================

普通表格优先按以下字段拆分：

    管辖行、管辖支行、所属管辖行、工作部门管辖行、人事部门管辖行等

如果某张表没有直接的管辖行字段，脚本会再尝试：

1. 按机构号拆分：
   例如“机构号”“网点机构号”“发起机构号”“创建机构”等字段。
   脚本会先从其他表中学习“机构号 -> 管辖行”的对应关系。

2. 按员工号拆分：
   例如“员工号”“员工编号”“EHR编号”“总行工号”“分行员工号”等字段。
   脚本会先从“员工信息（运营内控队伍复制）”等表中学习“员工号 -> 管辖行”。

如果可见数据表中的某一行无法判断属于当前管辖行，且不是空白/模板行，
脚本会删除它，并写入“拆分运行报告.json”。

============================== 三、不拆分的表 ==============================

以下几类工作表完整保留、不删行：

1. 所有隐藏工作表；
2. 工作表名称包含“评价方案”的工作表；
3. 工作表名称包含“校验”的工作表。

例如：
    1、运营内控副行长履职评价方案
    2、运营内控经理（含助理）履职评价方案
    1.2.2支行自主分配核准业务量积分校验
隐藏表不会取消隐藏。

============================== 四、运行方法 ==============================

首次运行前安装依赖：

    python3 -m pip install openpyxl

进入脚本所在目录：

    cd "/Users/wenzhou/Documents/履职评价拆表"

按默认路径运行：

    python3 split_q2_ops_internal_control_workbook_by_branch.py

指定其他路径运行：

    python3 split_q2_ops_internal_control_workbook_by_branch.py \\
      --source "/源文件路径.xlsx" \\
      --branches "/管辖行.xlsx路径" \\
      --output-dir "/输出目录"

默认输出目录：

    当前脚本目录/二季度运营内控拆分目录
"""

from __future__ import annotations

import argparse
import json
import re
import time
from copy import copy
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, range_boundaries


DEFAULT_SOURCE = Path("/Users/wenzhou/Downloads/附件1：2026年二季度运营内控队伍履职评价基础数据V1.xlsx")
DEFAULT_BRANCHES = Path("/Users/wenzhou/Downloads/管辖行.xlsx")
DEFAULT_OUTPUT = Path(__file__).resolve().parent / "二季度运营内控拆分目录"

# 输出文件命名配置。默认沿用“管辖行全称.xlsx”；具体入口脚本可以覆盖。
OUTPUT_FILE_PREFIX = ""
OUTPUT_USE_SHORT_BRANCH_NAME = False

# 可见工作表名称只要包含这些关键词，就完整保留、不拆分。
# 隐藏工作表无论名称是什么，都完整保留、不拆分。
PRESERVE_SHEET_KEYWORDS = ("评价方案", "校验")

# 这些字段通常直接表示一行数据属于哪个管辖行。
BRANCH_HEADER_KEYWORDS = (
    "管辖行",
    "管辖支行",
    "所属管辖行",
    "工作部门管辖行",
    "人事部门管辖行",
    "管辖行名称",
)

# 这些字段通常是机构号，可以通过“机构号 -> 管辖行”映射拆分。
ORG_HEADER_KEYWORDS = (
    "机构号",
    "网点机构号",
    "发起机构号",
    "创建机构",
    "工作部门ID",
    "机构号文本",
)

# 这些字段通常是员工号，可以通过“员工号 -> 管辖行”映射拆分。
EMPLOYEE_HEADER_KEYWORDS = (
    "员工号",
    "员工编号",
    "EHR编号",
    "EHR员工号",
    "总行工号",
    "总行员工号",
    "分行工号",
    "分行员工号",
    "工作证号",
    "柜员号",
    "柜员号文本",
)

# 这些字段通常是网点或机构名称，可以通过“网点/机构名称 -> 管辖行”映射拆分。
NAME_HEADER_KEYWORDS = (
    "网点",
    "机构名称",
    "机构",
    "所属部门",
    "工作部门",
    "派驻机构",
    # 部分加分项明细表的“管辖行、机构号、员工号”是公式，
    # 例如“6.1加分项（一、价值贡献）”通过姓名反查员工信息。
    # 所以这里把“姓名”也作为可学习、可匹配的归属字段。
    "姓名",
)


def parse_args() -> argparse.Namespace:
    """读取命令行参数；没有传参数时使用默认路径。"""
    parser = argparse.ArgumentParser(description="按管辖行拆分运营内控队伍履职评价工作簿。")
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE, help="待拆分的源工作簿")
    parser.add_argument("--branches", type=Path, default=DEFAULT_BRANCHES, help="管辖行名单工作簿")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT, help="拆分结果输出目录")
    return parser.parse_args()


def clean_text(value: Any) -> str:
    """把单元格内容转成干净文本，去掉空格、换行和制表符。"""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value))


def branch_key(value: Any) -> str:
    """
    将不同写法的管辖行统一成简称，便于匹配。

    示例：
        “宝安管辖支行” -> “宝安”
        “宝安支行” -> “宝安”
        “高新” -> “高新区”
        “分行营业部” -> “营业部”
        “[17836]中国银行深圳市分行营业部” -> “营业部”
        “河套皇岗管辖分行” -> “河套皇岗”
        “龙华管辖支行/红山支行” -> “龙华”
    """
    text = clean_text(value)
    if not text:
        return ""

    text = re.sub(r"^\[\d+\]", "", text)
    text = re.sub(r"^中国银行深圳(?:市)?", "", text)

    # 如果是“管辖行/网点”这种写法，只取斜杠前面的管辖行部分。
    if "/" in text:
        text = text.split("/", 1)[0]

    if text == "分行营业部":
        return "营业部"

    for suffix in ("管辖支行", "管辖分行", "管辖行", "支行", "分行"):
        if text.endswith(suffix):
            text = text[: -len(suffix)]
            break

    aliases = {
        "高新": "高新区",
        "中心": "中心区",
        "前海": "前海蛇口",
        "河套": "河套皇岗",
    }
    return aliases.get(text, text)


def normalize_id(value: Any) -> str:
    """
    标准化机构号或员工号。

    Excel 中编号可能是数字、文本、带制表符文本，甚至是“17747-中国银行深圳葵涌支行”。
    本函数会尽量提取开头的数字编号，并保留文本型编号的前导 0。
    """
    text = clean_text(value)
    if not text:
        return ""
    if text.endswith(".0"):
        text = text[:-2]
    match = re.match(r"^'?(\d+)", text)
    return match.group(1) if match else text


def read_branch_names(path: Path) -> list[str]:
    """读取“管辖行.xlsx”中的 18 个管辖行名称。"""
    workbook = load_workbook(path, read_only=True, data_only=True)
    names: list[str] = []
    for row in workbook.active.iter_rows():
        for cell in row:
            text = clean_text(cell.value)
            if text:
                names.append(text)
    workbook.close()

    if not names:
        raise ValueError(f"管辖行名单为空：{path}")
    if len(names) != len(set(names)):
        raise ValueError("管辖行名单存在重复，请先检查管辖行.xlsx")
    return names


def build_branch_lookup(branches: list[str]) -> dict[str, str]:
    """建立“简称 -> 正式名称”的查询字典。"""
    lookup: dict[str, str] = {}
    for branch in branches:
        key = branch_key(branch)
        if key in lookup:
            raise ValueError(f"管辖行简称冲突：{lookup[key]} / {branch}")
        lookup[key] = branch
    return lookup


def branch_from_text(value: Any, branch_lookup: dict[str, str]) -> str | None:
    """
    尝试直接从一段文字中判断管辖行。

    例如“宝安支行风险与运营管理部”没有单独的管辖行字段，
    但文字中包含“宝安支行”，可以推断为“宝安管辖支行”。
    """
    text = clean_text(value)
    if not text:
        return None

    direct = branch_lookup.get(branch_key(text))
    if direct:
        return direct

    for key, branch in branch_lookup.items():
        if key and key in text:
            return branch

    return None


def should_preserve_sheet(sheet_name: str) -> bool:
    """判断一张表是否属于“完整保留、不拆分”的表。"""
    return any(keyword in sheet_name for keyword in PRESERVE_SHEET_KEYWORDS)


def relevant_cells(worksheet, max_header_row: int = 12) -> list[tuple[int, int, Any]]:
    """
    读取表头区域内所有有值的单元格。

    不直接用 worksheet.max_column 全表扫描，是因为有些工作表虽然实际只有十几列，
    但 Excel 内部记录的最大列可能到 XFD，直接扫描会很慢。
    """
    cells: list[tuple[int, int, Any]] = []
    for (row, column), cell in worksheet._cells.items():
        if row <= max_header_row and cell.value is not None:
            cells.append((row, column, cell.value))
    return sorted(cells)


def header_matches(value: Any, keywords: tuple[str, ...]) -> bool:
    """判断表头文字是否包含某类关键词。"""
    text = clean_text(value)
    if not text:
        return False
    if len(text) > 40:
        return False
    return any(keyword in text for keyword in keywords)


def find_columns(worksheet) -> dict[str, list[tuple[int, int, str]]]:
    """
    找出一张表里可能用于拆分的列。

    返回：
        branch: 管辖行类字段
        org: 机构号类字段
        employee: 员工号类字段
        name: 网点/机构名称类字段
    """
    result = {"branch": [], "org": [], "employee": [], "name": []}

    for row, column, value in relevant_cells(worksheet):
        text = clean_text(value)

        # 排名、得分、说明类字段不是归属字段，跳过。
        if "排名" in text or "得分" in text or "说明" in text:
            continue

        if header_matches(value, BRANCH_HEADER_KEYWORDS):
            result["branch"].append((row, column, text))
        elif header_matches(value, ORG_HEADER_KEYWORDS):
            result["org"].append((row, column, text))
        elif header_matches(value, EMPLOYEE_HEADER_KEYWORDS):
            result["employee"].append((row, column, text))
        elif header_matches(value, NAME_HEADER_KEYWORDS):
            if "管辖" not in text and "员工" not in text and "工号" not in text:
                result["name"].append((row, column, text))

    return result


def choose_header_row(columns: dict[str, list[tuple[int, int, str]]]) -> int | None:
    """根据找到的字段，推断数据表头所在行。"""
    candidates = columns["branch"] or columns["org"] or columns["employee"] or columns["name"]
    if not candidates:
        return None
    return min(row for row, _, _ in candidates)


def build_lookup_maps(workbook, branch_lookup: dict[str, str]) -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    """
    从源工作簿学习两个映射关系：

    1. org_to_branch：机构号 -> 管辖行
    2. employee_to_branch：员工号 -> 管辖行
    3. name_to_branch：网点/机构名称 -> 管辖行

    这样没有直接“管辖行”字段的隐藏表，也可以尽量按机构号、员工号或网点名称拆分。
    """
    org_to_branch: dict[str, str] = {}
    employee_to_branch: dict[str, str] = {}
    name_to_branch: dict[str, str] = {}

    for worksheet in workbook.worksheets:
        columns = find_columns(worksheet)
        header_row = choose_header_row(columns)
        if header_row is None:
            continue

        branch_columns = [column for row, column, _ in columns["branch"] if row == header_row]
        org_columns = [column for row, column, _ in columns["org"] if row == header_row]
        employee_columns = [column for row, column, _ in columns["employee"] if row == header_row]
        name_columns = []
        for row, column, value in relevant_cells(worksheet):
            text = clean_text(value)
            if row == header_row and header_matches(value, NAME_HEADER_KEYWORDS):
                if "管辖" not in text and "员工" not in text and "工号" not in text:
                    name_columns.append(column)

        if not branch_columns:
            continue

        branch_column = branch_columns[0]
        for row_number in range(header_row + 1, worksheet.max_row + 1):
            branch = branch_lookup.get(branch_key(worksheet.cell(row_number, branch_column).value))
            if not branch:
                continue

            for org_column in org_columns:
                org_id = normalize_id(worksheet.cell(row_number, org_column).value)
                if org_id:
                    org_to_branch.setdefault(org_id, branch)

            for employee_column in employee_columns:
                employee_id = normalize_id(worksheet.cell(row_number, employee_column).value)
                if employee_id:
                    employee_to_branch.setdefault(employee_id, branch)

            for name_column in name_columns:
                name = clean_text(worksheet.cell(row_number, name_column).value)
                if name:
                    name_to_branch.setdefault(name, branch)

    return org_to_branch, employee_to_branch, name_to_branch


def row_branch(
    worksheet,
    row_number: int,
    columns: dict[str, list[tuple[int, int, str]]],
    branch_lookup: dict[str, str],
    org_to_branch: dict[str, str],
    employee_to_branch: dict[str, str],
    name_to_branch: dict[str, str],
) -> tuple[str | None, str]:
    """
    判断某一行属于哪个管辖行。

    返回：
        (管辖行正式名称, 匹配方式)

    如果无法判断，返回 (None, "")。
    """
    for _, column, _ in columns["branch"]:
        value = worksheet.cell(row_number, column).value
        # 公式单元格不是稳定的归属文本，跳过，改用机构号或员工号判断。
        if isinstance(value, str) and value.startswith("="):
            continue
        branch = branch_lookup.get(branch_key(value))
        if branch:
            return branch, "branch"
        branch = branch_from_text(value, branch_lookup)
        if branch:
            return branch, "branch_text"

    for _, column, _ in columns["org"]:
        org_id = normalize_id(worksheet.cell(row_number, column).value)
        branch = org_to_branch.get(org_id)
        if branch:
            return branch, "org"

    for _, column, _ in columns["employee"]:
        employee_id = normalize_id(worksheet.cell(row_number, column).value)
        branch = employee_to_branch.get(employee_id)
        if branch:
            return branch, "employee"

    for row, column, value in relevant_cells(worksheet):
        text = clean_text(value)
        if not header_matches(value, NAME_HEADER_KEYWORDS):
            continue
        if "管辖" in text or "员工" in text or "工号" in text:
            continue
        name = clean_text(worksheet.cell(row_number, column).value)
        branch = name_to_branch.get(name) or branch_from_text(name, branch_lookup)
        if branch:
            return branch, "name"

    return None, ""


def row_has_any_data(worksheet, row_number: int) -> bool:
    """判断一行是否有内容，用于区分空白模板行和无法识别的数据行。"""
    for (row, _), cell in worksheet._cells.items():
        if row == row_number and cell.value not in (None, ""):
            return True
    return False


def rows_with_any_data(worksheet) -> set[int]:
    """一次性返回所有含值行，避免大表逐行反复扫描全部单元格。"""
    return {
        row
        for (row, _), cell in worksheet._cells.items()
        if cell.value not in (None, "")
    }


def find_data_start_row(
    worksheet,
    header_row: int,
    columns: dict[str, list[tuple[int, int, str]]],
    branch_lookup: dict[str, str],
    org_to_branch: dict[str, str],
    employee_to_branch: dict[str, str],
    name_to_branch: dict[str, str],
) -> int:
    """
    推断真正的数据起始行。

    为什么不能简单地从“表头字段行 + 1”开始拆：
        有些工作表的表头不是一行，而是多行。
        例如第 2 行写“管辖行、姓名、员工号”，第 3-4 行继续写分组表头、
        指标说明或合并单元格说明，第 5 行才是第一条人员数据。

    本函数的做法：
        从字段行的下一行开始往下找；
        找到第一行能判断出属于某个管辖行的数据行；
        这行才是真正的数据起始行。

    这样第 1 行到数据起始行前一行都会被当作表头/说明区完整保留，
    颜色、宽度、合并单元格、公式和说明文字都不动。
    """
    for row_number in range(header_row + 1, worksheet.max_row + 1):
        matched_branch, _ = row_branch(
            worksheet,
            row_number,
            columns,
            branch_lookup,
            org_to_branch,
            employee_to_branch,
            name_to_branch,
        )
        if matched_branch is not None:
            return row_number

    # 如果整张表都没有找到可识别归属的数据，就退回到原来的保守起点。
    # 后续逻辑会把无法识别的非空行写入报告，便于人工检查。
    return header_row + 1


def copy_row_dimensions(worksheet, old_to_new: dict[int, int]) -> None:
    """删除行后恢复保留下来的行高、隐藏状态等行属性。"""
    original = {row: copy(dimension) for row, dimension in worksheet.row_dimensions.items()}

    for row in list(worksheet.row_dimensions):
        del worksheet.row_dimensions[row]

    for old_row, new_row in old_to_new.items():
        if old_row in original:
            worksheet.row_dimensions[new_row] = copy(original[old_row])


def translate_moved_formulas(worksheet, new_to_old: dict[int, int]) -> None:
    """
    行号移动后，调整行内相对引用公式。

    例如原来第 20 行的 “=B20/C20” 移到第 6 行后，
    会调整为 “=B6/C6”。
    """
    for new_row, old_row in new_to_old.items():
        if new_row == old_row:
            continue
        for cell in worksheet[new_row]:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                old_coordinate = f"{cell.column_letter}{old_row}"
                try:
                    cell.value = Translator(cell.value, origin=old_coordinate).translate_formula(cell.coordinate)
                except Exception:
                    # 少数公式 openpyxl 无法解析时，保留原公式，避免公式丢失。
                    pass


def copy_cell(source_cell, target_cell, source_coordinate: str) -> None:
    """
    把一个单元格完整复制到另一个单元格。

    这里不仅复制值，也复制样式、数字格式、批注、超链接等信息。
    如果复制的是公式，并且行号发生变化，还会把公式中的相对引用同步平移。
    """
    value = source_cell.value
    if isinstance(value, str) and value.startswith("="):
        try:
            value = Translator(value, origin=source_coordinate).translate_formula(target_cell.coordinate)
        except Exception:
            pass

    target_cell.value = value
    if source_cell.has_style:
        target_cell._style = copy(source_cell._style)
    if source_cell.number_format:
        target_cell.number_format = source_cell.number_format
    if source_cell.comment:
        target_cell.comment = copy(source_cell.comment)
    if source_cell.hyperlink:
        target_cell._hyperlink = copy(source_cell.hyperlink)


def compact_rows(worksheet, old_to_new: dict[int, int], original_max_row: int, fixed_header_rows: int) -> None:
    """
    将要保留的行压缩到上方，然后一次性删除尾部多余行。

    为什么不用 delete_rows 一段段删除：
        大明细表可能有两万多行，而且目标管辖行的数据分散在很多位置。
        如果一段段删除，Excel 内部要反复搬移单元格，非常慢。

    本函数的做法：
        1. 先记住每个保留行的所有单元格；
        2. 把这些保留行按顺序复制到新的行号；
        3. 一次性删除最后多出来的尾部行。

    这样通常比大量 delete_rows 快很多。
    """
    cells_by_old_row: dict[int, list[tuple[int, Any, str]]] = {}
    data_merges_to_restore: list[str] = []

    # openpyxl 删除行时不会自动同步合并区域。先解除数据区的合并，
    # 记录保留行对应的新位置，压缩完成后再按新行号恢复。
    # 否则旧的合并区域仍可能延伸到原来的最后一行，造成大量视觉空行。
    for merged_range in list(worksheet.merged_cells.ranges):
        if merged_range.min_row <= fixed_header_rows:
            continue

        source_rows = range(merged_range.min_row, merged_range.max_row + 1)
        mapped_rows = [old_to_new.get(row) for row in source_rows]
        worksheet.unmerge_cells(str(merged_range))

        if all(row is not None for row in mapped_rows):
            new_rows = [int(row) for row in mapped_rows]
            if new_rows == list(range(new_rows[0], new_rows[-1] + 1)):
                data_merges_to_restore.append(
                    f"{get_column_letter(merged_range.min_col)}{new_rows[0]}:"
                    f"{get_column_letter(merged_range.max_col)}{new_rows[-1]}"
                )

    for (row, column), cell in list(worksheet._cells.items()):
        # 表头区域不需要移动，保留原样，避免影响合并单元格。
        if row > fixed_header_rows and row in old_to_new:
            cells_by_old_row.setdefault(row, []).append((column, copy(cell), cell.coordinate))

    # 先删除将要写入区域的数据行单元格，避免目标行残留旧内容。
    data_start = fixed_header_rows + 1
    data_end = max(old_to_new.values())
    for key in list(worksheet._cells):
        row, _ = key
        if data_start <= row <= data_end:
            del worksheet._cells[key]

    for old_row in sorted(old_to_new):
        if old_row <= fixed_header_rows:
            continue
        new_row = old_to_new[old_row]
        for column, source_cell, source_coordinate in cells_by_old_row.get(old_row, []):
            target_cell = worksheet.cell(new_row, column)
            copy_cell(source_cell, target_cell, source_coordinate)

    copy_row_dimensions(worksheet, old_to_new)

    first_unused_row = len(old_to_new) + 1
    if first_unused_row <= original_max_row:
        worksheet.delete_rows(first_unused_row, original_max_row - first_unused_row + 1)

    for merged_range in data_merges_to_restore:
        worksheet.merge_cells(merged_range)

    # 同步普通筛选区域的最后一行，避免筛选下拉仍覆盖原工作表的空白尾行。
    if worksheet.auto_filter.ref:
        try:
            min_col, min_row, max_col, _ = range_boundaries(worksheet.auto_filter.ref)
            worksheet.auto_filter.ref = (
                f"{get_column_letter(min_col)}{min_row}:"
                f"{get_column_letter(max_col)}{len(old_to_new)}"
            )
        except (TypeError, ValueError):
            # 少数文件的筛选引用不是标准 CellRange，保持源设置。
            pass


def clear_visible_formula_errors(worksheet) -> int:
    """
    清理可见拆分表中的明显错误公式。

    这份源文件的可见表中存在极少量原始 #REF! 公式。
    用户要求最终文件中有公式的地方不要错，所以拆分后将这类公式清空，
    避免保留坏公式。
    """
    cleared = 0
    for row in worksheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("=") and "#REF!" in cell.value:
                cell.value = None
                cleared += 1
    return cleared


def delete_unwanted_rows(worksheet, first_data_row: int, last_row: int, keep_rows: set[int]) -> None:
    """按连续区间批量删除不需要的行，比逐行删除快很多。"""
    delete_blocks: list[tuple[int, int]] = []
    start: int | None = None
    end: int | None = None

    for row in range(first_data_row, last_row + 1):
        if row not in keep_rows:
            if start is None:
                start = row
            end = row
        elif start is not None and end is not None:
            delete_blocks.append((start, end))
            start = None
            end = None

    if start is not None and end is not None:
        delete_blocks.append((start, end))

    for start_row, end_row in reversed(delete_blocks):
        worksheet.delete_rows(start_row, end_row - start_row + 1)


def split_sheet(
    worksheet,
    branch: str,
    branch_lookup: dict[str, str],
    org_to_branch: dict[str, str],
    employee_to_branch: dict[str, str],
    name_to_branch: dict[str, str],
) -> dict[str, Any]:
    """把一张工作表筛选成当前管辖行的数据。"""
    if worksheet.sheet_state != "visible":
        return {
            "action": "preserved_hidden",
            "reason": "隐藏工作表按用户要求不拆分，保持原隐藏状态",
        }

    if should_preserve_sheet(worksheet.title):
        return {"action": "preserved", "reason": "评价方案/校验表不拆分"}

    columns = find_columns(worksheet)
    header_row = choose_header_row(columns)

    if header_row is None:
        return {"action": "no_split_key", "reason": "未找到管辖行、机构号或员工号字段"}

    data_start_row = find_data_start_row(
        worksheet,
        header_row,
        columns,
        branch_lookup,
        org_to_branch,
        employee_to_branch,
        name_to_branch,
    )

    # 数据起始行之前都视为表头/说明区，完整保留，不参与管辖行筛选。
    keep_rows = set(range(1, data_start_row))
    data_rows_with_values = rows_with_any_data(worksheet)
    deleted_unmatched_rows: list[dict[str, Any]] = []
    matched_by = {"branch": 0, "branch_text": 0, "org": 0, "employee": 0, "name": 0}

    for row_number in range(data_start_row, worksheet.max_row + 1):
        matched_branch, method = row_branch(
            worksheet,
            row_number,
            columns,
            branch_lookup,
            org_to_branch,
            employee_to_branch,
            name_to_branch,
        )

        if matched_branch == branch:
            keep_rows.add(row_number)
            matched_by[method] += 1
        elif matched_branch is None:
            if row_number in data_rows_with_values:
                # 严格模式：无法判断归属的非空行也删除，并写入报告。
                deleted_unmatched_rows.append({"row": row_number})
            # 数据区中的完全空白行不再保留，避免拆分后的子表留下大量空行。
            # 数据起始行之前的多行表头和说明区仍由 keep_rows 完整保留。
        # matched_branch 是其他管辖行时，不加入 keep_rows，稍后删除。

    original_max_row = worksheet.max_row
    old_to_new: dict[int, int] = {}
    next_row = 1
    for old_row in range(1, original_max_row + 1):
        if old_row in keep_rows:
            old_to_new[old_row] = next_row
            next_row += 1

    compact_rows(worksheet, old_to_new, original_max_row, fixed_header_rows=data_start_row - 1)
    cleared_formula_errors = clear_visible_formula_errors(worksheet)

    return {
        "action": "split",
        "header_row": header_row,
        "data_start_row": data_start_row,
        "branch_columns": columns["branch"],
        "org_columns": columns["org"],
        "employee_columns": columns["employee"],
        "name_columns": columns["name"],
        "kept_rows": len(keep_rows),
        "removed_rows": original_max_row - len(keep_rows),
        "removed_blank_rows": sum(
            1
            for row_number in range(data_start_row, original_max_row + 1)
            if row_number not in data_rows_with_values
        ),
        "matched_by": matched_by,
        "deleted_unmatched_rows": deleted_unmatched_rows,
        "cleared_formula_errors": cleared_formula_errors,
    }


def split_workbook(source: Path, branches_path: Path, output_dir: Path) -> dict[str, Any]:
    """执行完整拆分任务。"""
    if not source.is_file():
        raise FileNotFoundError(f"找不到源文件：{source}")
    if not branches_path.is_file():
        raise FileNotFoundError(f"找不到管辖行名单：{branches_path}")

    started_at = time.perf_counter()
    branches = read_branch_names(branches_path)
    branch_lookup = build_branch_lookup(branches)

    print(f"开始拆分，共 {len(branches)} 个管辖行...", flush=True)
    print("正在建立机构号/员工号到管辖行的映射...", flush=True)

    source_workbook = load_workbook(source, data_only=False)
    org_to_branch, employee_to_branch, name_to_branch = build_lookup_maps(source_workbook, branch_lookup)
    source_workbook.close()

    print(
        f"已识别机构号映射 {len(org_to_branch)} 条，"
        f"员工号映射 {len(employee_to_branch)} 条，"
        f"网点/机构名称映射 {len(name_to_branch)} 条。",
        flush=True,
    )

    output_dir.mkdir(parents=True, exist_ok=True)
    report: dict[str, Any] = {
        "source": str(source),
        "branches_file": str(branches_path),
        "output_dir": str(output_dir),
        "branch_count": len(branches),
        "preserve_keywords": PRESERVE_SHEET_KEYWORDS,
        "org_mapping_count": len(org_to_branch),
        "employee_mapping_count": len(employee_to_branch),
        "name_mapping_count": len(name_to_branch),
        "files": [],
        "sheets_without_split_key": [],
        "deleted_unmatched_rows": {},
    }

    for index, branch in enumerate(branches, start=1):
        branch_started_at = time.perf_counter()
        print(f"[{index:02d}/{len(branches)}] 正在处理：{branch}", flush=True)

        workbook = load_workbook(source, data_only=False)
        sheet_results: dict[str, Any] = {}

        for worksheet in workbook.worksheets:
            result = split_sheet(
                worksheet,
                branch,
                branch_lookup,
                org_to_branch,
                employee_to_branch,
                name_to_branch,
            )
            sheet_results[worksheet.title] = result

            if result["action"] == "no_split_key" and worksheet.title not in report["sheets_without_split_key"]:
                report["sheets_without_split_key"].append(worksheet.title)

            if result.get("deleted_unmatched_rows"):
                report["deleted_unmatched_rows"].setdefault(worksheet.title, result["deleted_unmatched_rows"])

        output_branch_name = branch_key(branch) if OUTPUT_USE_SHORT_BRANCH_NAME else branch
        if OUTPUT_FILE_PREFIX:
            output_filename = f"{OUTPUT_FILE_PREFIX}（{output_branch_name}）.xlsx"
        else:
            output_filename = f"{output_branch_name}.xlsx"
        output_path = output_dir / output_filename
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
        workbook.save(output_path)
        workbook.close()

        print(
            f"[{index:02d}/{len(branches)}] 已完成：{branch}（{time.perf_counter() - branch_started_at:.1f} 秒）",
            flush=True,
        )
        report["files"].append({"branch": branch, "path": str(output_path), "sheets": sheet_results})

    report_path = output_dir / "拆分运行报告.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"全部完成，总耗时 {time.perf_counter() - started_at:.1f} 秒。", flush=True)
    return report


def main() -> None:
    """程序入口。"""
    args = parse_args()
    report = split_workbook(args.source.resolve(), args.branches.resolve(), args.output_dir.resolve())

    print(f"已生成 {len(report['files'])} 个工作簿：{report['output_dir']}")
    if report["sheets_without_split_key"]:
        print("以下可见工作表未找到可拆分字段，已保留原样：")
        for sheet_name in report["sheets_without_split_key"]:
            print(f"  - {sheet_name}")
    if report["deleted_unmatched_rows"]:
        print("存在无法判断归属的数据行，已按严格规则删除；详见拆分运行报告.json")


if __name__ == "__main__":
    main()
