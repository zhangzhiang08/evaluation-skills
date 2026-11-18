#!/usr/bin/env python3
"""
按管辖行拆分履职评价汇总工作簿。

一、整体处理逻辑
----------------
1. 从“管辖行.xlsx”读取 18 个管辖行名称。
   该文件只提供名单，不会被拆分，也不会写入结果文件。
2. 针对名单中的每个管辖行，重新打开一次汇总工作簿。
3. 遇到 PRESERVE_SHEETS 中的工作表时，整张表原样保留、不拆分。
4. 其他工作表自动寻找拆分字段：
   “管辖支行” > “管辖行” > “责任机构” > “所属机构”。
5. 保留当前管辖行的数据，删除明确属于其他管辖行的数据。
6. 保存为“管辖行名称.xlsx”，最终生成 18 个工作簿。
7. 生成“拆分运行报告.json”，记录各表的处理情况和无法匹配的数据。

二、为什么使用“删除其他管辖行数据”的方式
----------------------------------------
直接在原工作簿副本中删除不需要的行，可以最大程度保留原 Excel 中的：
字体、颜色、边框、列宽、行高、合并单元格、公式、打印设置和数据验证等内容。

三、小白最常修改的位置
----------------------
1. 哪些工作表不拆分：修改 PRESERVE_SHEETS。
2. 可识别的拆分字段：修改 FIELD_PRIORITY。
3. 默认文件路径：修改 DEFAULT_SOURCE、DEFAULT_BRANCHES、DEFAULT_OUTPUT。

四、运行方式
------------
使用默认路径：
    python3 split_workbook_by_branch.py

指定其他文件路径：
    python3 split_workbook_by_branch.py \
        --source "/汇总文件路径/汇总.xlsx" \
        --branches "/名单文件路径/管辖行.xlsx" \
        --output-dir "/结果输出目录"

运行环境：
    Python 3.9 或更高版本
    openpyxl：python3 -m pip install openpyxl
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


# 不同工作表使用的字段名称不完全相同，因此按以下优先级寻找拆分字段。
# 如果一张表同时出现多个候选字段，会优先使用排在前面的字段。
FIELD_PRIORITY = ("管辖支行", "管辖行", "责任机构", "所属机构")

# 以下工作表在每个结果文件中都完整保留，不执行任何筛选或删除。
# 后续如需增加“不拆分”的工作表，只需在这里增加工作表名称。
PRESERVE_SHEETS = {
    "1.评定表格说明",
    "2-1 校验表",
    "2-2 综合服务经理履职评价方案",
}

# 默认文件路径。也可以在运行命令中通过参数指定其他路径。
DEFAULT_SOURCE = Path(
    "/Users/wenzhou/Downloads/2026年一季度综服经理队伍履职评价基础数据表（汇总）.xlsx"
)
DEFAULT_BRANCHES = Path("/Users/wenzhou/Downloads/管辖行.xlsx")
DEFAULT_OUTPUT = Path(__file__).resolve().parent / "拆分目录"


def parse_args() -> argparse.Namespace:
    """读取命令行参数；没有传入参数时使用上面的默认路径。"""
    parser = argparse.ArgumentParser(description="将履职评价汇总工作簿按管辖行拆分。")
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE, help="待拆分的汇总工作簿")
    parser.add_argument("--branches", type=Path, default=DEFAULT_BRANCHES, help="管辖行名单工作簿")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT, help="拆分结果输出目录")
    return parser.parse_args()


def check_input_files(source: Path, branches_path: Path) -> None:
    """运行前检查输入文件是否存在，避免出现难以理解的报错。"""
    if not source.is_file():
        raise FileNotFoundError(f"找不到汇总工作簿：{source}")
    if not branches_path.is_file():
        raise FileNotFoundError(f"找不到管辖行名单工作簿：{branches_path}")


def read_branch_names(path: Path) -> list[str]:
    """
    从管辖行名单文件的第一个工作表中读取所有非空单元格。

    返回示例：
        ["宝安管辖支行", "上步管辖支行", ..., "营业部"]
    """
    workbook = load_workbook(path, read_only=True, data_only=True)
    names: list[str] = []

    for row in workbook.active.iter_rows():
        for cell in row:
            if cell.value is not None and str(cell.value).strip():
                names.append(str(cell.value).strip())

    workbook.close()

    if not names:
        raise ValueError(f"管辖行名单为空：{path}")
    if len(names) != len(set(names)):
        raise ValueError("管辖行名单存在重复名称，请先检查管辖行.xlsx")

    return names


def branch_key(value: Any) -> str:
    """
    将不同写法的管辖行名称统一为可比较的简称。

    示例：
        “罗湖管辖支行”                       -> “罗湖”
        “[17815]中国银行深圳罗湖管辖支行”    -> “罗湖”
        “河套皇岗分行”                       -> “河套皇岗”

    这样即使不同工作表中的名称格式不一致，也能正确匹配。
    """
    if value is None:
        return ""

    text = re.sub(r"\s+", "", str(value))
    text = re.sub(r"^\[\d+\]", "", text)
    text = re.sub(r"^中国银行深圳(?:市)?", "", text)

    for suffix in ("管辖支行", "分行"):
        if text.endswith(suffix):
            text = text[: -len(suffix)]
            break

    return text


def build_branch_lookup(branches: list[str]) -> dict[str, str]:
    """
    建立“标准化简称 -> 名单中的正式名称”的查询字典。

    示例：
        {"宝安": "宝安管辖支行", "河套皇岗": "河套皇岗分行"}
    """
    lookup: dict[str, str] = {}

    for branch in branches:
        key = branch_key(branch)
        if key in lookup:
            raise ValueError(f"管辖行简称冲突：{lookup[key]} / {branch}")
        lookup[key] = branch

    return lookup


def find_split_field(worksheet) -> tuple[int, int, str] | None:
    """
    在工作表中寻找可用于拆分的字段。

    返回值：
        (表头所在行号, 字段所在列号, 字段名称)

    如果找不到任何候选字段，则返回 None，该工作表会保持原样。
    """
    found_fields: dict[str, tuple[int, int, str]] = {}

    for row in worksheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                value = cell.value.strip()
                if value in FIELD_PRIORITY and value not in found_fields:
                    found_fields[value] = (cell.row, cell.column, value)

    for field_name in FIELD_PRIORITY:
        if field_name in found_fields:
            return found_fields[field_name]

    return None


def copy_row_dimensions(worksheet, old_to_new: dict[int, int]) -> None:
    """
    删除行后，恢复保留下来的行高、隐藏状态等行属性。

    old_to_new 表示：
        删除前行号 -> 删除后行号
    """
    original_dimensions = {
        row_number: copy(dimension)
        for row_number, dimension in worksheet.row_dimensions.items()
    }

    for row_number in list(worksheet.row_dimensions):
        del worksheet.row_dimensions[row_number]

    for old_row, new_row in old_to_new.items():
        if old_row in original_dimensions:
            worksheet.row_dimensions[new_row] = copy(original_dimensions[old_row])


def translate_moved_formulas(worksheet, new_to_old: dict[int, int]) -> None:
    """
    数据行移动后，调整该行公式中的相对引用。

    示例：
        原第 20 行被移动到第 6 行时，
        “=E20/D20”会调整为“=E6/D6”。

    说明页、校验表、方案页不会进入本函数，因此会保持完全原样。
    """
    for new_row, old_row in new_to_old.items():
        if new_row == old_row:
            continue

        for cell in worksheet[new_row]:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                old_coordinate = f"{cell.column_letter}{old_row}"
                try:
                    cell.value = Translator(
                        cell.value,
                        origin=old_coordinate,
                    ).translate_formula(cell.coordinate)
                except Exception:
                    # 少数 Excel 扩展公式无法由 openpyxl 解析时，保留原公式，
                    # 防止因自动转换失败而丢失公式。
                    pass


def delete_unwanted_rows(
    worksheet,
    first_data_row: int,
    last_row: int,
    keep_rows: set[int],
) -> None:
    """
    从工作表底部开始，按连续区间批量删除不需要的数据行。

    为什么从底部删除：
        如果先删除上方行，下方行号会立即变化，后续删除位置就会出错。

    为什么批量删除：
        几千行的大表如果逐行删除会非常慢；连续行一次删除速度更快。
    """
    delete_blocks: list[tuple[int, int]] = []
    block_start: int | None = None
    block_end: int | None = None

    # 找出所有需要删除的连续行区间。
    for row_number in range(first_data_row, last_row + 1):
        if row_number not in keep_rows:
            if block_start is None:
                block_start = row_number
            block_end = row_number
        elif block_start is not None and block_end is not None:
            delete_blocks.append((block_start, block_end))
            block_start = None
            block_end = None

    if block_start is not None and block_end is not None:
        delete_blocks.append((block_start, block_end))

    # 从底部开始删除，确保行号不会影响尚未执行的删除区间。
    for start_row, end_row in reversed(delete_blocks):
        worksheet.delete_rows(start_row, end_row - start_row + 1)


def filter_sheet(worksheet, branch: str, branch_lookup: dict[str, str]) -> dict[str, Any]:
    """
    将一张普通数据表筛选为仅包含当前管辖行的数据。

    保留规则：
    1. 表头以及表头上方内容全部保留；
    2. 属于当前管辖行的数据保留；
    3. 空白模板行、分隔行、附注行保留；
    4. 无法匹配 18 个管辖行的数据保留，并记录到报告；
    5. 明确属于其他管辖行的数据删除。
    """
    split_field = find_split_field(worksheet)

    # 找不到拆分字段时，不对该工作表进行任何修改。
    if split_field is None:
        return {
            "action": "preserved",
            "reason": "未找到可用于拆分的字段",
            "field": None,
            "kept": None,
            "removed": None,
            "unmatched": [],
        }

    header_row, branch_column, field_name = split_field

    # 表头和表头上方的所有行均需保留。
    keep_rows = set(range(1, header_row + 1))
    unmatched: list[dict[str, Any]] = []

    for row_number in range(header_row + 1, worksheet.max_row + 1):
        raw_value = worksheet.cell(row_number, branch_column).value
        normalized_key = branch_key(raw_value)

        if not normalized_key:
            # 空白行通常是模板行、分隔行或供支行补充填写的预留行。
            keep_rows.add(row_number)
            continue

        matched_branch = branch_lookup.get(normalized_key)

        if matched_branch == branch:
            # 当前管辖行的数据需要保留。
            keep_rows.add(row_number)
        elif matched_branch is None:
            # 无法归属的数据保留到每个文件，避免数据被静默丢弃。
            keep_rows.add(row_number)
            unmatched.append({"row": row_number, "value": str(raw_value)})
        # matched_branch 是其他管辖行时，不加入 keep_rows，稍后统一删除。

    original_max_row = worksheet.max_row
    kept_data_rows = sum(
        1
        for row_number in keep_rows
        if row_number > header_row
        and worksheet.cell(row_number, branch_column).value not in (None, "")
    )

    # 记录删除前后的行号关系，用于恢复行高和调整行内公式。
    old_to_new: dict[int, int] = {}
    next_new_row = 1

    for old_row in range(1, original_max_row + 1):
        if old_row in keep_rows:
            old_to_new[old_row] = next_new_row
            next_new_row += 1

    delete_unwanted_rows(
        worksheet,
        first_data_row=header_row + 1,
        last_row=original_max_row,
        keep_rows=keep_rows,
    )

    copy_row_dimensions(worksheet, old_to_new)
    translate_moved_formulas(
        worksheet,
        {new_row: old_row for old_row, new_row in old_to_new.items()},
    )

    return {
        "action": "split",
        "field": field_name,
        "header_row": header_row,
        "kept": kept_data_rows,
        "removed": original_max_row - len(keep_rows),
        "unmatched": unmatched,
    }


def split_workbook(source: Path, branches_path: Path, output_dir: Path) -> dict[str, Any]:
    """
    执行完整拆分任务，并返回运行报告。

    每生成一个管辖行文件，都会重新加载一份完整的原始工作簿。
    这样各个结果文件互不影响，也能最大程度保留原工作簿内容。
    """
    started_at = time.perf_counter()
    check_input_files(source, branches_path)

    branches = read_branch_names(branches_path)
    branch_lookup = build_branch_lookup(branches)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"开始拆分，共 {len(branches)} 个管辖行...", flush=True)
    print("以下工作表保持原样：" + "、".join(sorted(PRESERVE_SHEETS)), flush=True)

    report: dict[str, Any] = {
        "source": str(source),
        "branches_file": str(branches_path),
        "output_dir": str(output_dir),
        "branch_count": len(branches),
        "preserved_sheets": sorted(PRESERVE_SHEETS),
        "files": [],
        "sheets_without_split_field": [],
        "unmatched_values": {},
    }

    for index, branch in enumerate(branches, start=1):
        branch_started_at = time.perf_counter()
        print(f"[{index:02d}/{len(branches)}] 正在处理：{branch}", flush=True)

        # data_only=False 表示读取并保留公式，而不是只读取公式计算结果。
        workbook = load_workbook(source, data_only=False)
        branch_sheet_results: dict[str, Any] = {}

        for worksheet in workbook.worksheets:
            # 明确要求不拆分的工作表，在这里直接跳过所有筛选逻辑。
            if worksheet.title in PRESERVE_SHEETS:
                branch_sheet_results[worksheet.title] = {
                    "action": "preserved",
                    "reason": "工作表位于 PRESERVE_SHEETS 原样保留名单中",
                }
                continue

            result = filter_sheet(worksheet, branch, branch_lookup)
            branch_sheet_results[worksheet.title] = result

            if (
                result["field"] is None
                and worksheet.title not in report["sheets_without_split_field"]
            ):
                report["sheets_without_split_field"].append(worksheet.title)

            if result["unmatched"]:
                report["unmatched_values"].setdefault(
                    worksheet.title,
                    result["unmatched"],
                )

        # 公式版工作簿拆分后，公式引用的行数和范围可能发生变化。
        # 这里要求 Excel 打开文件时做一次完整重算，确保显示结果按拆分后的数据刷新。
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"

        output_path = output_dir / f"{branch}.xlsx"
        workbook.save(output_path)
        workbook.close()

        elapsed = time.perf_counter() - branch_started_at
        print(
            f"[{index:02d}/{len(branches)}] 已完成：{branch}（{elapsed:.1f} 秒）",
            flush=True,
        )

        report["files"].append(
            {
                "branch": branch,
                "path": str(output_path),
                "sheets": branch_sheet_results,
            }
        )

    report_path = output_dir / "拆分运行报告.json"
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"全部完成，总耗时 {time.perf_counter() - started_at:.1f} 秒。", flush=True)
    return report


def main() -> None:
    """程序入口：读取参数、执行拆分，并在终端展示最终结果。"""
    args = parse_args()
    report = split_workbook(
        args.source.resolve(),
        args.branches.resolve(),
        args.output_dir.resolve(),
    )

    print(f"已生成 {len(report['files'])} 个工作簿：{report['output_dir']}")
    print("原样保留、不拆分的工作表：" + "、".join(report["preserved_sheets"]))

    if report["sheets_without_split_field"]:
        print("没有拆分字段、已原样保留的工作表：" + "、".join(report["sheets_without_split_field"]))

    if report["unmatched_values"]:
        print("存在无法匹配管辖行的数据，处理方式详见拆分运行报告.json")


if __name__ == "__main__":
    main()
