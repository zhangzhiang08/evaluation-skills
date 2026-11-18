#!/usr/bin/env python3
"""
拆分“2026年二季度综服经理队伍履职评价基础数据表（汇总）-公式版.xlsx”。

============================== 一、脚本目标 ==============================

这个脚本会把二季度综服经理履职评价基础数据汇总表，按照“管辖行.xlsx”
中的 18 个管辖行拆分成 18 个独立 Excel 文件。

脚本会尽量保留原 Excel 的：

1. 表头样式、颜色、边框；
2. 列宽、行高、合并单元格；
3. 单元格公式；
4. 原工作簿中的工作表顺序和工作表名称。

============================== 二、拆分逻辑 ==============================

本文件与之前的一季度“基础数据表”逻辑一致：

1. “管辖行.xlsx”只作为 18 个管辖行名单，不会被拆分；
2. 每个管辖行生成一个 Excel 文件；
3. 以下三张表完整保留，不拆分：
   - 1.评定表格说明
   - 2-1 校验表
   - 2-2 综合服务经理履职评价方案
4. 其他数据表会自动寻找以下字段进行拆分：
   - 管辖支行
   - 管辖行
   - 责任机构
   - 所属机构
5. 如果某行能判断为当前管辖行，就保留；
6. 如果某行明确属于其他管辖行，就删除；
7. 如果某行是“个人金融部”“办公室”“交易银行部”等非 18 个管辖行名单内的机构，
   也统一删除；
8. 如果“管辖支行/管辖行”列是公式，脚本会优先用员工号反查所属管辖行，
   避免公式列无法直接计算导致拆分错误。

============================== 三、二季度名称兼容 ==============================

二季度文件里有一些机构写法和“管辖行.xlsx”不完全一致，例如：

    “中心区支行”   -> “中心区管辖支行”
    “高新区支行”   -> “高新区管辖支行”
    “分行营业部”   -> “营业部”
    “河套皇岗”     -> “河套皇岗分行”

本脚本会先把这些写法转换成统一简称，再和 18 个正式管辖行名称匹配。

============================== 四、运行方法 ==============================

首次运行前安装依赖：

    python3 -m pip install openpyxl

进入脚本所在目录：

    cd "/Users/wenzhou/Documents/履职评价拆表"

按默认路径运行：

    python3 split_q2_service_base_workbook_by_branch.py

指定其他路径运行：

    python3 split_q2_service_base_workbook_by_branch.py \\
      --source "/源文件路径.xlsx" \\
      --branches "/管辖行.xlsx路径" \\
      --output-dir "/输出目录"

默认输出目录：

    当前脚本目录/二季度综服基础数据拆分目录
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

# 复用前面已经写好并验证过的通用拆分引擎。
# 通用引擎负责：读取名单、识别拆分字段、删除其他管辖行行、保留样式、调整公式、生成报告。
import split_workbook_by_branch as engine


DEFAULT_SOURCE = Path("/Users/wenzhou/Downloads/2026年二季度综服经理队伍履职评价基础数据表（汇总）-公式版.xlsx")
DEFAULT_BRANCHES = Path("/Users/wenzhou/Downloads/管辖行.xlsx")
DEFAULT_OUTPUT = Path(__file__).resolve().parent / "二季度综服基础数据拆分目录"

# 与一季度基础数据表保持一致：说明页、校验表、评价方案不拆。
PRESERVE_SHEETS = {
    "1.评定表格说明",
    "2-1 校验表",
    "2-2 综合服务经理履职评价方案",
}

# 全局员工映射表：员工号 -> 管辖行正式名称。
# 二季度文件里有个别行的“管辖支行”列本身是公式，openpyxl 不会计算公式结果；
# 遇到这种情况，就用员工号反查该员工所属管辖行，避免误判。
EMPLOYEE_TO_BRANCH: dict[str, str] = {}


def q2_branch_key(value: Any) -> str:
    """
    把不同写法的管辖行名称统一成简称。

    例如：
        “宝安管辖支行” -> “宝安”
        “宝安支行”     -> “宝安”
        “中心区支行”   -> “中心区”
        “分行营业部”   -> “营业部”
        “河套皇岗分行” -> “河套皇岗”

    返回简称后，通用引擎会和“管辖行.xlsx”的 18 个管辖行做匹配。
    """
    if value is None:
        return ""

    text = re.sub(r"\s+", "", str(value))
    if not text:
        return ""

    # 如果是带机构号或中行前缀的写法，先去掉前缀。
    text = re.sub(r"^\[\d+\]", "", text)
    text = re.sub(r"^中国银行深圳(?:市)?", "", text)

    if "/" in text:
        text = text.split("/", 1)[0]

    if text == "分行营业部":
        return "营业部"

    # 从较长后缀开始删除，避免“管辖支行”只被删成“管辖”。
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
    标准化员工号。

    Excel 中员工号可能是数字，也可能是文本；这里统一转成纯文本。
    """
    if value is None:
        return ""
    text = re.sub(r"\s+", "", str(value))
    if text.endswith(".0"):
        text = text[:-2]
    return text


def build_employee_to_branch(source: Path, branch_lookup: dict[str, str]) -> dict[str, str]:
    """
    建立“员工号 -> 管辖行”的映射。

    为什么需要这个：
    二季度公式版中，部分工作表的“管辖支行”列不是固定文字，而是 Excel 公式。
    Python 的 openpyxl 只保留公式，不会帮我们算出公式结果。
    所以遇到这种公式行时，脚本会改用员工号来判断它属于哪个管辖行。

    主要从两张表学习：
    1. “1.评价归属确认表”：A列员工号，D列管辖支行；
    2. “3-1 柜面及厅堂业务量”：E列员工号，C列管辖行。
    """
    workbook = load_workbook(source, data_only=False)
    mapping: dict[str, str] = {}

    sources = [
        ("1.评价归属确认表", 4, 1, 4),
        ("3-1 柜面及厅堂业务量", 2, 5, 3),
    ]

    for sheet_name, header_row, employee_col, branch_col in sources:
        if sheet_name not in workbook.sheetnames:
            continue
        worksheet = workbook[sheet_name]
        for row_number in range(header_row + 1, worksheet.max_row + 1):
            employee_id = normalize_id(worksheet.cell(row_number, employee_col).value)
            branch = branch_lookup.get(q2_branch_key(worksheet.cell(row_number, branch_col).value))
            if employee_id and branch:
                mapping.setdefault(employee_id, branch)

    workbook.close()
    return mapping


def employee_columns_for_sheet(worksheet, header_row: int) -> list[int]:
    """
    找出某张表中可能表示员工号的列。

    常见表头包括：
    “EHR员工号”“员工号”“柜员号”“总行员工号”等。
    """
    keywords = ("员工号", "EHR员工号", "柜员号", "总行员工号", "总行工号")
    columns: list[int] = []
    for column in range(1, worksheet.max_column + 1):
        value = worksheet.cell(header_row, column).value
        text = re.sub(r"\s+", "", str(value)) if value is not None else ""
        if any(keyword in text for keyword in keywords):
            columns.append(column)
    return columns


def row_branch_by_employee(worksheet, row_number: int, employee_columns: list[int]) -> str | None:
    """
    根据一行中的员工号反查管辖行。

    如果一行没有员工号，或者员工号不在映射表中，则返回 None。
    """
    for column in employee_columns:
        employee_id = normalize_id(worksheet.cell(row_number, column).value)
        if employee_id in EMPLOYEE_TO_BRANCH:
            return EMPLOYEE_TO_BRANCH[employee_id]
    return None


def strict_filter_sheet(worksheet, branch: str, branch_lookup: dict[str, str]) -> dict[str, Any]:
    """
    二季度专用的严格拆分逻辑。

    和通用引擎的区别：
    1. 非当前管辖行的数据全部删除；
    2. “个人金融部”“办公室”“交易银行部”等不在18个管辖行名单内的机构，也删除；
    3. 如果拆分字段是公式，尝试用员工号反查管辖行；
    4. 拆分字段为空但有实际内容的说明行会保留，完全空白行会清除。
    """
    split_field = engine.find_split_field(worksheet)

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
    employee_columns = employee_columns_for_sheet(worksheet, header_row)
    keep_rows = set(range(1, header_row + 1))
    rows_with_values = engine.rows_with_any_data(worksheet)
    deleted_unmatched: list[dict[str, Any]] = []

    for row_number in range(header_row + 1, worksheet.max_row + 1):
        raw_value = worksheet.cell(row_number, branch_column).value
        normalized_key = q2_branch_key(raw_value)
        matched_branch = branch_lookup.get(normalized_key)
        raw_is_formula = isinstance(raw_value, str) and raw_value.startswith("=")

        # 只有当管辖行列是公式或空白时，才用员工号反查。
        # 如果管辖行列明确写着“办公室”“交易银行部”等非管辖行文字，
        # 就不能再用员工号兜底，否则会把这些非所属管辖行机构误保留下来。
        if matched_branch is None and (raw_is_formula or not normalized_key):
            matched_branch = row_branch_by_employee(worksheet, row_number, employee_columns)

        if matched_branch == branch:
            keep_rows.add(row_number)
        elif not normalized_key and matched_branch is None:
            # 管辖行为空但本行有文字或公式时，视为说明/子表头并保留；
            # 整行完全空白则清除，避免拆分结果留下大量空行。
            if row_number in rows_with_values:
                keep_rows.add(row_number)
        else:
            # 严格模式：只要不是当前管辖行，就删除。
            deleted_unmatched.append({"row": row_number, "value": str(raw_value)})

    original_max_row = worksheet.max_row
    kept_data_rows = sum(1 for row in keep_rows if row > header_row)
    old_to_new: dict[int, int] = {}
    next_new_row = 1
    for old_row in range(1, original_max_row + 1):
        if old_row in keep_rows:
            old_to_new[old_row] = next_new_row
            next_new_row += 1

    engine.compact_rows(
        worksheet,
        old_to_new,
        original_max_row,
        fixed_header_rows=header_row,
    )

    return {
        "action": "split_strict",
        "field": field_name,
        "header_row": header_row,
        "employee_columns": employee_columns,
        "kept": kept_data_rows,
        "removed": original_max_row - len(keep_rows),
        "removed_blank_rows": sum(
            1
            for row_number in range(header_row + 1, original_max_row + 1)
            if row_number not in rows_with_values
        ),
        "unmatched": deleted_unmatched,
    }


def parse_args() -> argparse.Namespace:
    """读取命令行参数；没有传入参数时使用默认路径。"""
    parser = argparse.ArgumentParser(description="按管辖行拆分二季度综服经理基础数据表。")
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE, help="待拆分的二季度汇总工作簿")
    parser.add_argument("--branches", type=Path, default=DEFAULT_BRANCHES, help="管辖行名单工作簿")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT, help="拆分结果输出目录")
    return parser.parse_args()


def configure_engine() -> None:
    """
    把二季度文件的专用规则交给通用拆分引擎。

    主要配置两件事：
    1. 哪些工作表不拆；
    2. 如何把机构名称转换成统一简称。
    """
    engine.PRESERVE_SHEETS = PRESERVE_SHEETS
    engine.branch_key = q2_branch_key
    engine.filter_sheet = strict_filter_sheet


def main() -> None:
    """程序入口：配置规则，然后执行拆分。"""
    args = parse_args()
    configure_engine()

    branches = engine.read_branch_names(args.branches.resolve())
    branch_lookup = engine.build_branch_lookup(branches)
    global EMPLOYEE_TO_BRANCH
    EMPLOYEE_TO_BRANCH = build_employee_to_branch(args.source.resolve(), branch_lookup)
    print(f"已建立员工号到管辖行映射 {len(EMPLOYEE_TO_BRANCH)} 条。")

    report = engine.split_workbook(
        source=args.source.resolve(),
        branches_path=args.branches.resolve(),
        output_dir=args.output_dir.resolve(),
    )

    print(f"已生成 {len(report['files'])} 个二季度基础数据工作簿：{report['output_dir']}")
    print("原样保留、不拆分的工作表：" + "、".join(report["preserved_sheets"]))

    if report["sheets_without_split_field"]:
        print("没有拆分字段、已原样保留的工作表：" + "、".join(report["sheets_without_split_field"]))

    if report["unmatched_values"]:
        print("存在无法匹配管辖行的数据，已保留在全部文件中；详见拆分运行报告.json")


if __name__ == "__main__":
    main()
