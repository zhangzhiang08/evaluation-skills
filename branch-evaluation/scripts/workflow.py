"""GUI 和 Skill 的共同入口。每个拆分任务在独立 Python 进程中执行。"""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from uuid import uuid4

from openpyxl import load_workbook

from workbook_html import write_pages

ROOT = Path(__file__).resolve().parent
ROUNDS = {"first": "第一次拆表", "second": "第二次拆表"}
TEAMS = {"service": "综服经理", "ops": "运营内控经理"}


def split(source, branches, folder, stage, team):
    # 两个历史目录里有同名模块。每次启动时只导入所选轮次的版本。
    sys.path.insert(0, str(ROOT / ROUNDS[stage]))
    if team == "ops":
        import split_q2_ops_internal_control_workbook_by_branch as engine

        engine.PRESERVE_SHEET_KEYWORDS = ("评价方案", "名额分配")
    else:
        import split_q2_service_base_workbook_by_branch as service
        import split_workbook_by_branch as engine

        service.configure_engine()
        book = load_workbook(source, read_only=True)
        try:
            engine.PRESERVE_SHEETS = {
                name for name in book.sheetnames
                if any(word in name for word in ("说明", "校验", "评价方案", "名额分配"))
            }
        finally:
            book.close()
        lookup = engine.build_branch_lookup(engine.read_branch_names(branches))
        service.EMPLOYEE_TO_BRANCH = service.build_employee_to_branch(source, lookup)
    return engine.split_workbook(source, branches, folder), engine.branch_key


def prefix_for(source, stage, team):
    period = re.search(r"20\d{2}年[第]?[一二三四1-4]季度", source.stem)
    if not period:
        return source.stem
    attachment, title = ("附件1", "运营内控队伍") if team == "ops" else ("附件2", "综服经理队伍")
    draft = "【征求意见稿】" if stage == "second" else ""
    kind = "计分表" if stage == "second" else "基础数据"
    return f"{attachment}：{draft}{period.group()}{title}履职评价{kind}"


def main():
    parser = argparse.ArgumentParser(description="按轮次、队伍生成支行 Excel 和本地网页")
    parser.add_argument("--stage", choices=ROUNDS, required=True)
    parser.add_argument("--team", choices=TEAMS, required=True)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--html", action="store_true", help="同时生成工作簿数据网页")
    parser.add_argument("--output", type=Path, required=True, help="保存各批次的文件夹")
    args = parser.parse_args()
    source = args.source.expanduser().resolve()
    if not source.is_file() or source.suffix.lower() != ".xlsx":
        parser.error("请选择有效的 .xlsx 源文件")
    base = args.output.expanduser().resolve()
    base.mkdir(parents=True, exist_ok=True)
    label = ROUNDS[args.stage] + "（" + TEAMS[args.team] + "）"
    batch = base / f"{label}_{datetime.now():%Y%m%d_%H%M%S}_{uuid4().hex[:6]}"
    branches = ROOT / ROUNDS[args.stage] / "管辖行.xlsx"
    # Excel、网页和运行报告全部写完后，一起移入最终批次目录。
    with TemporaryDirectory(dir=base, prefix=".拆表-") as temp:
        folder = Path(temp) / "结果"
        report, short_name = split(source, branches, folder, args.stage, args.team)
        if len(report["files"]) != 18:
            raise ValueError("请核对管辖行.xlsx，预期生成18个支行文件")
        prefix = prefix_for(source, args.stage, args.team)
        for item in report["files"]:
            old = Path(item["path"])
            name = f"{prefix}（{short_name(item['branch'])}）.xlsx"
            if (folder / name).exists() and old.name != name:
                raise ValueError(f"支行简称重复：{name}")
            old.rename(folder / name)
            item["path"] = str(batch / name)
        report["output_dir"] = str(batch)
        (folder / "拆分运行报告.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        if args.html:
            print("正在生成各支行的工作表网页…", flush=True)
            write_pages(folder, report)
        folder.rename(batch)
    unclassified = report.get("sheets_without_split_key") or report.get("sheets_without_split_field")
    if unclassified:
        print("未识别拆分字段，整表保留：" + "、".join(unclassified))
    print(f"结果目录：{batch}", flush=True)
    if args.html:
        print(f"数据网页：{batch / 'index.html'}", flush=True)


if __name__ == "__main__":
    main()
