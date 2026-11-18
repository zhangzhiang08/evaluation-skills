"""从拆好的 Excel 生成离线网页。也可在 Excel 保存分数后单独刷新网页。"""

import argparse
import json
import shutil
import subprocess
from datetime import date, datetime
from html import escape
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import quote

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

STYLE = """body{font:16px system-ui,sans-serif;color:#172c39;max-width:96rem;margin:2rem auto;padding:0 2rem}
a{color:#00675d} h1{font-size:1.7rem} summary{cursor:pointer;padding:.8rem 0;font-weight:bold}
details{border-top:1px solid #cbd6d8}.scroll{overflow:auto;max-height:70vh}
table{border-collapse:collapse;font-size:.9rem}td{padding:.35rem .6rem;border:1px solid #ccd6da;white-space:pre-wrap}
li{margin:.6rem 0}"""


def page(title, body):
    return (f'<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" '
            f'content="width=device-width,initial-scale=1"><title>{escape(title)}</title>'
            f'<style>{STYLE}</style><body><h1>{escape(title)}</h1>{body}</body></html>')


def display(cell, cached):
    value = cell.value
    if cell.data_type == "f":
        value = cached.cell(cell.row, cell.column).value
        if value is None:
            return f'<span title="{escape(str(cell.value), quote=True)}">待重算</span>'
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        value = value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, (int, float)) and "%" in cell.number_format:
        pattern = cell.number_format.split("%")[0]
        digits = len(pattern.partition(".")[2].rstrip("#"))
        value = format(value, f".{min(digits, 12)}%")
    return escape(str(value))


def table(ws, cached):
    # 忽略仅带格式的尾部空行；保留数据之间的表头和合并区域。
    used = [cell for cell in ws._cells.values() if cell.value is not None]
    if not used:
        return "<p>空表</p>"
    last_row = max(c.row for c in used)
    last_col = max(c.column for c in used)
    hidden = set()
    for key, dim in ws.column_dimensions.items():
        if dim.hidden:
            start = dim.min or column_index_from_string(key)
            hidden.update(range(start, (dim.max or start) + 1))
    cols = [c for c in range(1, last_col + 1) if c not in hidden]
    rows = [r for r in range(1, last_row + 1) if not getattr(ws.row_dimensions.get(r), "hidden", False)]
    spans, covered = {}, set()
    for area in ws.merged_cells.ranges:
        rr = [r for r in rows if area.min_row <= r <= area.max_row]
        cc = [c for c in cols if area.min_col <= c <= area.max_col]
        if rr and cc:
            covered.update((r, c) for r in rr for c in cc)
            spans[rr[0], cc[0]] = f' rowspan="{len(rr)}" colspan="{len(cc)}"'
    parts = ['<div class="scroll"><table>']
    for r in rows:
        parts.append("<tr>")
        for c in cols:
            if (r, c) in covered and (r, c) not in spans:
                continue
            parts.append(f'<td{spans.get((r,c), "")}>{display(ws.cell(r,c), cached)}</td>')
        parts.append("</tr>")
    return "".join(parts) + "</table></div>"


def write_pages(folder, report, refresh=False):
    files = [(item, folder / Path(item["path"]).name) for item in report["files"]]
    links = []
    with TemporaryDirectory() as temp:
        cache = Path(temp)
        office = shutil.which("soffice")
        mac_office = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        office = office or (str(mac_office) if mac_office.is_file() else None)
        if office and not refresh:
            # 只对临时副本重算。交付的 Excel 保留拆表引擎写出的公式。
            proc = subprocess.run([office, "-env:UserInstallation=" + (cache / "profile").as_uri(),
                                   "--headless", "--convert-to", "xlsx", "--outdir", str(cache),
                                   *[str(path) for _, path in files]], capture_output=True)
            if proc.returncode:
                print("本机重算未全部完成，网页将标注缺少结果的公式。", flush=True)
        for number, (item, path) in enumerate(files, 1):
            branch = escape(item["branch"])
            link = f"branch-{number:02}.html"
            links.append(f'<li><a href="{link}">{branch}</a></li>')
            wb = load_workbook(path)
            copy = cache / path.name
            values = load_workbook(copy if copy.is_file() else path, data_only=True)
            try:
                parts = ['<p><a href="index.html">所有管辖行</a> · '
                         f'<a href="{quote(path.name)}">打开 Excel</a></p>']
                for ws in wb.worksheets:
                    if ws.sheet_state == "visible":
                        parts.append(f'<details><summary>{escape(ws.title)}</summary>'
                                     + table(ws, values[ws.title]) + '</details>')
                (folder / link).write_text(page(item["branch"], "".join(parts)), encoding="utf-8")
            finally:
                wb.close()
                values.close()
    warnings = report.get("sheets_without_split_key") or report.get("sheets_without_split_field") or []
    note = "<p>未识别归属字段，已整表保留：" + escape("、".join(warnings)) + "</p>" if warnings else ""
    title = Path(report["source"]).stem
    (folder / "index.html").write_text(page(title, note + "<ul>" + "".join(links) + "</ul>"), encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="从已保存的支行 Excel 刷新数据网页")
    parser.add_argument("folder", type=Path, help="含拆分运行报告.json的批次目录")
    folder = parser.parse_args().folder.resolve()
    write_pages(folder, json.loads((folder / "拆分运行报告.json").read_text(encoding="utf-8")), refresh=True)
