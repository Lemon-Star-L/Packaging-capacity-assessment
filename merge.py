# @Author: YiLong Li
# @Date: 2026/10/28 13:42
# @File: merge_results.py
# @Software: PyCharm


"""
输入：当前脚本目录 / output / 月度及年度result文件
输出：当前脚本目录 / output / A_result.xlsx、A_error.xlsx、
      A_logic.xlsx、delete.xlsx

主要功能：
    批量读取月度及年度工时矩阵，统一合并为最终求解矩阵。
    识别并删除重复方程及不符合统计要求的记录。

处理规则：
    按文件名升序、工作表原始行顺序读取数据。
    仅读取第一张“矩阵_去除全零行”工作表。
    以首个通过校验且包含有效数据的文件确定标准列数。
    列数不一致或存在无效单元格的文件整份跳过。
    删除仅一个模块数量非零且数量小于10的记录。
    对其余记录进行整行比例重复识别。
    比例判断包含全部模块数量及最后一列总工时。
    采用十进制精确比例比较，不使用近似容差。
    重复记录保留首次出现的有效记录。

输出规则：
    A_result.xlsx：合并后的最终工时矩阵，不含表头。
    A_error.xlsx：记录文件、表单、行号、异常内容及处理原因。
    A_logic.xlsx：记录单模块小数量逻辑删除记录。
    delete.xlsx：记录比例重复记录，每组包含保留行与删除行，
                 并标注来源文件及原始行号。

使用说明：
    月度文件命名格式：YYYY_M_result.xlsx。
    年度文件命名格式：YYYY_year_result.xlsx。
    将脚本放在对应项目目录下运行。
    模块数量与总工时必须为非负有限数值。
    保留记录的原始数值，不进行比例归一化修改。
    不修改输入文件，异常文件不参与本次汇总。
    Python 3.10+，依赖openpyxl。
"""


from __future__ import annotations

# ================= 常用配置 =================
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
INPUT_SHEET = "矩阵_去除全零行"  # 要求输入文件的第一张表是此表
RESULT_NAME = "A_result.xlsx"
ERROR_NAME = "A_error.xlsx"
DELETE_NAME = "delete.xlsx"
DELETE_SHEET = "删除对照"
LOGIC_NAME = "A_logic.xlsx"
LOGIC_SHEET = "逻辑删除"
SINGLE_MODULE_THRESHOLD = 10  # 严格小于此值；末列工时不参与模块计数
# ============================================

import math
import re
import shutil
import sys
import tempfile
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from functools import reduce
from typing import Any

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Alignment, Font, PatternFill
except ImportError:
    raise SystemExit("缺少依赖，请在项目终端运行：uv add openpyxl")

Row = list[Any]
Sheets = list[tuple[str, Iterable[Sequence[Any]]]]
PATTERN = re.compile(r"(?P<year>[0-9]{4})_(?P<period>[0-9]{1,2}|year)_result", re.IGNORECASE)
ERROR_HEADER = [
    "级别", "文件", "表单", "原始行号", "列号",
    "标准列数", "实际列数", "原始值", "错误原因", "处理方式",
]


@dataclass
class ParsedFile:
    rows: list[list[int | float]]
    signatures: list[tuple[int, ...]]
    source_lines: list[int]
    width: int
    blank_rows: int
    zero_rows: int


@dataclass
class SourceRow:
    filename: str
    line: int
    values: list[int | float]


def deletion_rows(pairs: list[tuple[SourceRow, SourceRow]], width: int) -> Iterable[Row]:
    """每个删除事件输出一对：保留行在上，被删除行在下。"""
    yield ["序号", "备注"] + [f"x{i}" for i in range(1, width)] + ["总工时"]
    for index, (kept, removed) in enumerate(pairs, 1):
        yield [index, f"保留行｜{kept.filename}｜{INPUT_SHEET}｜第{kept.line}行"] + kept.values
        yield [None, f"删除行｜{removed.filename}｜{INPUT_SHEET}｜第{removed.line}行"] + removed.values


def matches_logic(values: Sequence[int | float]) -> bool:
    """仅一个模块有数量，且0 < 数量 < 阈值；最后一列为工时。"""
    nonzero = [value for value in values[:-1] if value != 0]
    return len(nonzero) == 1 and 0 < nonzero[0] < SINGLE_MODULE_THRESHOLD


def logic_rows(removed: list[SourceRow], width: int) -> Iterable[Row]:
    yield ["序号", "备注"] + [f"x{i}" for i in range(1, width)] + ["总工时"]
    for index, source in enumerate(removed, 1):
        module, quantity = next(
            (i, value) for i, value in enumerate(source.values[:-1], 1) if value != 0
        )
        note = (
            f"删除行｜{source.filename}｜{INPUT_SHEET}｜第{source.line}行｜"
            f"仅x{module}非零，数量{quantity}小于{SINGLE_MODULE_THRESHOLD}"
        )
        yield [index, note] + source.values


def issue(
    errors: list[Row], path: Path, reason: str,
    sheet: str = "", line: int | str = "", column: int | str = "",
    expected: int | str = "", actual: int | str = "", value: Any = "",
) -> None:
    errors.append([
        "错误", path.name, sheet, line, column, expected, actual,
        "" if value is None else str(value), reason, "整份文件跳过，不参与本次汇总",
    ])


def numeric(value: Any) -> tuple[int | float, Decimal]:
    # 输入为程序已生成的纯数字矩阵。空格、文本、公式、布尔值均不默认为0。
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError("矩阵单元格必须为数值，不能是空值、文本、公式或布尔值")
    try:
        decimal = Decimal(str(value))
    except (InvalidOperation, ValueError) as error:
        raise ValueError("无法识别的数值") from error
    if not decimal.is_finite() or decimal < 0:
        raise ValueError("矩阵单元格必须为非负有限数值")
    return value, decimal


def proportion_key(values: list[Decimal]) -> tuple[int, ...]:
    """消去十进制分母，再除以整行最大公约数，得到精确比例指纹。

    不使用浮点除法，不将任一非零列省略；全零行由调用方排除。
    """
    ratios = [value.as_integer_ratio() for value in values]
    denominator = math.lcm(*(den for _, den in ratios))
    integers = [num * (denominator // den) for num, den in ratios]
    divisor = reduce(math.gcd, integers, 0)
    if divisor == 0:
        raise ValueError("全零行不应参与比例计算")
    return tuple(value // divisor for value in integers)


def discover(errors: list[Row]) -> list[Path]:
    paths: list[Path] = []
    for path in sorted(OUTPUT_DIR.iterdir(), key=lambda item: item.name):
        if not path.is_file() or path.name.startswith("~$") or path.suffix.lower() != ".xlsx":
            continue
        if path.name.casefold() in {RESULT_NAME.casefold(), ERROR_NAME.casefold(), DELETE_NAME.casefold(), LOGIC_NAME.casefold()}:
            continue
        match = PATTERN.fullmatch(path.stem)
        if match is None:
            continue
        try:
            year = int(match.group("year"))
            period = match.group("period").lower()
            date(year, 1 if period == "year" else int(period), 1)
        except ValueError as error:
            issue(errors, path, f"结果文件名中的年月无效：{error}")
            continue
        paths.append(path)
    return paths


def read_file(path: Path, expected_width: int | None, errors: list[Row]) -> ParsedFile | None:
    start_errors = len(errors)
    rows: list[list[int | float]] = []
    signatures: list[tuple[int, ...]] = []
    source_lines: list[int] = []
    width: int | None = None
    blank_rows = 0
    zero_rows = 0
    try:
        # data_only=False：公式应作为异常报告，不能将无缓存公式当空白跳过。
        workbook = load_workbook(path, read_only=True, data_only=False)
        try:
            worksheet = workbook.worksheets[0]
            sheet = worksheet.title
            if sheet != INPUT_SHEET:
                issue(errors, path, f"第一张表应为“{INPUT_SHEET}”，实际为“{sheet}”", sheet)
                return None
            # write_only生成的结果表可能没有dimension；按实际单元格读取。
            worksheet.reset_dimensions()
            for line, raw in enumerate(worksheet.iter_rows(values_only=True), 1):
                if all(value is None for value in raw):
                    blank_rows += 1
                    continue
                actual = len(raw)
                if width is None:
                    width = actual
                target = expected_width if expected_width is not None else width
                if actual != target:
                    issue(errors, path, "列数不一致", sheet, line,
                          expected=target, actual=actual)
                    continue
                if actual < 2:
                    issue(errors, path, "至少需要一列模块数量和一列工时", sheet, line,
                          expected="至少2", actual=actual)
                    continue
                values: list[int | float] = []
                decimals: list[Decimal] = []
                valid = True
                for column, value in enumerate(raw, 1):
                    try:
                        original, decimal = numeric(value)
                    except ValueError as error:
                        issue(errors, path, str(error), sheet, line, column,
                              expected=target, actual=actual, value=value)
                        valid = False
                    else:
                        values.append(original)
                        decimals.append(decimal)
                if not valid:
                    continue
                if all(value == 0 for value in decimals):
                    zero_rows += 1
                    continue
                rows.append(values)
                source_lines.append(line)
                signatures.append(proportion_key(decimals))
        finally:
            workbook.close()
    except Exception as error:
        issue(errors, path, f"读取或校验文件失败：{error}")
    if len(errors) > start_errors:
        return None
    return ParsedFile(rows, signatures, source_lines, width or 0, blank_rows, zero_rows)


def write_workbook(path: Path, sheets: Sheets) -> None:
    workbook = Workbook(write_only=True)
    try:
        for title, rows in sheets:
            worksheet = workbook.create_sheet(title)
            is_delete = title == DELETE_SHEET
            is_report = is_delete or title == LOGIC_SHEET
            if is_report:
                worksheet.freeze_panes = "C2"
                worksheet.column_dimensions["A"].width = 9
                worksheet.column_dimensions["B"].width = 75
                worksheet.sheet_format.defaultColWidth = 12
            for line, row in enumerate(rows, 1):
                if line > 1048576 or len(row) > 16384:
                    raise ValueError(f"表单{title}超过Excel行列上限，未发布本次结果")
                cells = []
                for column, value in enumerate(row, 1):
                    cell = WriteOnlyCell(worksheet, value=value)
                    if isinstance(value, str):
                        cell.data_type = "s"
                    if is_report:
                        if line == 1:
                            cell.font = Font(name="微软雅黑", bold=True, color="FFFFFF")
                            cell.fill = PatternFill("solid", fgColor="305496")
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        elif column <= 2:
                            cell.font = Font(name="微软雅黑", size=10)
                            cell.alignment = Alignment(
                                horizontal="center" if column == 1 else "left",
                                vertical="center", wrap_text=True,
                            )
                            if column == 2:
                                cell.fill = PatternFill(
                                    "solid", fgColor="E2F0D9" if is_delete and line % 2 == 0 else "FCE4D6"
                                )
                    cells.append(cell)
                if is_report:
                    worksheet.row_dimensions[line].height = 28 if line == 1 else 34
                    # 第1行为表头；每组序号合并A2:A3、A4:A5……
                    # write_only使用合并区域集合，在保存时写入合并信息。
                    if is_delete and line >= 2 and line % 2 == 0:
                        worksheet.merged_cells.add(f"A{line}:A{line + 1}")
                worksheet.append(cells)
        workbook.save(path)
    finally:
        workbook.close()


def publish(files: Mapping[Path, Sheets | None]) -> None:
    """准备成功再替换；失败回滚。None表示删除对应旧文件。

    不保证断电或强制退出时的跨文件原子性。
    """
    stage = Path(tempfile.mkdtemp(prefix=".merge_", dir=OUTPUT_DIR))
    staged: dict[Path, Path] = {}
    backups: dict[Path, Path] = {}
    installed: list[Path] = []
    preserve = False
    try:
        for index, (target, sheets) in enumerate(files.items()):
            if sheets is not None:
                temp = stage / f"new_{index}.xlsx"
                write_workbook(temp, sheets)
                staged[target] = temp
        try:
            for index, target in enumerate(files):
                if target.exists():
                    backup = stage / f"backup_{index}.xlsx"
                    target.replace(backup)
                    backups[target] = backup
            for target, temp in staged.items():
                temp.replace(target)
                installed.append(target)
        except Exception as publish_error:
            rollback_errors: list[str] = []
            for target in reversed(installed):
                try:
                    target.unlink(missing_ok=True)
                except OSError as error:
                    rollback_errors.append(f"移除{target.name}失败：{error}")
            for target, backup in backups.items():
                try:
                    backup.replace(target)
                except OSError as error:
                    rollback_errors.append(f"恢复{target.name}失败：{error}")
            if rollback_errors:
                preserve = True
                raise RuntimeError(f"发布失败：{publish_error}；{'；'.join(rollback_errors)}；恢复文件位于{stage}") from publish_error
            raise RuntimeError(f"发布失败，已恢复原文件：{publish_error}；请检查文件是否被Excel占用") from publish_error
    finally:
        if not preserve:
            try:
                shutil.rmtree(stage)
            except OSError:
                print(f"[提示] 临时目录未清理：{stage}")


def main() -> int:
    if not OUTPUT_DIR.is_dir():
        raise FileNotFoundError(f"输出目录不存在：{OUTPUT_DIR}")
    errors: list[Row] = []
    paths = discover(errors)
    print("=" * 64)
    print(f"读取目录：{OUTPUT_DIR}")
    print("范围：YYYY_M_result.xlsx、YYYY_year_result.xlsx；只读第一张去除全零行的矩阵")
    print(f"先行删除：仅一个模块非零且数量小于{SINGLE_MODULE_THRESHOLD}；末列工时不参与模块计数")
    print("规则：全部列含工时参与任意正倍数判断；精确比较；保留最先出现的一行")
    print("注意：列数校验不能验证模块含义；输入文件必须使用相同的模块列顺序")
    if not paths and not errors:
        raise ValueError("未找到符合命名规则的结果文件；未更新A_result.xlsx，旧文件不能视为本次结果")

    expected_width: int | None = None
    seen: dict[tuple[int, ...], SourceRow] = {}
    deleted_pairs: list[tuple[SourceRow, SourceRow]] = []
    logic_removed: list[SourceRow] = []
    merged: list[list[int | float]] = []
    accepted = 0
    empty_files = 0
    skipped = len({row[1] for row in errors})
    before_rows = 0
    duplicate_rows = 0
    zero_rows = 0
    blank_rows = 0
    for path in paths:
        parsed = read_file(path, expected_width, errors)
        if parsed is None:
            skipped += 1
            print(f"[跳过] {path.name}：校验失败，详见{ERROR_NAME}")
            continue
        zero_rows += parsed.zero_rows
        blank_rows += parsed.blank_rows
        if not parsed.rows:
            empty_files += 1
            print(f"[无有效行] {path.name}：空行{parsed.blank_rows}，全零行{parsed.zero_rows}")
            continue
        if expected_width is None:
            expected_width = parsed.width
            print(f"[列数基准] {path.name}：{expected_width}列（含末列工时）")
        accepted += 1
        before_rows += len(parsed.rows)
        added = 0
        removed = 0
        logic_count = 0
        for row, signature, line in zip(parsed.rows, parsed.signatures, parsed.source_lines):
            current = SourceRow(path.name, line, row)
            if matches_logic(row):
                logic_removed.append(current)
                logic_count += 1
                continue
            if signature in seen:
                deleted_pairs.append((seen[signature], current))
                removed += 1
            else:
                seen[signature] = current
                merged.append(row)
                added += 1
        duplicate_rows += removed
        print(f"[汇总] {path.name}：有效{len(parsed.rows)}行；逻辑删除{logic_count}行；删除相同/成比例{removed}行；新增{added}行")

    error_sheets: Sheets | None = [("错误记录", [ERROR_HEADER] + errors)] if errors else None
    result_path = OUTPUT_DIR / RESULT_NAME
    error_path = OUTPUT_DIR / ERROR_NAME
    delete_path = OUTPUT_DIR / DELETE_NAME
    logic_path = OUTPUT_DIR / LOGIC_NAME
    if expected_width is None:
        publish({error_path: error_sheets})
        print("[未生成] 没有可汇总的有效行，未更新A_result.xlsx、delete.xlsx和A_logic.xlsx；旧结果不能视为本次结果")
        if errors:
            print(f"[错误报告] {error_path}")
        return 1

    publish({
        result_path: [("汇总矩阵", merged)],
        error_path: error_sheets,
        delete_path: [(DELETE_SHEET, deletion_rows(deleted_pairs, expected_width))] if deleted_pairs else None,
        logic_path: [(LOGIC_SHEET, logic_rows(logic_removed, expected_width))] if logic_removed else None,
    })
    print("=" * 64)
    print(f"文件：参与{accepted}个；无有效行{empty_files}个；错误跳过{skipped}个")
    print(f"行数：合并前{before_rows}行 - 逻辑删除{len(logic_removed)}行 - 相同/成比例{duplicate_rows}行 = 保留{len(merged)}行")
    print(f"另外跳过：空行{blank_rows}行；全零行{zero_rows}行（仅统计通过校验的文件）")
    print(f"[结果] {result_path}；无表头、无日期列，保持原列顺序")
    if deleted_pairs:
        print(f"[删除对照] {delete_path}；{len(deleted_pairs)}组，每组2行（保留行/删除行）")
    else:
        print(f"[删除对照] 无记录，不生成{DELETE_NAME}，已清理同名旧报告（如有）")
    if logic_removed:
        print(f"[逻辑删除] {logic_path}；{len(logic_removed)}行")
    else:
        print(f"[逻辑删除] 无记录，不生成{LOGIC_NAME}，已清理同名旧报告（如有）")
    if not merged:
        print("[提示] 有效行全部被逻辑规则删除；已用空矩阵替换旧A_result.xlsx")
    if errors:
        print(f"[错误报告] {error_path}；共{len(errors)}项问题；当前结果仅含通过校验的文件")
    else:
        print("[检查] 无错误；不生成A_error.xlsx，已清理同名旧错误报告（如有）")
    # 部分文件跳过时仍保存有效结果，但返回1提示需要查看报告。
    return 1 if errors else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as startup_error:
        print(f"[运行失败] {startup_error}")
        sys.exit(1)
