# @Author: YiLong Li
# @Date: 2026/10/23 16:28
# @File: pre_execute.py
# @Software: PyCharm


"""
主要功能：
    读取业务数据、人员表及 Resource.xlsx，
    根据“件号—备件模块—自变量”的对应关系，
    按日期汇总各模块业务数量，并计算当天总工时。

处理规则：
    1. 支持月度和年度文件批量处理，均按日汇总。
    2. 当天总工时 = 当天人数 × 每人工时系数。
    3. 当天全部模块数量为0时，总工时也置0。
    4. 检查输入数据，导出异常记录，并核对数量平衡。

输入位置：
    业务表和人员表：脚本所在目录下的 data/execute 文件夹。
    Resource.xlsx：脚本所在目录下的 data 文件夹。

输出内容：
    结果保存至 output 文件夹。
    包含“矩阵_去除全零行”和“矩阵_保留全零行”两个表单，
    并根据处理情况生成错误明细和预检查报告。

使用说明：
    矩阵不含表头和日期列，最后一列为当天总工时。
    本脚本仅生成计算所需的数据，不直接求解模块单件工时。
"""


# ================== 常用配置 ===================

PROCESS_MODE = 2        # 1：月度模式；2：年度表模式（对应截图2026_year.xlsx）
FIELD_MODE = 1          # 1：中文业务字段；2：原英文字段（与PROCESS_MODE独立）
HOURS_PER_PERSON = 11   # 当天总工时 = 当天人数 × 此系数

# ==============================================

from pathlib import Path
from datetime import datetime, date
from collections import defaultdict, Counter
import math
import re
import shutil
import sys
import tempfile
try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.cell import WriteOnlyCell
except ImportError:
    raise SystemExit("缺少依赖，请在项目终端运行：uv add openpyxl")


# ================= 其他配置 =================

BASE_DIR = Path(__file__).resolve().parent
DATA_ROOT = BASE_DIR / "data"
DATA_SUBDIR = "execute" # data下的业务数据子文件夹；填""则直接读取data
DATA_DIR = DATA_ROOT / DATA_SUBDIR
OUTPUT_DIR = BASE_DIR / "output"
RESOURCE_FILE = DATA_ROOT / "Resource.xlsx"

# 留空：按照当前模式自动扫描。
# 模式1示例：["2026_8.xlsx", "2026_9.xlsx"]
# 模式2示例：["2025_year.xlsx", "2026_year.xlsx"]
DATA_FILES = []

# 0表示第一个表单，也可以填写具体表单名称。
DATA_SHEET = 0
PERSONNEL_SHEET = 0

PARTS_SHEET = "备件信息"
MODULES_SHEET = "模块信息"

# 按字段名称读取，不依赖列的位置。
DATA_FIELD_PROFILES = {
    1: {
        "part": "件号",
        "quantity": "预期收货数量",
        "date": "收货日期时间",
    },
    2: {
        "part": "SUSR1",
        "quantity": "QTYRECEIVED_1",
        "date": "DATERECEIVED",
    },
}


def get_data_fields() -> dict[str, str]:
    """仅切换业务表字段；人员表及Resource字段保持原配置。"""
    if type(FIELD_MODE) is not int or FIELD_MODE not in (1, 2):
        raise ValueError("FIELD_MODE只能设置为整数1（中文）或2（英文）")
    return DATA_FIELD_PROFILES[FIELD_MODE].copy()

PERSONNEL_FIELDS = {
    "date": "日期",
    "people": "人数",
}

PARTS_FIELDS = {
    "part": "件号",
    "module": "备件模块",
}

MODULES_FIELDS = {
    "module": "备件模块",
    "variable": "自变量",
}

REPORT_HEADER = [
    "级别",
    "文件",
    "表单",
    "原始行号",
    "日期",
    "件号",
    "字段",
    "原始值",
    "问题说明",
    "处理方式",
]

# ============================================


class ReportedFailure(Exception):
    """问题已经写入预检查报告。"""


def key(value):
    """去除首尾空格，保留件号内部字符。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def header_key(value):
    """表头忽略首尾空格及英文字母大小写。"""
    return key(value).casefold()


def as_date(value):
    """提取自然日，忽略时、分、秒。"""
    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = key(value).replace("/", "-")

    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass

    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        raise ValueError(f"无法识别日期：{value!r}")


def number(value):
    """读取非负有限数值。"""
    if value is None or isinstance(value, bool) or key(value) == "":
        raise ValueError("数值为空或无效")

    try:
        result = float(value)
    except (ValueError, TypeError, OverflowError):
        raise ValueError(f"不是有效数字：{value!r}")

    if not math.isfinite(result) or result < 0:
        raise ValueError(f"必须为非负有限数字：{value!r}")

    return result



def quantity_number(value):
    """数量支持原生数值、普通数值文本及中英文逗号千位分隔符。

    示例：5,640 / 5，640 / 1,234,567 / 1,234.50。
    逗号前后允许空白；按三位分组校验，避免将5,64误读为564。
    解析仅用于计算，错误明细保留原始单元格值。
    """
    if isinstance(value, str):
        text = value.strip().replace("，", ",")
        if "," in text:
            text = re.sub(r"\s*,\s*", ",", text)
            if not re.fullmatch(r"[+-]?[0-9]{1,3}(?:,[0-9]{3})+(?:\.[0-9]+)?", text):
                raise ValueError(
                    f"数量的千位分隔格式不正确：{value!r}；示例：5,640 或 1,234.50"
                )
            text = text.replace(",", "")
        return number(text)
    return number(value)


def validate_config():
    get_data_fields()  # 提前校验字段配置，防止用错表头。

    if (
        isinstance(PROCESS_MODE, bool)
        or not isinstance(PROCESS_MODE, int)
        or PROCESS_MODE not in (1, 2)
    ):
        raise ValueError("PROCESS_MODE只能设置为整数1或2")

    if (
        isinstance(HOURS_PER_PERSON, bool)
        or not isinstance(HOURS_PER_PERSON, (int, float))
    ):
        raise ValueError("HOURS_PER_PERSON必须是大于0的有限数值")

    try:
        valid = (
            math.isfinite(HOURS_PER_PERSON)
            and HOURS_PER_PERSON > 0
        )
    except (TypeError, OverflowError):
        valid = False

    if not valid:
        raise ValueError("HOURS_PER_PERSON必须是大于0的有限数值")

    if not DATA_DIR.is_dir():
        raise FileNotFoundError(f"数据文件夹不存在：{DATA_DIR}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def read_rows(path, sheet, fields):
    """
    第一行为表头，按字段名称定位。
    保留全部原始列和原始行号。
    不排序、不去重业务数据。
    """
    if not path.is_file():
        raise FileNotFoundError(f"找不到文件：{path}")

    workbook = load_workbook(
        path,
        read_only=True,
        data_only=True,
    )

    try:
        try:
            worksheet = (
                workbook.worksheets[sheet]
                if isinstance(sheet, int)
                else workbook[sheet]
            )
        except (IndexError, KeyError) as exc:
            raise ValueError(
                f"{path.name}：找不到指定表单 {sheet!r}"
            ) from exc

        iterator = worksheet.iter_rows(values_only=True)
        header = next(iterator, None)

        if header is None:
            raise ValueError(
                f"{path.name}/{worksheet.title}：没有表头"
            )

        positions = defaultdict(list)

        for index, value in enumerate(header):
            name = header_key(value)
            if name:
                positions[name].append(index)

        duplicate_headers = [
            name
            for name, indexes in positions.items()
            if len(indexes) > 1
        ]

        if duplicate_headers:
            raise ValueError(
                "存在重复字段名：" + "、".join(duplicate_headers)
            )

        columns = {}

        for logical_name, expected_name in fields.items():
            normalized = header_key(expected_name)

            if not normalized or normalized not in positions:
                raise ValueError(f"缺少字段“{expected_name}”")

            columns[logical_name] = positions[normalized][0]

        if len(set(columns.values())) != len(columns):
            raise ValueError("配置中的必需字段重复")

        records = [
            (row_number, row)
            for row_number, row in enumerate(iterator, start=2)
            if any(value is not None for value in row)
        ]

        return list(header), records, columns, worksheet.title

    finally:
        workbook.close()


def add_issue(
    issues,
    level,
    path,
    sheet="",
    row_number="",
    day="",
    part="",
    field="",
    value="",
    message="",
    action="",
):
    issues.append([
        level,
        path.name,
        sheet,
        row_number,
        key(day),
        key(part),
        field,
        key(value),
        message,
        action,
    ])


def checked_read(path, sheet, fields, issues):
    """文件读取错误写入预检查报告。"""
    try:
        return read_rows(path, sheet, fields)
    except Exception as exc:
        add_issue(
            issues,
            "错误",
            path,
            sheet=str(sheet),
            message=str(exc),
            action="修正后重新运行；本次不生成新矩阵",
        )
        return None


def checked_value(
    parser,
    value,
    issues,
    path,
    sheet,
    row_number,
    field,
    day="",
    part="",
):
    """记录单元格错误，继续检查其他数据。"""
    try:
        return parser(value)
    except (ValueError, TypeError, OverflowError) as exc:
        add_issue(
            issues,
            "错误",
            path,
            sheet,
            row_number,
            day,
            part,
            field,
            value,
            str(exc),
            "修正后重新运行；本次不生成新矩阵",
        )
        return None


def filename_pattern():
    if PROCESS_MODE == 1:
        return r"(\d{4})_(\d{1,2})"
    return r"(\d{4})_year"


def period_info(path):
    """从文件名识别所属年月或年份。"""
    match = re.fullmatch(filename_pattern(), path.stem)

    if not match or path.suffix.lower() != ".xlsx":
        example = (
            "2026_9.xlsx"
            if PROCESS_MODE == 1
            else "2026_year.xlsx"
        )
        raise ValueError(
            f"文件名不符合当前模式，例如：{example}"
        )

    year = int(match.group(1))

    if PROCESS_MODE == 1:
        month = int(match.group(2))
        date(year, month, 1)
        return year, month

    date(year, 1, 1)
    return (year,)


def check_period(day, period):
    """只校验所属期间，不要求记录按日期排列。"""
    if PROCESS_MODE == 1:
        valid = (day.year, day.month) == period
        expected = f"{period[0]}年{period[1]}月"
    else:
        valid = day.year == period[0]
        expected = f"{period[0]}年"

    if not valid:
        raise ValueError(f"日期{day}不属于{expected}")


def checked_day(
    value,
    period,
    issues,
    path,
    sheet,
    row_number,
    field,
):
    day = checked_value(
        as_date,
        value,
        issues,
        path,
        sheet,
        row_number,
        field,
    )

    if day is None:
        return None

    try:
        check_period(day, period)
    except ValueError as exc:
        add_issue(
            issues,
            "错误",
            path,
            sheet,
            row_number,
            day=day,
            field=field,
            value=value,
            message=str(exc),
            action="修正后重新运行；本次不生成新矩阵",
        )

    return day


def write_workbook(path, sheets):
    """
    sheets格式：
    [
        ("表单名称", 可迭代的行数据),
        ...
    ]
    """
    workbook = Workbook(write_only=True)

    try:
        for title, rows in sheets:
            worksheet = workbook.create_sheet(title)

            for row in rows:
                cells = []

                for value in row:
                    cell = WriteOnlyCell(worksheet, value=value)

                    # 字符串按文本保存，避免被识别为公式。
                    if isinstance(value, str):
                        cell.data_type = "s"

                    cells.append(cell)

                worksheet.append(cells)

        workbook.save(path)

    finally:
        workbook.close()


def publish_files(files):
    """
    files：
        目标路径 -> 表单列表：生成或替换文件。
        目标路径 -> None：删除对应旧文件。

    先准备所有新文件，再统一替换。
    可捕获的发布失败会尝试回滚。
    不保证断电或强制结束进程时的跨文件原子性。
    """
    stage = Path(
        tempfile.mkdtemp(prefix=".matrix_", dir=OUTPUT_DIR)
    )

    staged = {}
    backups = {}
    installed = []
    preserve_stage = False

    try:
        for index, (target, sheets) in enumerate(files.items()):
            if sheets is not None:
                temporary = stage / f"new_{index}.xlsx"
                write_workbook(temporary, sheets)
                staged[target] = temporary

        try:
            for index, target in enumerate(files):
                if target.exists():
                    backup = stage / f"backup_{index}.xlsx"
                    target.replace(backup)
                    backups[target] = backup

            for target, temporary in staged.items():
                temporary.replace(target)
                installed.append(target)

        except Exception as publish_error:
            rollback_errors = []

            for target in reversed(installed):
                try:
                    target.unlink(missing_ok=True)
                except OSError as exc:
                    rollback_errors.append(
                        f"移除新文件{target.name}失败：{exc}"
                    )

            for target, backup in backups.items():
                try:
                    backup.replace(target)
                except OSError as exc:
                    rollback_errors.append(
                        f"恢复{target.name}失败：{exc}"
                    )

            if rollback_errors:
                preserve_stage = True
                raise RuntimeError(
                    f"发布失败：{publish_error}；"
                    f"回滚未全部完成：{'；'.join(rollback_errors)}；"
                    f"恢复文件保留在：{stage}"
                ) from publish_error

            raise RuntimeError(
                f"发布失败，已恢复发布前状态：{publish_error}。"
                "请检查输出文件是否被Excel占用。"
            ) from publish_error

    finally:
        if not preserve_stage:
            try:
                shutil.rmtree(stage)
            except OSError:
                print(f"[提示] 临时目录未能清理：{stage}")


def report_sheets(issues):
    return [("预检查报告", [REPORT_HEADER] + issues)]


def stop_if_errors(issues, report_path):
    error_count = sum(row[0] == "错误" for row in issues)

    if error_count:
        publish_files({
            report_path: report_sheets(issues),
        })

        print(
            f"[预检查] 错误{error_count}项；"
            f"提示{len(issues) - error_count}项"
        )
        print(f"[报告] {report_path}")

        raise ReportedFailure(
            f"预检查发现{error_count}项错误，未生成本次新结果"
        )


def mappings():
    """所有业务文件共用同一份Resource映射。"""
    issues = []
    report_path = OUTPUT_DIR / "Resource_precheck.xlsx"

    parts_data = checked_read(
        RESOURCE_FILE,
        PARTS_SHEET,
        PARTS_FIELDS,
        issues,
    )

    modules_data = checked_read(
        RESOURCE_FILE,
        MODULES_SHEET,
        MODULES_FIELDS,
        issues,
    )

    part_map = defaultdict(set)
    module_columns = defaultdict(list)
    module_count = 0

    if parts_data is not None:
        _, parts, columns, _ = parts_data

        for _, row in parts:
            part = key(row[columns["part"]])
            module = key(row[columns["module"]])
            part_map[part].add(module)

    if modules_data is not None:
        _, modules, columns, sheet = modules_data
        module_count = len(modules)

        if not modules:
            add_issue(
                issues,
                "错误",
                RESOURCE_FILE,
                sheet,
                message="没有模块数据",
                action="补充模块信息后重新运行",
            )

        for column_index, (row_number, row) in enumerate(modules):
            module = key(row[columns["module"]])
            variable = key(row[columns["variable"]])
            expected = f"x{column_index + 1}"

            if not module:
                add_issue(
                    issues,
                    "错误",
                    RESOURCE_FILE,
                    sheet,
                    row_number,
                    field=MODULES_FIELDS["module"],
                    message="备件模块为空",
                    action="补充后重新运行",
                )

            if variable != expected:
                add_issue(
                    issues,
                    "错误",
                    RESOURCE_FILE,
                    sheet,
                    row_number,
                    field=MODULES_FIELDS["variable"],
                    value=variable,
                    message=(
                        f"自变量应为{expected}，实际为{variable!r}"
                    ),
                    action="复核编号和行顺序；程序不会自动重排",
                )

            module_columns[module].append(column_index)

    # #N/A等仍按原来的模块匹配规则处理。
    # 重复模块仍在业务匹配时报告。
    stop_if_errors(issues, report_path)

    if report_path.exists():
        publish_files({report_path: None})

    return part_map, module_columns, module_count


def resolve_module(part, part_map, module_columns):
    """返回矩阵列位置及错误原因。"""
    choices = part_map.get(part, set())

    if not part or not choices:
        return None, "件号无法匹配"

    if len(choices) > 1:
        return None, "件号对应多个不同模块"

    module = next(iter(choices))
    indexes = module_columns.get(module, [])

    if not module:
        return None, "件号对应的备件模块为空"

    if not indexes:
        return None, f"备件模块无法匹配自变量：{module}"

    if len(indexes) > 1:
        return None, f"备件模块对应多个自变量：{module}"

    return indexes[0], ""


def finite_sum(values):
    try:
        total = math.fsum(values)
    except (OverflowError, ValueError) as exc:
        raise ValueError("数量汇总溢出或无效") from exc

    if not math.isfinite(total):
        raise ValueError("数量汇总不是有限数字")

    return total


def process(path, part_map, module_columns, module_count):
    data_fields = get_data_fields()
    period = period_info(path)
    personnel_path = path.with_name(path.stem + "_P.xlsx")

    result_path = OUTPUT_DIR / f"{path.stem}_result.xlsx"
    error_suffix = "erro" if PROCESS_MODE == 1 else "error"
    error_path = OUTPUT_DIR / f"{path.stem}_{error_suffix}.xlsx"
    report_path = OUTPUT_DIR / f"{path.stem}_precheck.xlsx"

    issues = []

    business_data = checked_read(
        path,
        DATA_SHEET,
        data_fields,
        issues,
    )

    personnel_data = checked_read(
        personnel_path,
        PERSONNEL_SHEET,
        PERSONNEL_FIELDS,
        issues,
    )

    people = {}
    people_lines = {}
    hours_by_day = {}
    parsed_records = []

    # ---------- 1. 人员数据预检查 ----------

    if personnel_data is not None:
        _, people_rows, columns, people_sheet = personnel_data

        for row_number, row in people_rows:
            day = checked_day(
                row[columns["date"]],
                period,
                issues,
                personnel_path,
                people_sheet,
                row_number,
                PERSONNEL_FIELDS["date"],
            )

            raw_people = row[columns["people"]]

            if key(raw_people) == "":
                count = None
            else:
                count = checked_value(
                    number,
                    raw_people,
                    issues,
                    personnel_path,
                    people_sheet,
                    row_number,
                    PERSONNEL_FIELDS["people"],
                    day=day,
                )

            if day is None:
                continue

            if day in people:
                add_issue(
                    issues,
                    "错误",
                    personnel_path,
                    people_sheet,
                    row_number,
                    day=day,
                    field=PERSONNEL_FIELDS["date"],
                    value=row[columns["date"]],
                    message=(
                        f"日期重复，首次出现在第{people_lines[day]}行"
                    ),
                    action="合并或修正重复日期后重新运行",
                )
                continue

            people[day] = count
            people_lines[day] = row_number

            if count is not None:
                try:
                    hours = count * HOURS_PER_PERSON

                    if not math.isfinite(hours):
                        raise ValueError("总工时不是有限数字")

                    hours_by_day[day] = hours

                except (ValueError, OverflowError) as exc:
                    add_issue(
                        issues,
                        "错误",
                        personnel_path,
                        people_sheet,
                        row_number,
                        day=day,
                        field=PERSONNEL_FIELDS["people"],
                        value=raw_people,
                        message=f"人数乘工时系数失败：{exc}",
                        action="修正人数或工时系数后重新运行",
                    )

    # ---------- 2. 业务数据预检查 ----------

    if business_data is not None:
        header, records, columns, business_sheet = business_data

        for row_number, row in records:
            part = key(row[columns["part"]])

            day = checked_day(
                row[columns["date"]],
                period,
                issues,
                path,
                business_sheet,
                row_number,
                data_fields["date"],
            )

            quantity = checked_value(
                quantity_number,
                row[columns["quantity"]],
                issues,
                path,
                business_sheet,
                row_number,
                data_fields["quantity"],
                day=day,
                part=part,
            )

            column_index, mapping_error = resolve_module(
                part,
                part_map,
                module_columns,
            )

            if mapping_error:
                add_issue(
                    issues,
                    "提示",
                    path,
                    business_sheet,
                    row_number,
                    day=day,
                    part=part,
                    field=data_fields["part"],
                    value=part,
                    message=mapping_error,
                    action=(
                        "该条业务进入错误文件；"
                        "当天模块数量全为0时整行置0；"
                        "部分匹配成功且人数有效时保留全天工时"
                    ),
                )

            if day is not None and quantity is not None:
                parsed_records.append((
                    row_number,
                    row,
                    day,
                    quantity,
                    part,
                    column_index,
                    mapping_error,
                ))

    # 两份表尽量检查完后，一次性报告严重错误。
    # 严重错误存在时，不生成本次新矩阵。
    stop_if_errors(issues, report_path)

    # ---------- 3. 日期及数量汇总 ----------

    dates = sorted(people)

    date_to_row = {
        day: index
        for index, day in enumerate(dates)
    }

    matrix = [
        [0] * (module_count + 1)
        for _ in dates
    ]

    raw_quantities = defaultdict(list)

    for _, _, day, quantity, _, _, _ in parsed_records:
        raw_quantities[day].append(quantity)

    active_days = set(raw_quantities)
    source_daily = {}

    for day, quantities in raw_quantities.items():
        try:
            source_daily[day] = finite_sum(quantities)
        except ValueError as exc:
            add_issue(
                issues,
                "错误",
                path,
                business_sheet,
                day=day,
                message=str(exc),
                action="核查当天数量；本次不生成新矩阵",
            )

    stop_if_errors(issues, report_path)

    missing_date_days = active_days - set(people)

    empty_people_days = {
        day
        for day in active_days & set(people)
        if people[day] is None
    }

    zero_people_days = {
        day
        for day in active_days & set(people)
        if people[day] == 0 and source_daily[day] > 0
    }

    no_business_days = set(people) - active_days

    for day in sorted(missing_date_days):
        add_issue(
            issues,
            "提示",
            personnel_path,
            people_sheet,
            day=day,
            field=PERSONNEL_FIELDS["date"],
            message="业务日期未在人员表中出现",
            action="当天业务进入错误文件；矩阵不增加该日期",
        )

    for day in sorted(empty_people_days):
        add_issue(
            issues,
            "提示",
            personnel_path,
            people_sheet,
            people_lines[day],
            day=day,
            field=PERSONNEL_FIELDS["people"],
            message="当天有业务，但人数为空",
            action="当天整行置0；当天业务进入错误文件",
        )

    for day in sorted(zero_people_days):
        add_issue(
            issues,
            "提示",
            personnel_path,
            people_sheet,
            people_lines[day],
            day=day,
            field=PERSONNEL_FIELDS["people"],
            value=0,
            message=(
                "需要核实的异常样本：人数为0，"
                f"业务数量为{source_daily[day]:g}"
            ),
            action="当天整行置0；当天业务进入错误文件",
        )

    # ---------- 4. 按日期、模块归集 ----------

    grouped_quantities = defaultdict(list)
    error_quantities = defaultdict(list)
    errors = []
    reason_counts = Counter()

    for (
        row_number,
        row,
        day,
        quantity,
        part,
        column_index,
        mapping_error,
    ) in parsed_records:
        reasons = []

        if mapping_error:
            reasons.append(mapping_error)

        if day in missing_date_days:
            reasons.append("缺少当天工时：工时表无该日期记录")
        elif day in empty_people_days:
            reasons.append("缺少当天工时：工时表当天人数为空")
        elif day in zero_people_days:
            reasons.append(
                "人数为0但业务数量大于0，待核实；当天整行置0"
            )

        if reasons:
            # 一条业务记录只导出一次，多个原因合并。
            errors.append(list(row) + ["；".join(reasons)])
            error_quantities[day].append(quantity)
            reason_counts.update(reasons)
        else:
            # 不依赖业务表行顺序。
            grouped_quantities[(day, column_index)].append(quantity)

    for (day, column_index), quantities in grouped_quantities.items():
        try:
            matrix[date_to_row[day]][column_index] = finite_sum(
                quantities
            )
        except ValueError as exc:
            add_issue(
                issues,
                "错误",
                path,
                business_sheet,
                day=day,
                field=f"x{column_index + 1}",
                message=str(exc),
                action="核查数量；本次不生成新矩阵",
            )

    stop_if_errors(issues, report_path)

    # ---------- 5. 写入工时 ----------

    # 统计原本有正工时，但因模块数量全为0而不写入工时的日期。
    zero_module_positive_hours_days = set()

    for day in dates:
        row_index = date_to_row[day]

        # 人员表补齐日期，但业务表当天没有记录：整行0。
        if day not in active_days:
            continue

        # 人数为空：整行0。
        if people[day] is None:
            continue

        # 人数为0、业务数量大于0：整行0。
        if day in zero_people_days:
            continue

        # 统一规则：
        # 当天全部模块数量为0，工时也必须为0。
        # 包括全部模块匹配失败、原始业务数量均为0等情况。
        if all(value == 0 for value in matrix[row_index][:-1]):
            matrix[row_index][-1] = 0

            if hours_by_day[day] > 0:
                zero_module_positive_hours_days.add(day)

            continue

        # 至少一个模块数量大于0，才写入当天完整工时。
        # 部分模块匹配失败时，仍保留成功数量和全天工时。
        matrix[row_index][-1] = hours_by_day[day]

    # ---------- 6. 每日及总数量核对 ----------

    matrix_daily = {}
    error_daily = {}

    try:
        for day in sorted(set(dates) | active_days):
            source_total = source_daily.get(day, 0.0)
            error_total = finite_sum(
                error_quantities.get(day, [])
            )

            matrix_total = (
                finite_sum(matrix[date_to_row[day]][:-1])
                if day in date_to_row
                else 0.0
            )

            combined = finite_sum([matrix_total, error_total])

            if not math.isclose(
                source_total,
                combined,
                rel_tol=1e-10,
                abs_tol=1e-8,
            ):
                raise ValueError(
                    f"{day}：数量核对失败；"
                    f"原始={source_total:g}，"
                    f"矩阵={matrix_total:g}，"
                    f"错误记录={error_total:g}"
                )

            matrix_daily[day] = matrix_total
            error_daily[day] = error_total

        source_total = finite_sum(source_daily.values())
        matrix_total = finite_sum(matrix_daily.values())
        error_total = finite_sum(error_daily.values())

        if not math.isclose(
            source_total,
            finite_sum([matrix_total, error_total]),
            rel_tol=1e-10,
            abs_tol=1e-8,
        ):
            raise ValueError("总数量核对失败")

    except ValueError as exc:
        add_issue(
            issues,
            "错误",
            path,
            business_sheet,
            message=str(exc),
            action="核查数据；本次不生成新矩阵",
        )

    stop_if_errors(issues, report_path)

    # ---------- 7. 去除全零行 ----------

    # 第一个表单：删除包含工时列在内全部为0的行。
    # 保持剩余日期的原有升序，不重新排序。
    filtered_matrix = [
        row
        for row in matrix
        if any(value != 0 for value in row)
    ]

    removed_count = len(matrix) - len(filtered_matrix)

    # ---------- 8. 发布输出 ----------

    publish_files({
        result_path: [
            ("矩阵_去除全零行", filtered_matrix),
            ("矩阵_保留全零行", matrix),
        ],
        error_path: (
            [("错误记录", [header + ["错误原因"]] + errors)]
            if errors
            else None
        ),
        report_path: (
            report_sheets(issues)
            if issues
            else None
        ),
    })

    # ---------- 9. 控制台只输出总览 ----------

    print(
        f"[数据] 业务记录{len(records)}条；"
        f"业务日期{len(active_days)}天；"
        f"人员表日期{len(dates)}天"
    )

    print(
        f"[矩阵] 原矩阵{len(matrix)}行 × {module_count + 1}列；"
        f"去除全零行{removed_count}行；"
        f"保留{len(filtered_matrix)}行"
    )

    print(
        f"[日期] 人员表有日期但无业务{len(no_business_days)}天；"
        f"人数为空且有业务{len(empty_people_days)}天；"
        f"人数为0且业务量大于0：{len(zero_people_days)}天；"
        f"业务日期缺少人员记录{len(missing_date_days)}天"
    )

    print(
        "[工时] 有业务记录、原工时大于0，"
        "但因模块数量全为0而将工时置0："
        f"{len(zero_module_positive_hours_days)}天"
    )

    print(
        f"[记录] 正常归集{len(records) - len(errors)}条；"
        f"错误记录{len(errors)}条"
    )

    if reason_counts:
        summary_counts = Counter()

        for reason, count in reason_counts.items():
            if reason.startswith("备件模块无法匹配自变量："):
                category = "备件模块无法匹配自变量"
            elif reason.startswith("备件模块对应多个自变量："):
                category = "备件模块对应多个自变量"
            else:
                category = reason

            summary_counts[category] += count

        print(
            "[错误分类] 同一记录可能包含多个原因，"
            "分类计数可能重叠"
        )

        for reason, count in summary_counts.most_common():
            print(f"  {reason}：{count}条")

    print(
        f"[核对通过] 原始数量 {source_total:g} = "
        f"矩阵数量 {matrix_total:g} + "
        f"错误记录数量 {error_total:g}"
    )

    print(f"[结果] {result_path}")

    if errors:
        print(f"[错误明细] {error_path}")

    if issues:
        print(f"[预检查报告] {report_path}")
    else:
        print("[预检查] 无异常，无需生成报告")

    return {
        "file": path.name,
        "status": "成功",
        "original_rows": len(matrix),
        "removed_rows": removed_count,
        "retained_rows": len(filtered_matrix),
        "error_records": len(errors),
        "message": (
            f"原{len(matrix)}行，去除{removed_count}行，"
            f"保留{len(filtered_matrix)}行，"
            f"错误记录{len(errors)}条"
        ),
    }


def discover_files():
    pattern = filename_pattern()

    available = sorted(
        (
            path
            for path in DATA_DIR.iterdir()
            if path.is_file()
            and path.suffix.lower() == ".xlsx"
            and not path.name.startswith("~$")
        ),
        key=lambda path: path.name,
    )

    business_files = [
        path
        for path in available
        if re.fullmatch(pattern, path.stem)
    ]

    business_groups = defaultdict(list)

    for path in business_files:
        try:
            period = period_info(path)
        except ValueError:
            continue

        business_groups[period].append(path)

    business_stems = {
        path.stem
        for path in business_files
    }

    orphan_count = sum(
        path.stem.endswith("_P")
        and re.fullmatch(pattern, path.stem[:-2]) is not None
        and path.stem[:-2] not in business_stems
        for path in available
    )

    if orphan_count:
        print(
            f"[扫描] {orphan_count}个人员文件没有对应业务文件，"
            "已跳过"
        )

    selected = (
        [DATA_DIR / name for name in DATA_FILES]
        if DATA_FILES
        else business_files
    )

    selected_groups = defaultdict(list)
    failures = []

    for path in selected:
        try:
            period = period_info(path)
        except ValueError as exc:
            failures.append({
                "file": path.name,
                "status": "失败",
                "message": str(exc),
            })
            continue

        selected_groups[period].append(path)

    valid_paths = []

    for period in sorted(selected_groups):
        paths = selected_groups[period]
        available_for_period = business_groups.get(period, [])

        if len(paths) > 1 or len(available_for_period) > 1:
            names = sorted({
                path.name
                for path in paths + available_for_period
            })

            failures.append({
                "file": "_".join(map(str, period)),
                "status": "失败",
                "message": (
                    "同期间重复文件或重复配置：" + "、".join(names)
                ),
            })
            continue

        valid_paths.append(paths[0])

    return valid_paths, failures


def main():
    validate_config()

    mode_name = (
        "月度模式"
        if PROCESS_MODE == 1
        else "年度表模式"
    )

    print("=" * 64)
    print(f"当前模式：{PROCESS_MODE}（{mode_name}）")
    field_name = "中文字段" if FIELD_MODE == 1 else "英文字段"
    print(f"字段配置：{FIELD_MODE}（{field_name}）")
    print("业务字段：" + " / ".join(get_data_fields().values()))
    print("数量支持中英文逗号千位分隔符，例如：5,640 / 5，640")
    print(f"工时系数：{HOURS_PER_PERSON}小时/人")
    print("业务数据按日期及模块归集，不要求原表按时间排序")
    print("矩阵行按人员表日期升序排列，不补齐自然日")
    print("全部模块数量为0时，工时也为0，整行置0")
    print("第1表单去除全零行；第2表单保留全零行")
    print(f"读取目录：{DATA_DIR}")
    print(f"映射文件：{RESOURCE_FILE}")
    print(f"输出目录：{OUTPUT_DIR}")
    print("=" * 64)

    part_map, module_columns, module_count = mappings()

    print(
        f"[Resource] 基础结构检查通过，模块数量：{module_count}"
    )

    paths, summaries = discover_files()

    if not paths and not summaries:
        example = (
            "2026_9.xlsx"
            if PROCESS_MODE == 1
            else "2026_year.xlsx"
        )
        raise ValueError(
            f"未找到符合当前模式的业务文件，例如：{example}"
        )

    for path in paths:
        print(f"\n{'=' * 64}")
        print(f"开始处理：{path.name}")

        try:
            result = process(
                path,
                part_map,
                module_columns,
                module_count,
            )
            summaries.append(result)

        except Exception as exc:
            summaries.append({
                "file": path.name,
                "status": "失败",
                "message": str(exc),
            })

            print(f"[失败] {path.name}：{exc}")
            print(
                "该文件未完成，已有旧输出不能视为本次结果。"
            )

    print(f"\n{'=' * 64}")
    print("批量处理汇总")

    for item in summaries:
        print(
            f"[{item['status']}] "
            f"{item['file']}：{item['message']}"
        )

    successful = [
        item
        for item in summaries
        if item["status"] == "成功"
    ]

    success_count = len(successful)
    failure_count = len(summaries) - success_count

    print(
        f"\n处理结束：成功{success_count}项，"
        f"失败{failure_count}项。"
    )

    if successful:
        total_original = sum(
            item["original_rows"]
            for item in successful
        )
        total_removed = sum(
            item["removed_rows"]
            for item in successful
        )
        total_retained = sum(
            item["retained_rows"]
            for item in successful
        )
        total_errors = sum(
            item["error_records"]
            for item in successful
        )

        print(
            f"成功文件合计：原矩阵{total_original}行；"
            f"去除全零行{total_removed}行；"
            f"保留{total_retained}行；"
            f"错误记录{total_errors}条。"
        )

    return 1 if failure_count else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"[运行失败] {exc}")
        sys.exit(1)