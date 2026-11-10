# @Author: YiLong Li
# @Date: 2026/10/26 11:05
# @File: model.py
# @Software: PyCharm


"""
输入：当前脚本目录 / model / 备件模块信息.xlsx
输出：当前脚本目录 / model / 映射表（近似）.xlsx
字典：当前脚本目录 / model / direct / CSV字典文件

主要功能：
    读取备件模块信息，仅处理“模块信息_Operate”表中
    “备件模块”为 #N/A 的记录，完成备件名称与模块的自动归类。

匹配方式：
    优先采用人工修正，再依次进行日用品识别、
    原映射精确匹配、主体名称匹配、名称近似匹配
    及汽车配件字典匹配。
    对无法可靠确定的记录进行临时推定或兜底处理。

输出规则：
    日用品识别：填写“未分类（日用品）”。
    汽车字典匹配：填写“未分类（具体模块名）”。
    近似匹配：填写推荐模块。
    最终无法确定：填写“未分类”。
    人工修正及原映射匹配：保留对应模块名称。

输入位置：
    源文件：model/备件模块信息.xlsx。
    人工修正表：model/人工修正.csv。
    字典文件：model/direct/目录下的CSV文件。

输出位置：
    model/映射表（近似）.xlsx。
    包含“近似”“匹配明细”“处理统计”三个表单。

使用说明：
    保留待处理记录的原始顺序和重复记录。
    明细记录匹配来源、可信等级、候选模块及判定依据。
    临时推定结果标记为低可信，需要人工复核。
    不修改源文件，不自动更新Resource.xlsx。
"""


import csv
import heapq
import os
import re
import sys
import tempfile
import unicodedata
from collections import Counter, defaultdict
from difflib import SequenceMatcher
from functools import lru_cache
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
except ImportError:
    raise SystemExit('缺少依赖，请在 PyCharm 当前解释器安装：pip install openpyxl')

VERSION = 'V3'
BASE = Path(__file__).resolve().parent / 'model'
INPUT = BASE / '备件模块信息.xlsx'
OUTPUT = BASE / '映射表（近似）.xlsx'
DIRECT = BASE / 'direct'
# 默认仍只处理 #N/A，改为 True 时也重新处理空白及“未分类…”记录。
REPROCESS_UNCLASSIFIED = False

MIN_SCORE, MIN_MARGIN, SHORTLIST = 82, 8, 100
INVALID = {'', '#N/A', 'N/A', 'NONE', 'NULL', 'NAN', '未分类', '虚拟模块', '待核实'}
ROLES = {
    '支架': '支架|托架|安装架|保持架', '螺栓': '螺栓|螺钉|螺丝',
    '螺母': '螺母', '护罩': '护罩|护板|防尘罩|隔热罩|隔热板|防护板',
    '线束': '线束|馈线|转接线', '管': '软管|硬管|水管|油管|气管|通风管',
    '盖板': '盖板|堵盖|孔盖|饰盖|端盖|箱盖', '饰板': '饰板|饰条|内衬',
    '密封': '密封条|密封圈|密封垫|油封', '卡扣': '卡扣|管夹|卡箍|卡子',
    '垫片': '垫圈|垫片', '传感器': '传感器|感应塞', '开关': '开关',
    '控制器': '控制器|控制单元|ECU|BCM', '拉索': '拉索|钢索|拉线',
    '玻璃': '玻璃', '轴承': '轴承', '齿轮': '齿轮', '灯': '大灯|雾灯|尾灯|顶灯',
    '泵': '水泵|油泵|真空泵|洗涤泵', '电机': '电机|起动机|发电机',
    '滤清器': '滤芯|滤清器|空滤器', '散热器': '散热器|冷凝器|蒸发器',
    '减震器': '减震器|减振器', '弹簧': '弹簧', '皮带': '皮带',
    '轮胎': '轮胎', '轮毂': '轮毂', '门锁': '门锁|锁体',
    '手柄': '手柄|把手|拉手', '盒': '保险盒|储物盒|工具盒',
    '地毯': '后备箱垫|后备厢垫|行李箱垫|脚垫|地毯',
}


def text(value):
    return '' if value is None else str(value).strip()


@lru_cache(maxsize=150000)
def norm(value):
    value = unicodedata.normalize('NFKC', text(value)).upper()
    value = re.sub(r'^(协议包装|拦截包装|前置包装)', '', value)
    return re.sub(r'[^A-Z0-9\u4e00-\u9fff]', '', value)


@lru_cache(maxsize=100000)
def invalid(value):
    return norm(value) in {norm(x) for x in INVALID} or norm(value).startswith('未分类')


@lru_cache(maxsize=150000)
def core(value):
    value = unicodedata.normalize('NFKC', text(value)).upper()
    # 去括号附加说明，但不去掉主体中的左右、前后和部件名。
    value = re.sub(r'\([^()]*\)|\[[^\[\]]*\]|【[^【】]*】', '', value)
    value = re.sub(r'^(协议包装|拦截包装|前置包装)', '', value)
    value = re.sub(r'^(全新)?(长安汽车|长安|启源|深蓝|欧尚)', '', value)
    value = re.sub(r'[A-Z]+[-_ ]*\d+[A-Z0-9_-]*', '', value)
    value = re.sub(r'铝镁合金|铝合金|不锈钢|碳纤维|\d+\s*KM', '', value)
    for a, b in [('减振', '减震'), ('雨刷', '雨刮'), ('刹车', '制动'),
                 ('消音', '消声'), ('后备厢', '后备箱'), ('电驱动', '电驱')]:
        value = value.replace(a, b)
    return re.sub(r'(分总成|组装件|总成|组件)$', '', norm(value))


@lru_cache(maxsize=100000)
def role(value):
    name = norm(re.sub(r'\([^()]*\)|\[[^\[\]]*\]', '', unicodedata.normalize('NFKC', value)))
    hits = [(name.rfind(w) + len(w), len(w), k)
            for k, words in ROLES.items() for w in words.split('|') if w in name]
    return max(hits)[2] if hits else ''


@lru_cache(maxsize=200000)
def compatible(query, reference, module):
    a, b = role(query), role(reference)
    if a != b and (a or b):
        return False
    q, r = norm(query), norm(reference)
    systems = [('电池', '电瓶'), ('电驱', '电机'), ('制动', '刹车', '驻车'),
               ('燃油', '汽油', '柴油'), ('转向',), ('洗涤', '雨刮', '雨刷')]
    qs = {i for i, words in enumerate(systems) if any(w in q for w in words)}
    rs = {i for i, words in enumerate(systems) if any(w in r for w in words)}
    if qs and rs and qs.isdisjoint(rs):
        return False
    for x, y in [('左', '右'), ('前', '后')]:
        for first, second in [(x, y), (y, x)]:
            if first in module and second not in module and second in q and first not in q:
                return False
    return True


def read_csv(path, required):
    if not path.exists():
        raise FileNotFoundError(f'未找到文件：{path}\n请将本次提供的文件放到该位置。')
    for encoding in ('utf-8-sig', 'gb18030'):
        try:
            with path.open(encoding=encoding, newline='') as f:
                reader = csv.DictReader(f)
                headers = [text(x) for x in (reader.fieldnames or [])]
                if not set(required) <= set(headers):
                    raise ValueError(f'{path.name} 缺少列：{set(required) - set(headers)}')
                reader.fieldnames = headers
                rows = [{k: text(v) for k, v in r.items() if k} for r in reader]
                print(f'【已读取】{path.resolve()}；{len(rows):,} 条', flush=True)
                return rows
        except UnicodeDecodeError:
            continue
    raise ValueError(f'{path.name} 编码无法读取，请保存为 UTF-8 CSV。')


class Matcher:
    def __init__(self, mapping, daily, auto):
        self.refs, self.cores = defaultdict(dict), defaultdict(dict)
        self.modules, self.excluded = set(), 0
        for name, module in mapping:
            name, module = text(name), text(module)
            if invalid(name) or invalid(module):
                self.excluded += 1
                continue
            self.refs[norm(name)][module] = name
            if core(name):
                self.cores[core(name)][module] = name
            self.modules.add(module)
        if not self.refs:
            raise ValueError('映射表没有有效名称与模块。')
        self.keys = list(self.cores)
        self.index = defaultdict(set)
        for key in self.keys:
            for gram in self.grams(key):
                self.index[gram].add(key)
        self.daily_exact, self.daily_words = {}, []
        for r in daily:
            mode = r.get('匹配方式', '') or '包含'
            if mode not in {'精确', '包含'}:
                raise ValueError(f'日用品字典匹配方式错误：{mode}')
            if not norm(r['关键词']):
                continue
            if mode == '精确':
                self.daily_exact[norm(r['关键词'])] = r['关键词']
            else:
                self.daily_words.append((norm(r['关键词']), r['关键词'],
                    tuple(norm(w) for w in r.get('排除词', '').split('|') if w)))
        self.auto = [(core(r['关键词']), r, tuple(norm(w) for w in r.get('排除词', '').split('|') if w)) for r in auto]

    @staticmethod
    def grams(value):
        return {value[i:i + 2] for i in range(len(value) - 1)} or {value}

    def daily_hit(self, name):
        q = norm(name)
        if q in self.daily_exact:
            return '已核对商品名称：' + self.daily_exact[q]
        for word, original, exclusions in self.daily_words:
            if word and word in q and not any(w in q for w in exclusions):
                return '日用品关键词：' + original
        return ''

    def auto_hits(self, name):
        hits = []
        query, normalized = core(name), norm(name)
        for word, r, exclusions in self.auto:
            if not word or word not in query:
                continue
            if any(w in normalized for w in exclusions):
                continue
            if not compatible(name, r['关键词'], r['模块']):
                continue
            hits.append((len(word), r['模块'], r['关键词']))
        hits.sort(key=lambda x: (-x[0], x[1], x[2]))
        return hits

    @lru_cache(maxsize=100000)
    def match(self, name):
        # 返回最终模块、来源、需要复核、依据、前三候选。
        def result(module, stage, reason, candidates=(), review='是'):
            return module, stage, review, reason, list(candidates)

        if invalid(name):
            return result('未分类', '最终兜底', '名称为空或无效')
        daily = self.daily_hit(name)
        if daily:
            return result('未分类（日用品）', '日用品', daily, review='否')
        exact = self.refs.get(norm(name), {})
        if len(exact) == 1:
            m, ref = next(iter(exact.items()))
            return result(m, '原映射精确', '标准化名称一致', [(100, m, ref)], '否')
        if len(exact) > 1:
            return result('未分类', '映射冲突', '同名对应多个模块：' + '；'.join(sorted(exact)))
        simple = self.cores.get(core(name), {})
        simple = {m: r for m, r in simple.items() if compatible(name, r, m)}
        if len(simple) == 1:
            m, ref = next(iter(simple.items()))
            return result(m, '主体名称匹配', '去车型、材质、括号附注后相同', [(100, m, ref)])
        # 以完整的参考主体名称结尾，而非任意短词包含。
        anchors = []
        for key in {core(name)[i:] for i in range(max(0, len(core(name)) - 3))}:
            if len(key) >= 4 and key in self.cores:
                for m, ref in self.cores[key].items():
                    if compatible(name, ref, m):
                        anchors.append((len(key), m, ref))
        if anchors:
            longest = max(x[0] for x in anchors)
            best = [x for x in anchors if x[0] == longest]
            if len({x[1] for x in best}) == 1 and len(simple) <= 1:
                _, m, ref = best[0]
                score = round(100 * SequenceMatcher(None, core(name), core(ref), autojunk=False).ratio(), 2)
                return result(m, '主体词近似', '命中唯一最长主体词：' + core(ref), [(score, m, ref)])
        q = core(name)
        counts = Counter(k for g in self.grams(q) for k in self.index.get(g, ()))
        keys = heapq.nlargest(SHORTLIST, counts, key=lambda k: (2 * counts[k] / (len(self.grams(q)) + len(self.grams(k))), k))
        best = {}
        for key in keys:
            score = round(100 * SequenceMatcher(None, q, key, autojunk=False).ratio(), 2)
            for m, ref in self.cores[key].items():
                if compatible(name, ref, m) and (m not in best or score > best[m][0]):
                    best[m] = (score, m, ref)
        candidates = sorted(best.values(), key=lambda x: (-x[0], x[1], x[2]))[:3]
        hits = self.auto_hits(name)
        top_modules = {m for size, m, _ in hits if hits[0][0] - size < 2} if hits else set()
        confident = (candidates and candidates[0][0] >= MIN_SCORE
                     and (len(candidates) == 1 or candidates[0][0] - candidates[1][0] >= MIN_MARGIN)
                     and len(simple) <= 1)
        if confident and (not top_modules or top_modules == {candidates[0][1]}):
            return result(candidates[0][1], '近似匹配', '主体相似度和分差达标，零件类型相容', candidates)
        if len(top_modules) == 1:
            m = next(iter(top_modules))
            return result(f'未分类（{m}）', '汽车字典', '关键词：' + hits[0][2], candidates)
        note = '字典有多个模块冲突' if len(top_modules) > 1 else '未达到可靠匹配条件'
        return result('未分类', '最终兜底', note + '；候选仅供人工复核', candidates)


def headers(sheet, required):
    row = next(sheet.iter_rows(max_row=1, values_only=True), ())
    names = [text(x) for x in row]
    if any(n and names.count(n) > 1 for n in names):
        raise ValueError(f'{sheet.title} 存在重复表头')
    index = {v: i for i, v in enumerate(names) if v}
    if not set(required) <= index.keys():
        raise ValueError(f'{sheet.title} 缺少列：{set(required) - index.keys()}')
    return index


def put(sheet, values):
    sheet.append(values)
    # 件号和外部文本按文本写入，避免前导零丢失或被当成公式。
    sheet._written_count = getattr(sheet, '_written_count', 0) + 1
    for cell in sheet.iter_rows(min_row=sheet._written_count, max_row=sheet._written_count, max_col=len(values)):
        for item in cell:
            if isinstance(item.value, str):
                item.data_type = 's'
    return sheet._written_count


def format_sheet(sheet, widths):
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = sheet.dimensions
    for cell in sheet[1]:
        cell.fill = PatternFill('solid', fgColor='245A7C')
        cell.font = Font(color='FFFFFF', bold=True)
    for i, width in enumerate(widths, 1):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical='center', wrap_text=True)


def run():
    BASE.mkdir(parents=True, exist_ok=True)
    DIRECT.mkdir(parents=True, exist_ok=True)
    print(f'【版本】{VERSION}', flush=True)
    print(f'【脚本】{Path(__file__).resolve()}', flush=True)
    print(f'【输入】{INPUT}', flush=True)
    print(f'【输出】{OUTPUT}', flush=True)
    print('【处理范围】' + ('#N/A、空白及未分类记录' if REPROCESS_UNCLASSIFIED else '仅 #N/A；不修改原表'), flush=True)
    if not INPUT.exists():
        raise FileNotFoundError(f'未找到输入文件：{INPUT}\n请放入该文件，或修改代码中的 INPUT 文件名。')
    daily = read_csv(DIRECT / '日用品字典.csv', ['关键词'])
    auto = read_csv(DIRECT / '汽车配件字典.csv', ['关键词', '模块'])
    auto = [r for r in auto if norm(r['关键词']) and not invalid(r['模块'])]
    cp = BASE / '人工修正.csv'
    if not cp.exists():
        with cp.open('w', encoding='utf-8-sig', newline='') as f:
            csv.writer(f).writerow(['件号', '备件名称', '确认模块', '备注'])
    corrections = {}
    for r in read_csv(cp, ['件号', '备件名称', '确认模块']):
        m = r['确认模块']
        if not m:
            continue
        key = ('件号', r['件号']) if r['件号'] else ('名称', norm(r['备件名称']))
        if not key[1] or (invalid(m) and not m.startswith('未分类')):
            raise ValueError('人工修正存在空标识或无效模块')
        if key in corrections and corrections[key] != m:
            raise ValueError(f'人工修正冲突：{key}')
        corrections[key] = m
    print('读取源表和字典...', flush=True)
    wb = openpyxl.load_workbook(INPUT, read_only=True, data_only=True)
    try:
        for name in ['映射表', '模块信息_Operate']:
            if name not in wb.sheetnames:
                raise ValueError(f'缺少工作表：{name}')
        mp = headers(wb['映射表'], ['备件名称'])
        mf = '模块名称' if '模块名称' in mp else '备件模块'
        if mf not in mp:
            raise ValueError('映射表缺少模块名称或备件模块列')
        matcher = Matcher(((r[mp['备件名称']], r[mp[mf]]) for r in wb['映射表'].iter_rows(min_row=2, values_only=True)), daily, auto)
        print(f'有效映射名称：{len(matcher.refs):,}；排除无效映射：{matcher.excluded:,}', flush=True)
        source = wb['模块信息_Operate']
        h = headers(source, ['备件名称', '区', '备件模块'])
        pf = next((x for x in ['件号', '备件编码', '零件号'] if x in h), None)
        if not pf:
            raise ValueError('源表缺少件号、备件编码或零件号列，无法输出件号')
        out = openpyxl.Workbook()
        main = out.active
        main.title = '近似'
        detail = out.create_sheet('匹配明细')
        stats_sheet = out.create_sheet('处理统计')
        put(main, ['件号', '货品名称', '区', '备件模块'])
        put(detail, ['原表行号', '备件SKU', '件号', '备件名称', '区', '原分类', '最终模块', '匹配来源', '需要复核', '判定依据', '候选1', '候选2', '候选3'])
        stats = Counter()
        for number, row in enumerate(source.iter_rows(min_row=2, values_only=True), 2):
            original = text(row[h['备件模块']])
            if original.upper() != '#N/A' and not (REPROCESS_UNCLASSIFIED and (not original or original.startswith('未分类'))):
                continue
            name, area, part = text(row[h['备件名称']]), text(row[h['区']]), text(row[h[pf]])
            override = corrections.get(('件号', part)) or corrections.get(('名称', norm(name)))
            result = (override, '人工修正', '否', '件号优先于名称', []) if override else matcher.match(name)
            module, stage, review, reason, candidates = result
            put(main, [part, name, area, module])
            sku = text(row[h['备件SKU']]) if '备件SKU' in h else ''
            refs = [f'{ref} → {m}（{score:.2f}）' for score, m, ref in candidates]
            put(detail, [number, sku, part, name, area, original, module, stage, review, reason] + refs + [''] * (3 - len(refs)))
            color = 'E2F0D9' if review == '否' else ('FCE4D6' if module == '未分类' else 'FFF2CC')
            main.cell(main._written_count, 4).fill = PatternFill('solid', fgColor=color)
            stats[stage] += 1
            if sum(stats.values()) % 2000 == 0:
                print(f'已处理 {sum(stats.values()):,} 行', flush=True)
        put(stats_sheet, ['统计项', '记录数或说明'])
        put(stats_sheet, ['程序版本', VERSION])
        put(stats_sheet, ['日用品字典路径', str(DIRECT / '日用品字典.csv')])
        put(stats_sheet, ['汽车字典路径', str(DIRECT / '汽车配件字典.csv')])
        put(stats_sheet, ['日用品字典条数', len(daily)])
        put(stats_sheet, ['汽车字典有效条数', len(auto)])
        for k, v in stats.items():
            put(stats_sheet, [k, v])
        put(stats_sheet, ['合计', sum(stats.values())])
        put(stats_sheet, ['排除无效映射', matcher.excluded])
        format_sheet(main, [26, 80, 16, 32])
        format_sheet(detail, [12, 23, 26, 65, 12, 24, 32, 22, 12, 65, 70, 70, 70])
        format_sheet(stats_sheet, [30, 90])
        fd, temp = tempfile.mkstemp(suffix='.xlsx', dir=BASE)
        os.close(fd)
        try:
            out.save(temp)
            os.replace(temp, OUTPUT)
        finally:
            if os.path.exists(temp):
                os.remove(temp)
            out.close()
        if not stats:
            print('【提示】本次没有符合条件的记录，请核对输入表的备件模块列。', flush=True)
        print(f'完成：{sum(stats.values()):,} 行\n输出：{OUTPUT}', flush=True)
    finally:
        wb.close()


if __name__ == '__main__':
    try:
        run()
    except PermissionError as exc:
        print(f'【文件被占用或没有权限】请关闭 Excel 后重试。\n{exc}')
        sys.exit(1)
    except Exception as exc:
        print(f'【处理失败】{type(exc).__name__}: {exc}')
        sys.exit(1)