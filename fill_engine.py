"""
轻型动力触探检测报告自动生成工具 - 填充引擎
基于内容模式匹配，不依赖段落顺序
"""
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os, re


def _all_paragraphs(doc):
    """遍历文档中所有段落 (body, tables, headers)，返回 (location, paragraph)"""
    for p in doc.paragraphs:
        yield ('body', p)
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    yield (f'table_{ti}_{ri}_{ci}', p)
    for si, section in enumerate(doc.sections):
        header = section.header
        if header:
            for p in header.paragraphs:
                yield (f'header_{si}', p)


def _red_runs(p):
    """返回段落中红色run的 (run_index, run, text) 列表"""
    result = []
    for i, run in enumerate(p.runs):
        if run.font.color and run.font.color.rgb:
            if str(run.font.color.rgb).upper() == 'FF0000':
                result.append((i, run, run.text))
    return result


def _set_runs_text(p, new_text, red_only=True):
    """设置段落中红run的文本并统一为黑色。多个红run:第一个放文本,其余清空"""
    BLACK = RGBColor(0, 0, 0)
    reds = _red_runs(p)
    if red_only:
        if reds:
            for i, (ri, run, _) in enumerate(reds):
                run.text = str(new_text) if i == 0 else ''
                run.font.color.rgb = BLACK
            return True
    else:
        for i, run in enumerate(p.runs):
            run.text = str(new_text) if i == 0 else ''
            run.font.color.rgb = BLACK
        return True
    return False


def _set_specific_red_run(p, run_index, new_text):
    """设置段落中指定位置的红色run并移除红色标记"""
    reds = _red_runs(p)
    for ri, run, _ in reds:
        if ri == run_index:
            run.text = str(new_text)
            run.font.color.rgb = RGBColor(0, 0, 0)
            return True
    return False


def _match_and_replace_all_exact(paras, exact_match_text, new_text):
    """仅匹配段落文本完全等于exact_match_text的段落"""
    count = 0
    for loc, p in paras:
        full = p.text.strip()
        if full == exact_match_text:
            if _set_runs_text(p, str(new_text), True):
                count += 1
    return count


def _match_and_replace_contains(paras, contains_text, new_text, exclude_pattern=None):
    """匹配包含特定文本的段落，可选排除模式"""
    count = 0
    for loc, p in paras:
        full = p.text.strip()
        if contains_text in full:
            if exclude_pattern and exclude_pattern in full:
                continue
            if _set_runs_text(p, str(new_text), True):
                count += 1
    return count


def fill_document(template_path, output_path, data):
    """
    核心填充函数
    基于内容模式匹配，稳健可靠
    """
    doc = Document(template_path)
    from lxml import etree as etree2

    # ===== 0.1 空白行清理（必须在 fill 前完成，避免 paragraph 列表损坏）=====
    empty_indices = []
    for pi, p in enumerate(doc.paragraphs):
        if not p.text.strip():
            empty_indices.append(pi)
    to_remove = set()
    for i in range(1, len(empty_indices)):
        if empty_indices[i] == empty_indices[i-1] + 1:
            to_remove.add(empty_indices[i])
    for pi in sorted(to_remove, reverse=True):
        p_elem = doc.paragraphs[pi]._element
        p_elem.getparent().remove(p_elem)

    # 保存并重新加载，刷新 paragraph 列表
    doc.save(output_path)
    doc = Document(output_path)

    # ===== 0.2 首页关键段落间距调整 =====
    # 空白清理把连续空行删到只留1个，需要补2个空段落达到3个空行效果
    # 两个区域："检测报告"↔"工程地点"、"报告编号"↔"湖北建夷"
    from docx.oxml import OxmlElement as _OxmlElement
    spacing_targets = ["检测报告", "报告编号：DT2026-00196"]
    for target in spacing_targets:
        for p in doc.paragraphs:
            stripped = p.text.strip()
            no_space = stripped.replace(' ', '').replace('\u3000', '')
            if stripped == target or no_space == target:
                ref = p._element
                for _ in range(6):
                    blank = _OxmlElement('w:p')
                    ref.addnext(blank)
                    ref = blank
                break

    all_p = list(_all_paragraphs(doc))

    # ===== 核心策略：按段落全文模式精确匹配，避免误伤 =====

    # 1. 工程名称 — 匹配以"宜昌市共同南路"开头的段落（不是"由"开头的概述段）
    for loc, p in all_p:
        full = p.text.strip()
        if full.startswith('宜昌市共同南路'):
            _set_runs_text(p, data.get('project_name', ''), True)

    # 2. 工程地点 — 精确匹配完整文本
    for loc, p in all_p:
        full = p.text.strip()
        if full.startswith('工程地点：') and '宜昌市' in full:
            _set_runs_text(p, data.get('project_location', ''), True)
        elif full == '宜昌市伍家岗区橘乡大道':
            _set_runs_text(p, data.get('project_location', ''), True)

    # 3. 委托单位 — 精确匹配
    for loc, p in all_p:
        full = p.text.strip()
        if full.startswith('委托单位：') and '宜昌市' in full:
            _set_runs_text(p, data.get('client_name', ''), True)
        elif full == '宜昌市城市建设投资开发有限公司':
            _set_runs_text(p, data.get('client_name', ''), True)

    # 4. 报告编号 — 包含"报告编号："和"DT"的段落
    for loc, p in all_p:
        full = p.text.strip()
        if full.startswith('报告编号：') and 'DT' in full:
            _set_runs_text(p, data.get('report_number', ''), True)
    # 页眉
    for loc, p in all_p:
        if loc.startswith('header') and 'DT' in p.text:
            reds = _red_runs(p)
            for ri, run, _ in reds:
                if 'DT' in run.text:
                    run.text = data.get('report_number', '')
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    break

    # 5. 检测日期
    date_str = data.get('test_date', '')
    first_date = date_str.split('、')[0].strip() if '、' in date_str else date_str

    # 5a. 首页检测日期（startswith "检测日期："）
    for loc, p in all_p:
        full = p.text.strip()
        if full.startswith('检测日期：') and not loc.startswith('table'):
            reds = _red_runs(p)
            if len(reds) >= 6:
                parts = re.match(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', first_date)
                if parts:
                    y, m, d = parts.groups()
                    texts = ['20', y[2:], '年', m.zfill(2), '月', f'{d.zfill(2)}日']
                    for i, (ri, run, _) in enumerate(reds):
                        run.text = texts[i] if i < len(texts) else ''
                        run.font.color.rgb = RGBColor(0, 0, 0)
                else:
                    for i, (ri, run, _) in enumerate(reds):
                        run.text = first_date if i == 0 else ''
                        run.font.color.rgb = RGBColor(0, 0, 0)
            break

    # 5b. 短格式日期 "2026.05.08"（精确匹配）
    for loc, p in all_p:
        full = p.text.strip()
        if re.match(r'^\d{4}\.\d{2}\.\d{2}$', full):
            parts = re.match(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', first_date)
            if parts:
                y, m, d = parts.groups()
                short = f'{y}.{m.zfill(2)}.{d.zfill(2)}'
            else:
                short = first_date
            _set_runs_text(p, short, True)

    # 6. 报告日期
    report_date = data.get('report_date', '')
    for loc, p in all_p:
        full = p.text.strip()
        if re.match(r'^2026年\d{2}月\d{2}日$', full):
            reds = _red_runs(p)
            if reds:
                parts = re.match(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', report_date)
                if parts and len(reds) >= 4:
                    y, m, d = parts.groups()
                    texts = [y, '年', m.zfill(2), '月', d.zfill(2), '日']
                    for i, (ri, run, _) in enumerate(reds):
                        run.text = texts[i] if i < len(texts) else ''
                        run.font.color.rgb = RGBColor(0, 0, 0)
                else:
                    for i, (ri, run, _) in enumerate(reds):
                        run.text = report_date if i == 0 else ''
                        run.font.color.rgb = RGBColor(0, 0, 0)

    # 7. 承载力特征值 — 自动加 ≥ 前缀
    caps = data.get('bearing_capacities', '')
    # 确保承载力的每个值都带 ≥ 前缀
    if caps:
        parts = re.split(r'[、,，\s]+', caps)
        formatted = []
        for part in parts:
            part = part.strip()
            if part and not part.startswith('≥'):
                part = '≥' + part
            formatted.append(part)
        caps_display = '、'.join(formatted)
    else:
        caps_display = ''

    for loc, p in all_p:
        full = p.text.strip()
        if full.startswith('≥') and ('100' in full or '200' in full or '120' in full or '130' in full or '150' in full):
            _set_runs_text(p, caps_display, True)

    # 7b. Table1 row[7]: cell[1]=承载力, cell[3]=地基面积
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        if len(t1.rows) > 7:
            if len(t1.rows[7].cells) > 1:
                for p in t1.rows[7].cells[1].paragraphs:
                    _set_runs_text(p, caps_display, True)
        # 地基面积 in cell[3]
        foundation_area = data.get('foundation_area', '')
        if foundation_area and len(t1.rows) > 7 and len(t1.rows[7].cells) > 3:
            for p in t1.rows[7].cells[3].paragraphs:
                _set_runs_text(p, foundation_area, True)

    # 7c. 检测依据 row[5].cell[1] (非红色，直接替换)
    testing_standards_page1 = data.get('testing_standards_page1', '')
    if testing_standards_page1 and len(doc.tables) > 1:
        t1 = doc.tables[1]
        if len(t1.rows) > 5 and len(t1.rows[5].cells) > 1:
            from docx.shared import Pt
            disp = testing_standards_page1
            # 内容过长时自动缩小字体，防止撑高首页把日期挤到下一页
            shrink_size = None
            if len(disp) > 50:
                shrink_size = Pt(9)
            elif len(disp) > 30:
                shrink_size = Pt(10.5)
            for p in t1.rows[5].cells[1].paragraphs:
                for run in p.runs:
                    if 'JGJ' in run.text or 'DB42' in run.text or '检测依据' in run.text:
                        run.text = disp
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        if shrink_size:
                            run.font.size = shrink_size

    # 8. 地基类型/土层 — 仅匹配 Table1 中的独立段落
    for loc, p in all_p:
        full = p.text.strip()
        if full == '天然地基' and loc.startswith('table_1'):
            _set_runs_text(p, data.get('foundation_type', ''), True)

    for loc, p in all_p:
        full = p.text.strip()
        if full == '素填土':
            _set_runs_text(p, data.get('soil_layer', ''), True)

    # 9. 抽检数量/总进尺/检测深度 — 在正文段落中
    # 模板红色run顺序：抽检数量 → 检测深度 → 总进尺
    for loc, p in all_p:
        full = p.text.strip()
        if full.startswith('1、现场检测由委托方进行抽样') and not loc.startswith('table'):
            reds = _red_runs(p)
            if len(reds) >= 3:
                # 抽检数量加"点"后缀（模板中"点"字在红色run之外，直接拼进去）
                sc = str(data.get('sample_count', ''))
                if sc and not sc.endswith('点'):
                    sc += '点'
                reds[0][1].text = sc
                reds[1][1].text = str(data.get('test_depth_meters', ''))
                reds[2][1].text = str(data.get('total_depth', ''))
            break

    # 10. 检测桩号范围 — 仅匹配前有"（具体点位见附图"的段落（桩号描述段）
    for loc, p in all_p:
        full = p.text.strip()
        if '（具体点位见附图' in full and not loc.startswith('table'):
            _set_runs_text(p, data.get('pile_range', ''), True)
            break

    # 11. 检测结论（首页 + 第九章）
    conclusion = data.get('test_conclusion', '')
    for loc, p in all_p:
        full = p.text.strip()
        if '动力触探试验结果统计显示' in full and '换算得出' in full:
            # 正文中的结论段落（非表格）
            reds = _red_runs(p)
            if reds:
                for i, (ri, run, _) in enumerate(reds):
                    run.text = conclusion if i == 0 else ''
                    run.font.color.rgb = RGBColor(0, 0, 0)
            # 也处理结论中分散的数字runs
            # 对于100kPa、6等分散runs，全部合并到第一个red run
            for i, (ri, run, _) in enumerate(reds):
                run.text = conclusion if i == 0 else ''
                run.font.color.rgb = RGBColor(0, 0, 0)

    # ===== 19. 结论、建议 =====
        # ===== 19. 结论、建议 =====
    suggestion_on = data.get('suggestion_on', True)
    suggestion_type = data.get('suggestion_type', 'qualified')
    
    # 根据类型直接生成建议内容
    if suggestion_type == 'qualified':
        suggestion_content = "2、基础施工过程中，望有关部门加强截排水及验槽工作。"
    else:
        suggestion_content = "2、建议对不满足设计要求的地基采取有效方式进行相应处理后再进行下一步施工。"

    for loc, p in all_p:
        full = p.text.strip()
        # 标题行处理
        if full == '九、结论、建议':
            reds = _red_runs(p)
            if reds:
                if suggestion_on:
                    reds[0][1].text = '与建议'
                    reds[0][1].font.color.rgb = RGBColor(0, 0, 0)
                else:
                    for run in p.runs:
                        if run.text == '、建议':
                            run.text = ''
        # 建议段落处理
        elif full.startswith('2') and ('望有关部门' in full or '基础施工' in full or '建议对不满足' in full):
            if suggestion_on:
                reds = _red_runs(p)
                if reds:
                    reds[0][1].text = suggestion_content
                    reds[0][1].font.color.rgb = RGBColor(0, 0, 0)
                else:
                    for i, run in enumerate(p.runs):
                        run.text = suggestion_content if i == 0 else ''
                        run.font.color.rgb = RGBColor(0, 0, 0)
            else:
                for run in p.runs:
                    run.text = ''
                    run.font.color.rgb = RGBColor(0, 0, 0)

    # 12. 参建单位 — 精确匹配完整公司名
    unit_map = {
        '中国兵器工业北方勘察设计研究院有限公司': data.get('project_units', {}).get('survey', ''),
        '宜昌市城市规划设计研究院有限公司': data.get('project_units', {}).get('design', ''),
        '宜昌建投园林有限公司': data.get('project_units', {}).get('construction', ''),
        '湖北虹源工程咨询有限公司': data.get('project_units', {}).get('supervision', ''),
        '杨勇': data.get('project_units', {}).get('manager', ''),
        '宜昌市市政工程质量安全监督站': data.get('project_units', {}).get('quality_station', ''),
    }
    for loc, p in all_p:
        full = p.text.strip()
        if full in unit_map and unit_map[full]:
            _set_runs_text(p, unit_map[full], True)

    # 13. 地质概况概述段落
    geo_mode = data.get('geo_mode', 'full')
    pile = data.get('pile_range', '')
    ft = data.get('foundation_type', '')
    sl = data.get('soil_layer', '')

    for loc, p in all_p:
        full = p.text.strip()
        if full.startswith('由') and '提供的《岩土工程勘察报告》' in full and not loc.startswith('table'):
            if geo_mode == 'simple':
                # 简化模式：替换为 "本次检测{pile_range}，地基类型为{foundation_type}，换填的主要土层为{soil_layer}。"
                simple_text = f'本次检测{pile}，地基类型为{ft}，换填的主要土层为{sl}。'
                _set_runs_text(p, simple_text, False)
            else:
                # 完整模式：逐个红run替换
                reds = _red_runs(p)
                survey = data.get('project_units', {}).get('survey', '')
                pile_handled = False

                for ri, run, old_text in reds:
                    ot = old_text.strip()
                    # 勘察单位
                    if '中国兵器' in ot or '勘察研究' in ot:
                        run.text = survey
                        run.font.color.rgb = RGBColor(0, 0, 0)
                    # 桩号范围（可能分两个run: YS7-YS9 + 雨水管道沟槽）
                    elif ('YS7' in ot or 'YS8' in ot) and not pile_handled:
                        run.text = pile
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        pile_handled = True
                    # 桩号范围后半部分（只有"雨水管道沟槽"等描述文字）
                    elif ('雨水管道沟槽' in ot or '污水管道沟槽' in ot) and pile_handled:
                        run.text = ''
                        run.font.color.rgb = RGBColor(0, 0, 0)
                    # 地基类型
                    elif '天然地基' in ot:
                        run.text = ft
                        run.font.color.rgb = RGBColor(0, 0, 0)
                    # 土层描述
                    elif ot == '素填土':
                        run.text = sl
                        run.font.color.rgb = RGBColor(0, 0, 0)
            break

    # 14. 检测依据 — 第三章 第1条（替换，保留序号格式）
    testing_standards_item1 = data.get('testing_standards_item1', '')
    if testing_standards_item1:
        for loc, p in all_p:
            full = p.text.strip()
            if full == '1、《建筑地基检测技术规范》（JGJ 340-2015）；':
                for i, run in enumerate(p.runs):
                    run.text = f'1、{testing_standards_item1}；' if i == 0 else ''
                    run.font.color.rgb = RGBColor(0, 0, 0)
                break

    # 14b. 检测依据 — 第三章 第2条（替换，保留第3条不变）
    testing_standards_chapter3 = data.get('testing_standards_chapter3', '')
    if testing_standards_chapter3:
        for loc, p in all_p:
            full = p.text.strip()
            if full == '2、《岩土工程勘察规程》（DB42/T 169-2022）；':
                for i, run in enumerate(p.runs):
                    run.text = f'2、{testing_standards_chapter3}；' if i == 0 else ''
                    run.font.color.rgb = RGBColor(0, 0, 0)
                break

    # 15. 地质概况表格 (Table 3: 6x2)
    geo_layers = data.get('geo_layers', [])

    if len(doc.tables) > 3:
        geo_table = doc.tables[3]
        if geo_mode == 'simple':
            while len(geo_table.rows) > 1:
                tr = geo_table.rows[-1]._tr
                geo_table._tbl.remove(tr)
        else:
            # 填充已有行
            for i, layer in enumerate(geo_layers):
                ri = i + 1
                if ri < len(geo_table.rows):
                    row = geo_table.rows[ri]
                    name = layer.get('name', '')
                    # 只填第一个 paragraph，避免两处重复
                    cell0 = row.cells[0]
                    if len(cell0.paragraphs) > 0:
                        _set_runs_text(cell0.paragraphs[0], name, True)
                        # 清空其余 paragraph
                        for pi in range(1, len(cell0.paragraphs)):
                            for run in cell0.paragraphs[pi].runs:
                                run.text = ''
                                run.font.color.rgb = RGBColor(0, 0, 0)
                    desc = layer.get('description', '')
                    for p in row.cells[1].paragraphs:
                        reds = _red_runs(p)
                        if reds:
                            for j, (rj, run, _) in enumerate(reds):
                                run.text = desc if j == 0 else ''
                                run.font.color.rgb = RGBColor(0, 0, 0)
                            break
                else:
                    new_row = geo_table.add_row()
                    new_row.cells[0].text = layer.get('name', '')
                    new_row.cells[1].text = layer.get('description', '')
            # 删除多余行
            while len(geo_table.rows) > len(geo_layers) + 1:
                tr = geo_table.rows[-1]._tr
                geo_table._tbl.remove(tr)

    # 15. 仪器表格 (Table 4: 3x4)
    instruments = data.get('instruments', [])
    if len(doc.tables) > 4 and instruments:
        inst_table = doc.tables[4]
        for ri in range(min(len(instruments), 2)):
            row = inst_table.rows[ri + 1]
            inst = instruments[ri]
            for ci, col_key in enumerate(['name', 'number', 'calib_date', 'cert_number']):
                if ci < len(row.cells):
                    for p in row.cells[ci].paragraphs:
                        _set_runs_text(p, str(inst.get(col_key, '')), True)

    # 16. 原始数据表 (Table 9: 7x3)
    raw_data = data.get('raw_data', [])
    sample_count = int(data.get('sample_count', 0) or 0)
    if sample_count > 0:
        # 按抽检数量补足行数（只补不删）
        while len(raw_data) < sample_count:
            raw_data.append({'point_id': '', 'depth': '', 'blows': ''})
    if len(doc.tables) > 9 and raw_data:
        raw_table = doc.tables[9]
        while len(raw_table.rows) > 1:
            tr = raw_table.rows[-1]._tr
            raw_table._tbl.remove(tr)
        for rd in raw_data:
            new_row = raw_table.add_row()
            for ci, col_key in enumerate(['point_id', 'depth', 'blows']):
                new_row.cells[ci].paragraphs[0].text = str(rd.get(col_key, ''))

        # 17. 汇总表 (Table 9)
    summary_data = data.get('summary_data', [])
    sample_count = int(data.get('sample_count', 0) or 0)
    if sample_count > 0:
        while len(summary_data) < sample_count:
            summary_data.append({'soil_layer': '素填土', 'point_id': '', 'elevation': '', 'avg_blows': '', 'bearing_capacity': ''})
    
    # 查找表9（通过表头识别，不依赖固定索引）
    sum_table = None
    for table in doc.tables:
        if len(table.rows) > 0:
            header_text = ' '.join([cell.text for cell in table.rows[0].cells])
            if '土层' in header_text and '承载力' in header_text:
                sum_table = table
                break
    
    if sum_table is not None and summary_data:
        # 删除所有数据行（保留表头）
        while len(sum_table.rows) > 1:
            tr = sum_table.rows[-1]._tr
            sum_table._tbl.remove(tr)
        
        # 添加新数据行，强制第一列为素填土
        for sd in summary_data:
            new_row = sum_table.add_row()
            # 第一列强制设置为素填土
            new_row.cells[0].text = '素填土'
            # 填充其他列（第2-5列）
            if len(new_row.cells) > 1:
                new_row.cells[1].text = str(sd.get('point_id', ''))
            if len(new_row.cells) > 2:
                new_row.cells[2].text = str(sd.get('elevation', ''))
            if len(new_row.cells) > 3:
                new_row.cells[3].text = str(sd.get('avg_blows', ''))
            if len(new_row.cells) > 4:
                new_row.cells[4].text = str(sd.get('bearing_capacity', ''))
        
        # 确保所有数据行的第一列都是素填土
        for row_idx in range(1, len(sum_table.rows)):
            if sum_table.rows[row_idx].cells[0].text.strip() != '素填土':
                sum_table.rows[row_idx].cells[0].text = '素填土'
        
        # 设置列宽
        tbl_grid = sum_table._tbl.find(qn('w:tblGrid'))
        if tbl_grid is not None:
            grid_cols = tbl_grid.findall(qn('w:gridCol'))
            col_widths = [1134, 1418, 1418, 1418, 1418]
            for i, w in enumerate(col_widths):
                if i < len(grid_cols):
                    grid_cols[i].set(qn('w:w'), str(w))
        
        # 设置单元格格式（垂直居中、水平居中、禁止换行）
        for row in sum_table.rows:
            for cell in row.cells:
                tc = cell._tc
                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is None:
                    tcPr = etree2.SubElement(tc, qn('w:tcPr'))
                noWrap = tcPr.find(qn('w:noWrap'))
                if noWrap is None:
                    etree2.SubElement(tcPr, qn('w:noWrap'))
                vAlign = tcPr.find(qn('w:vAlign'))
                if vAlign is None:
                    vAlign = etree2.SubElement(tcPr, qn('w:vAlign'))
                vAlign.set(qn('w:val'), 'center')
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 合并第一列（素填土垂直合并）
        rows = sum_table.rows
        if len(rows) > 1:
            i = 1
            while i < len(rows):
                cell_text = rows[i].cells[0].text.strip()
                if '素填土' in cell_text:
                    start = i
                    while i < len(rows) and '素填土' in rows[i].cells[0].text.strip():
                        i += 1
                    end = i - 1
                    if start < end:
                        for mi in range(start, end + 1):
                            tc = rows[mi].cells[0]._tc
                            tcPr = tc.find(qn('w:tcPr'))
                            if tcPr is None:
                                tcPr = etree2.SubElement(tc, qn('w:tcPr'))
                            vMerge = tcPr.find(qn('w:vMerge'))
                            if vMerge is None:
                                vMerge = etree2.SubElement(tcPr, qn('w:vMerge'))
                            if mi == start:
                                vMerge.set(qn('w:val'), 'restart')
                            else:
                                for p in rows[mi].cells[0].paragraphs:
                                    for run in p.runs:
                                        run.text = ''
                                if qn('w:val') in vMerge.attrib:
                                    del vMerge.attrib[qn('w:val')]
                else:
                    i += 1
        
        # 设置行高0.9cm
        from docx.shared import Pt
        for row in sum_table.rows:
            row.height = Pt(25.4 * 0.9)  # 0.9cm

        # 18. 合并 素填土 单元格 (Table[10] col 0)
        rows = sum_table.rows
        if len(rows) > 1:
            i = 1
            while i < len(rows):
                cell_text = rows[i].cells[0].text.strip()
                if '素填土' in cell_text:
                    start = i
                    while i < len(rows) and '素填土' in rows[i].cells[0].text.strip():
                        i += 1
                    end = i - 1
                    if start < end:
                        for mi in range(start, end + 1):
                            tc = rows[mi].cells[0]._tc
                            tcPr = tc.find(qn('w:tcPr'))
                            if tcPr is None:
                                tcPr = etree2.SubElement(tc, qn('w:tcPr'))
                            vMerge = tcPr.find(qn('w:vMerge'))
                            if vMerge is None:
                                vMerge = etree2.SubElement(tcPr, qn('w:vMerge'))
                            if mi == start:
                                vMerge.set(qn('w:val'), 'restart')
                            else:
                                for p in rows[mi].cells[0].paragraphs:
                                    for run in p.runs:
                                        run.text = ''
                                if qn('w:val') in vMerge.attrib:
                                    del vMerge.attrib[qn('w:val')]
                else:
                    i += 1

    # 19. Table1 额外字段（工程概况表中跨列合并单元格）
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        if len(t1.rows) > 6:
            if len(t1.rows[6].cells) > 1:
                for p in t1.rows[6].cells[1].paragraphs:
                    _set_runs_text(p, data.get('foundation_type', ''), True)
            if len(t1.rows[6].cells) > 3:
                for p in t1.rows[6].cells[3].paragraphs:
                    _set_runs_text(p, data.get('soil_layer', ''), True)
        if len(t1.rows) > 8:
            if len(t1.rows[8].cells) > 1:
                for p in t1.rows[8].cells[1].paragraphs:
                    sc_table = str(data.get('sample_count', ''))
                    if sc_table and not sc_table.endswith('点'):
                        sc_table += '点'
                    _set_runs_text(p, sc_table, True)
            if len(t1.rows[8].cells) > 3:
                parts = re.match(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', first_date)
                if parts:
                    y, m, d = parts.groups()
                    short_date = f'{y}.{m.zfill(2)}.{d.zfill(2)}'
                else:
                    short_date = first_date
                for p in t1.rows[8].cells[3].paragraphs:
                    _set_runs_text(p, short_date, False)
        if len(t1.rows) > 9:
            if len(t1.rows[9].cells) > 1:
                for p in t1.rows[9].cells[1].paragraphs:
                    _set_runs_text(p, data.get('total_depth', ''), True)
            if len(t1.rows[9].cells) > 3:
                for p in t1.rows[9].cells[3].paragraphs:
                    _set_runs_text(p, data.get('test_depth_range', ''), True)

    # 19b. Table1 Row11 检测结论 — 只替换红色 run（后半句），保留黑色前缀
    conclusion = data.get('test_conclusion', '')
    if conclusion and len(doc.tables) > 1:
        t1 = doc.tables[1]
        if len(t1.rows) > 11:
            # Row11 是合并单元格，Col0 是「检 测 结 论」标签（不动），只改 Col1~Col3
            for ci in range(1, min(len(t1.rows[11].cells), 4)):
                cell = t1.rows[11].cells[ci]
                for p in cell.paragraphs:
                    reds = _red_runs(p)
                    if reds:
                        # 只替换红色 run，保留黑色前缀不变
                        for i, (ri, run, _) in enumerate(reds):
                            run.text = conclusion if i == 0 else ''
                            run.font.color.rgb = RGBColor(0, 0, 0)

    # 20. Table2 工程概况表（第二页）
    if len(doc.tables) > 2:
        t2 = doc.tables[2]
        for row_idx in range(len(t2.rows)):
            if row_idx == 0:
                for ci in range(1, min(len(t2.rows[0].cells), 4)):
                    for p in t2.rows[0].cells[ci].paragraphs:
                        _set_runs_text(p, data.get('project_name', ''), True)
            elif row_idx == 1:
                for ci in range(1, min(len(t2.rows[1].cells), 4)):
                    for p in t2.rows[1].cells[ci].paragraphs:
                        _set_runs_text(p, data.get('project_location', ''), True)
            elif row_idx == 9:
                # 结构型式 / 基础类型
                if len(t2.rows[9].cells) > 1:
                    for p in t2.rows[9].cells[1].paragraphs:
                        _set_runs_text(p, data.get('structure_type', ''), True)
                if len(t2.rows[9].cells) > 3:
                    for p in t2.rows[9].cells[3].paragraphs:
                        _set_runs_text(p, data.get('base_type', ''), True)
            elif row_idx == 10:
                # 设计承载力特征值 / 基底高程
                if len(t2.rows[10].cells) > 1:
                    for p in t2.rows[10].cells[1].paragraphs:
                        _set_runs_text(p, data.get('design_bearing_capacity', ''), True)
                if len(t2.rows[10].cells) > 3:
                    for p in t2.rows[10].cells[3].paragraphs:
                        _set_runs_text(p, data.get('elevation_range', ''), True)
            elif row_idx == 12:
                # 备注
                if len(t2.rows[12].cells) > 1:
                    for p in t2.rows[12].cells[1].paragraphs:
                        _set_runs_text(p, data.get('remarks', ''), True)

    # 21. 附图插入
    images = data.get('images', [])
    if images:
        from docx.shared import Inches, Pt
        
        # 查找"十、附图"段落
        target_para = None
        target_index = None
        for pi, p in enumerate(doc.paragraphs):
            if '十' in p.text and '附图' in p.text:
                target_para = p
                target_index = pi
                break
        
        # 如果没找到，在文档末尾添加
        if target_para is None:
            target_para = doc.add_paragraph()
            target_para.text = '十、附图'
        
        # 在附图标题后插入图片
        for img_info in images:
            img_path = img_info.get('path', '')
            caption = img_info.get('caption', '')
            if img_path and os.path.exists(img_path):
                # 添加图片标题
                if caption:
                    cap_p = doc.add_paragraph()
                    cap_run = cap_p.add_run(caption)
                    cap_run.bold = True
                    cap_run.font.size = Pt(10.5)
                # 添加图片
                img_p = doc.add_paragraph()
                img_run = img_p.add_run()
                img_run.add_picture(img_path, width=Inches(5.5))
                img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 23. 表格间距设为1.00cm — 仅 表8(Table[9]) 和 表9(Table[10])
    #from lxml import etree
    #SPACING_TABLES = [9, 10]  # 表8=原始数据, 表9=结果汇总
    #for ti in SPACING_TABLES:
      #  if ti >= len(doc.tables):
       #     continue
       # table = doc.tables[ti]
      #  tbl_pr = table._tbl.find(qn('w:tblPr'))
       # if tbl_pr is None:
      #      tbl_pr = etree.SubElement(table._tbl, qn('w:tblPr'))
        # Remove existing margins
      #  existing = tbl_pr.find(qn('w:tblCellMar'))
       # if existing is not None:
       #     tbl_pr.remove(existing)
      #  mar = etree.SubElement(tbl_pr, qn('w:tblCellMar'))
      #  for side in ['top', 'left', 'bottom', 'right']:
      #      el = etree.SubElement(mar, qn(f'w:{side}'))
      #      el.set(qn('w:w'), '567')  # 1cm = 567 twips
       #     el.set(qn('w:type'), 'dxa')

        # ===== 24. 首页末尾分页（已禁用，避免空白页）=====
    # 如需分页，取消下面代码的注释
    # last_home_para = None
    # for p in doc.paragraphs:
    #     txt = p.text.strip()
    #     if re.match(r'^\d{4}年\d{2}月\d{2}日$', txt):
    #         last_home_para = p
    # 
    # if last_home_para is not None:
    #     has_page_break = False
    #     for run in last_home_para.runs:
    #         if run._element.find(qn('w:br')) is not None:
    #             has_page_break = True
    #             break
    #     if not has_page_break:
    #         run = last_home_para.add_run()
    #         run._element.append(etree2.Element(qn('w:br'), {qn('w:type'): 'page'}))

    # 25. 表格对齐 + 禁止换行 + 列宽固定 + 行高0.9cm
    if len(doc.tables) > 9:
        raw_table = doc.tables[9]
        # 设置表8列宽（点号:2.5cm, 孔深:2cm, 锤击数:剩余）
        tbl_grid = raw_table._tbl.find(qn('w:tblGrid'))
        if tbl_grid is not None:
            grid_cols = tbl_grid.findall(qn('w:gridCol'))
            col_widths = [1418, 1134, 6804]  # 2.5cm, 2cm, rest (~12cm)
            for i, w in enumerate(col_widths):
                if i < len(grid_cols):
                    grid_cols[i].set(qn('w:w'), str(w))

        from docx.shared import Pt
        for row in raw_table.rows:
            # 设置行高0.9cm
            row.height = Pt(28.35 * 0.9)  # 0.9cm = 约25.4磅
            # 单元格禁止换行 + 对齐 + 垂直居中
            for ci, cell in enumerate(row.cells):
                tc = cell._tc
                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is None:
                    tcPr = etree2.SubElement(tc, qn('w:tcPr'))
                # 禁止换行
                noWrap = tcPr.find(qn('w:noWrap'))
                if noWrap is None:
                    etree2.SubElement(tcPr, qn('w:noWrap'))
                # 垂直居中
                vAlign = tcPr.find(qn('w:vAlign'))
                if vAlign is None:
                    vAlign = etree2.SubElement(tcPr, qn('w:vAlign'))
                vAlign.set(qn('w:val'), 'center')
                # 对齐
                for p in cell.paragraphs:
                    if ci <= 1:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # 26. 表9：全居中 + 合并素填土 + 固定列宽 + 行高0.9cm
    if len(doc.tables) > 10:
        sum_table = doc.tables[10]
        # 设置表9列宽
        tbl_grid = sum_table._tbl.find(qn('w:tblGrid'))
        if tbl_grid is not None:
            grid_cols = tbl_grid.findall(qn('w:gridCol'))
            col_widths = [1134, 1418, 1418, 1418, 1418, 1985]  # 约2,2.5,2.5,2.5,2.5,3.5cm
            for i, w in enumerate(col_widths):
                if i < len(grid_cols):
                    grid_cols[i].set(qn('w:w'), str(w))

        from docx.shared import Pt
        for row in sum_table.rows:
            # 设置行高0.9cm
            row.height = Pt(28.35 * 0.9)
            for cell in row.cells:
                tc = cell._tc
                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is None:
                    tcPr = etree2.SubElement(tc, qn('w:tcPr'))
                # 禁止换行
                noWrap = tcPr.find(qn('w:noWrap'))
                if noWrap is None:
                    etree2.SubElement(tcPr, qn('w:noWrap'))
                # 垂直居中
                vAlign = tcPr.find(qn('w:vAlign'))
                if vAlign is None:
                    vAlign = etree2.SubElement(tcPr, qn('w:vAlign'))
                vAlign.set(qn('w:val'), 'center')
                # 水平居中
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # ===== 删除空白页（跳过含图片的段落）=====
    # 删除分页符后面的空段落
    for pi in range(len(doc.paragraphs) - 1, -1, -1):
        p = doc.paragraphs[pi]
        if not p.text.strip():
            # 跳过含图片/绘图的段落
            has_drawing = any(tag.endswith('}drawing') for tag in [e.tag for e in p._element.iter()])
            if has_drawing:
                continue
            if pi == len(doc.paragraphs) - 1:
                try:
                    p._element.getparent().remove(p._element)
                except:
                    pass
            elif pi > 0:
                prev_p = doc.paragraphs[pi - 1]
                for run in prev_p.runs:
                    br = run._element.find(qn('w:br'))
                    if br is not None and br.get(qn('w:type')) == 'page':
                        try:
                            p._element.getparent().remove(p._element)
                        except:
                            pass
                        break
    
    # 删除文档末尾连续的空段落（跳过含图片的段落）
    while len(doc.paragraphs) > 0 and not doc.paragraphs[-1].text.strip():
        last_p = doc.paragraphs[-1]
        has_drawing = any(tag.endswith('}drawing') for tag in [e.tag for e in last_p._element.iter()])
        if has_drawing:
            break
        try:
            last_p._element.getparent().remove(last_p._element)
        except:
            break

    # ===== 全局扫尾：所有残留红色文字 → 黑色 =====
    BLACK = RGBColor(0, 0, 0)
    # 正文段落
    for p in doc.paragraphs:
        for run in p.runs:
            if run.font.color and run.font.color.rgb and str(run.font.color.rgb).upper() == 'FF0000':
                run.font.color.rgb = BLACK
    # 表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.font.color and run.font.color.rgb and str(run.font.color.rgb).upper() == 'FF0000':
                            run.font.color.rgb = BLACK
    # 页眉/页脚
    for section in doc.sections:
        for hdr in [section.header, section.first_page_header, section.even_page_header]:
            if hdr:
                for p in hdr.paragraphs:
                    for run in p.runs:
                        if run.font.color and run.font.color.rgb and str(run.font.color.rgb).upper() == 'FF0000':
                            run.font.color.rgb = BLACK

    doc.save(output_path)
    return output_path


def _safe_log(msg):
    """安全写入 stderr，避免 Windows 编码崩溃"""
    try:
        sys.stderr.buffer.write((msg + '\n').encode('utf-8', 'replace'))
        sys.stderr.buffer.flush()
    except Exception:
        pass


def _kill_word():
    """强制清理残留 WINWORD 进程"""
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'],
                       capture_output=True, timeout=5)
        time.sleep(2)
    except Exception:
        pass


def _refresh_toc_worker(docx_path, report_number, result):
    """refresh_toc 的 worker 函数，在线程中运行，支持超时控制"""
    import sys, subprocess, time, zipfile, pythoncom

    native_path = os.path.abspath(docx_path)

    # ===== 0. 文件完整性预检 =====
    try:
        with zipfile.ZipFile(native_path, 'r') as zf:
            bad = zf.testzip()
            if bad:
                result['ok'] = False
                result['msg'] = f"[TOC] 文件已损坏，跳过: {bad}"
                return
    except Exception:
        result['ok'] = False
        result['msg'] = "[TOC] 文件不是合法 docx，跳过"
        return

    # ===== 1. 等待文件释放 =====
    time.sleep(1)

    # ===== 2. 初始化 COM =====
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    # ===== 3. 打开文档（带一次重试）=====
    doc = None
    word = None
    for attempt in range(2):
        try:
            import win32com.client
            word = win32com.client.gencache.EnsureDispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(native_path,
                                      ConfirmConversions=False,
                                      ReadOnly=False)
            break  # 成功打开，跳出重试循环
        except Exception as open_err:
            if attempt == 0:
                _kill_word()
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    pass
                continue
            else:
                result['ok'] = False
                result['msg'] = f"[TOC] 无法打开文件: {str(open_err)[:120]}"
                return

    if doc is None:
        result['ok'] = False
        result['msg'] = "[TOC] doc is None"
        return

    # ===== 4. 页眉处理：首页不同 + PAGE 域 =====
    try:
        wdHeaderFooterPrimary = 1
        wdHeaderFooterFirstPage = 2
        wdFieldPage = 33
        wdFieldNumPages = 26
        wdCollapseEnd = 0

        for si in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(si)
            try:
                sec.PageSetup.DifferentFirstPageHeaderFooter = True
            except Exception:
                pass

            # 清空首页页眉
            try:
                hdr_first = sec.Headers(wdHeaderFooterFirstPage)
                hdr_first.Range.Text = ""
            except Exception:
                pass

            # 设置正文页眉：报告编号 + 标题 + 第 {PAGE} 页 共 {NUMPAGES} 页
            try:
                hdr = sec.Headers(wdHeaderFooterPrimary)
                rng = hdr.Range
                rng.Text = ""
                rng.Collapse(wdCollapseEnd)

                # 左侧：报告编号
                if report_number:
                    rng.InsertAfter(report_number)
                    rng = hdr.Range
                    rng.Collapse(wdCollapseEnd)

                # 中间：报告标题（用制表符推到中间位置）
                rng.InsertAfter("\t圆锥动力触探试验检测报告\t")
                rng = hdr.Range
                rng.Collapse(wdCollapseEnd)

                # 右侧：第 {PAGE} 页 共 {NUMPAGES} 页
                rng.InsertAfter("第 ")
                rng = hdr.Range
                rng.Collapse(wdCollapseEnd)

                rng.Fields.Add(rng, wdFieldPage)
                rng = hdr.Range
                rng.Collapse(wdCollapseEnd)

                rng.InsertAfter(" 页 共 ")
                rng = hdr.Range
                rng.Collapse(wdCollapseEnd)

                rng.Fields.Add(rng, wdFieldNumPages)
                rng = hdr.Range
                rng.Collapse(wdCollapseEnd)

                rng.InsertAfter(" 页")

                # 设置页眉段落对齐为两端对齐（让制表符生效）
                try:
                    hdr.Range.ParagraphFormat.TabStops.ClearAll()
                    # 添加居中制表位和右对齐制表位
                    sec.PageSetup.TextColumns.SetWidth(1, sec.PageSetup.TextWidth)
                    # 设置段落对齐
                    hdr.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
                except Exception:
                    pass
            except Exception:
                pass
    except Exception:
        pass

    # ===== 5. 手动重建目录（含超链接） =====
    # 模板 TOC 条目是纯文本：一、项目概况………………………………………………………6（U+2026 点+页码，无 \t）
    # U+2026 过滤区分 TOC 条目与正文标题；old_hdr 非空时跳过旧条目避免重复匹配
    try:
        heading_prefixes = ['一、', '二、', '三、', '四、', '五、',
                           '六、', '七、', '八、', '九、', '十、']
        ellipsis = '\u2026'  # 模板 TOC 条目用 U+2026 做前导点

        # 5a. 收集正文一级标题段落，添加书签
        # 必须同时排除 \t 和 U+2026，否则会把 TOC 条目当成标题
        heading_items = []  # [(text, page_num, bookmark_name)]
        old_hdr_texts = set()  # 去重：一、二、... 各只需一个

        for pi in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(pi)
            txt = p.Range.Text.strip()
            for prefix in heading_prefixes:
                if txt.startswith(prefix) and '\t' not in txt and ellipsis not in txt:
                    # 检查是否已有同前缀的标题（避免收集到重复段落）
                    if prefix not in old_hdr_texts:
                        try:
                            pg = p.Range.Information(3)  # wdActiveEndPageNumber = 3
                        except Exception:
                            pg = 0
                        bookmark_name = f'_TocHdr_{len(heading_items) + 1}'
                        try:
                            doc.Bookmarks.Add(bookmark_name, p.Range)
                        except Exception:
                            pass
                        heading_items.append((txt, pg, bookmark_name))
                        old_hdr_texts.add(prefix)
                        break
            if len(heading_items) >= 10:
                break

        # 5b. 删除 Word TOC 域（如果存在）
        if doc.TablesOfContents.Count > 0:
            for i in range(doc.TablesOfContents.Count, 0, -1):
                try:
                    doc.TablesOfContents(i).Range.Delete()
                except Exception:
                    pass

        # 5c. 定位"目  录"段落和第一个正文标题段落
        toc_para_idx = None
        first_hdr_idx = None
        for pi in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(pi)
            no_space = p.Range.Text.strip().replace(' ', '').replace('\u3000', '')
            if toc_para_idx is None and no_space == '目录':
                toc_para_idx = pi
            if first_hdr_idx is None:
                for prefix in heading_prefixes:
                    if no_space.startswith(prefix) and '\t' not in p.Range.Text and ellipsis not in p.Range.Text:
                        first_hdr_idx = pi
                        break
            if toc_para_idx is not None and first_hdr_idx is not None:
                break

        if toc_para_idx is not None and first_hdr_idx is not None:
            # 5d. 逐段删除"目  录"和第一个标题之间的旧 TOC 条目（从后往前删，防索引错位）
            for pi in range(first_hdr_idx - 1, toc_para_idx, -1):
                try:
                    doc.Paragraphs(pi).Range.Delete()
                except Exception:
                    pass

            # 5e. 重新定位"目  录"（删除后索引变化）
            for pi in range(1, doc.Paragraphs.Count + 1):
                no_space = doc.Paragraphs(pi).Range.Text.strip().replace(' ', '').replace('\u3000', '')
                if no_space == '目录':
                    toc_para_idx = pi
                    break

            toc_p = doc.Paragraphs(toc_para_idx)

            # 计算制表位：右对齐 + 前导点
            page_width = doc.PageSetup.PageWidth
            left_margin = doc.PageSetup.LeftMargin
            right_margin = doc.PageSetup.RightMargin
            tab_pos = page_width - left_margin - right_margin

            # 5f. 逐条插入新 TOC 条目
            rng = toc_p.Range.Duplicate
            rng.Collapse(0)  # wdCollapseEnd

            for text, page_num, bookmark_name in heading_items:
                rng.InsertAfter('\r')
                rng.Collapse(0)
                rng.InsertAfter(f'{text}\t{page_num}')
                rng.Collapse(0)

            # 5g. 设置样式 + 制表位 + 超链接
            # 重新定位"目  录"索引（插入后段落数变了）
            for pi in range(1, doc.Paragraphs.Count + 1):
                no_space = doc.Paragraphs(pi).Range.Text.strip().replace(' ', '').replace('\u3000', '')
                if no_space == '目录':
                    toc_para_idx = pi
                    break

            for offset, (text, page_num, bookmark_name) in enumerate(heading_items):
                entry_idx = toc_para_idx + 1 + offset
                if entry_idx <= doc.Paragraphs.Count:
                    try:
                        p = doc.Paragraphs(entry_idx)
                        p.Range.Style = doc.Styles('toc 1')
                        p.Range.ParagraphFormat.SpaceAfter = 6
                        try:
                            p.Range.ParagraphFormat.TabStops.Add(tab_pos, 2, 2)  # wdAlignTabRight=2, wdTabLeaderDots=2
                        except Exception:
                            pass
                        # 超链接：只链标题文本部分（不含 \t 和页码）
                        link_range = p.Range.Duplicate
                        link_range.End = link_range.Start + len(text)
                        try:
                            doc.Hyperlinks.Add(link_range, '', bookmark_name)
                        except Exception:
                            pass
                    except Exception:
                        pass

    except Exception:
        pass

    # ===== 6. 更新所有域 =====
    try:
        doc.Fields.Update()
    except Exception:
        pass

    try:
        doc.ComputeStatistics(2)  # 2 = wdStatisticPages
    except Exception:
        pass

    # ===== 7. 保存并退出 =====
    try:
        doc.Save()
        doc.Close()
        word.Quit()
    except Exception:
        _kill_word()
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    result['ok'] = True
    result['msg'] = "[TOC] 完成"


def refresh_toc(docx_path, report_number=''):
    """用 win32com 刷新目录字段、设置页眉 PAGE 域、更新总页数

    策略：
    1. 先检查文件是否为合法的 docx（zip）格式
    2. 不主动 taskkill，避免杀掉用户正在用的 Word
    3. 打开失败时，再 taskkill 清理残留进程并重试一次
    4. 整个操作有 60 秒超时，防止 Word 卡死导致程序无限等待
    """
    import threading

    native_path = os.path.abspath(docx_path)
    result = {'ok': False, 'msg': ''}

    t = threading.Thread(target=_refresh_toc_worker, args=(docx_path, report_number, result))
    t.daemon = True
    t.start()
    t.join(timeout=60)  # 最多等待 60 秒

    if t.is_alive():
        # 超时了，杀掉残留 Word 进程
        _safe_log("[TOC] 超时（60秒），强制终止 Word 进程")
        _kill_word()
        return False

    if not result.get('ok', False):
        _safe_log(result.get('msg', '[TOC] 未知错误'))

    return result.get('ok', False)




