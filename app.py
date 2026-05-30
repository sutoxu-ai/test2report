"""
轻型动力触探检测报告自动生成工具 v2.0
2026-05-26 重构：新模板 + UI纵向排布 + 按报告顺序
"""
import streamlit as st
import os, re, traceback
from datetime import datetime, timedelta
from fill_engine import fill_document

# ===== 轻型动力触探查表（JGJ340-2015 表8.4.9-1）=====
_REF_TABLES = {
    '一般黏性土地基': {
        'ref': [(5,50),(10,70),(15,90),(20,115),(25,135),(30,160),(35,180),(40,200),(45,220),(50,240)],
        'inputs': [6.4,14,17,21,28,30.5,35.5,41.5,46,48],
    },
    '黏性素填土地基': {
        'ref': [(5,60),(10,80),(15,95),(20,110),(25,120),(30,130),(35,140),(40,150),(45,160),(50,170)],
        'inputs': [5,10,15,23,28.5,32,37,41.3,45,50],
    },
    '粉土、粉细砂土地基': {
        'ref': [(5,55),(10,70),(15,80),(20,90),(25,100),(30,110),(35,125),(40,140),(45,150),(50,160)],
        'inputs': [5,12,None,30,26.7,33.3,37,30,45,50],
    },
}

def _trend_interp(ref_table, x):
    if x is None:
        return None
    pts = sorted(ref_table, key=lambda p: p[0])
    if x <= pts[0][0]:
        return round(pts[0][1], 1)
    if x >= pts[-1][0]:
        return round(pts[-1][1], 1)
    for i in range(len(pts) - 1):
        x1, y1 = pts[i]
        x2, y2 = pts[i + 1]
        if x1 <= x <= x2:
            if x2 == x1:
                return round(y1, 1)
            return round(y1 + (y2 - y1) * (x - x1) / (x2 - x1), 1)
    return None

st.set_page_config(page_title="轻型动力触探检测报告生成", page_icon="📋", layout="wide")

st.markdown("""
<style>
    .stButton > button { font-weight: bold; }
    .section-divider { border-top: 2px solid #e5e7eb; margin: 20px 0; }
    .stTextArea textarea { font-family: 'Consolas', monospace; font-size: 13px; }
    .fixed-label { color: #888; font-size: 12px; }
</style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.docx")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
PROJECTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "projects.json")

def get_output_path(filename=None):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    if filename:
        return os.path.join(OUTPUT_DIR, filename)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(OUTPUT_DIR, f"检测报告_{timestamp}.docx")

def build_report_filename(report_number, project_name, suggestion_type):
    safe_number = report_number.strip().replace('/', '-').replace('\\', '-')
    safe_name = project_name.strip().replace('/', '-').replace('\\', '-')
    result = '合格' if suggestion_type == 'qualified' else '不合格'
    return f"{safe_number}{safe_name}轻型{result}.docx"

def _load_projects():
    import json
    if os.path.exists(PROJECTS_FILE):
        try:
            with open(PROJECTS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def _save_all_projects(projects):
    import json
    with open(PROJECTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(projects, f, ensure_ascii=False, indent=2)

def _get_all_keys():
    return [
        'project_name', 'project_location', 'client_name', 'test_dates', 'date_count',
        'report_date', 'report_number', 'bearing_capacities', 'foundation_type',
        'soil_layer', 'sample_count', 'total_depth', 'test_depth_range',
        'test_depth_meters', 'pile_range', 'test_conclusion', 'geo_mode',
        'geo_description', 'simple_pile_range', 'simple_foundation_type', 'simple_soil_layer',
        'foundation_area', 'testing_standards_page1', 'testing_items',
        'test_unit_info',
        'survey_company', 'design_company', 'construction_company', 'supervision_company',
        'construction_unit', 'quality_station', 'geo_layers', 'instruments',
        'raw_data', 'summary_data', 'suggestion_on', 'suggestion_type',
        'witness', 'structure_type', 'base_type', 'base_elevation', 'table1_remark',
    ]

def _save_current_project():
    import json, copy
    projects = _load_projects()
    pname = st.session_state.get('project_name', '未命名项目').strip()
    if not pname:
        return
    data = {}
    for k in _get_all_keys():
        val = st.session_state.get(k)
        if isinstance(val, (str, int, float, bool, list, dict, type(None))):
            data[k] = copy.deepcopy(val)
    projects[pname] = data
    _save_all_projects(projects)

def _load_project(pname):
    projects = _load_projects()
    if pname in projects:
        data = projects[pname]
        for k, v in data.items():
            st.session_state[k] = v
        return True
    return False

def init_state():
    defaults = {
        'project_name': '宜昌市共同南路（共同东路-桔乡路）市政工程',
        'project_location': '宜昌市伍家岗区橘乡大道',
        'client_name': '宜昌市城市建设投资开发有限公司',
        'test_dates': ['2026年05月08日'],
        'report_date': '2026年05月10日',
        'report_number': 'DT2026-00196',
        'bearing_capacities': '≥120、≥130、≥150',
        'foundation_type': '天然地基',
        'soil_layer': '素填土',
        'sample_count': '6',
        'total_depth': '2.70',
        'test_depth_range': '0.00～0.45',
        'test_depth_meters': '0.45',
        'pile_range': 'YS7-YS9雨水管道沟槽',
        'test_conclusion': '所检测6个点地基土承载力特征值均大于100kPa，符合设计要求。',
        'geo_mode': '完整',
        'suggestion_on': True,
        'suggestion_type': 'qualified',
        'foundation_area': '200',
        'testing_standards_page1': 'JGJ 340-2015、DB42/T 169-2022',
        'testing_items': [
            '《建筑地基检测技术规范》（JGJ 340-2015）',
            '《岩土工程勘察规程》（DB42/T 169-2022）',
        ],
        'survey_company': '中国兵器工业北方勘察设计研究院有限公司',
        'design_company': '宜昌市城市规划设计研究院有限公司',
        'construction_company': '宜昌建投园林有限公司',
        'supervision_company': '湖北虹源工程咨询有限公司',
        'construction_unit': '宜昌市城市建设投资开发有限公司',
        'quality_station': '宜昌市市政工程质量安全监督站',
        'elevation_range': '81.782～81.959',
        'date_count': 1,
        'simple_pile_range': 'K0+534~K1+160段过街污水管道基础',
        'simple_foundation_type': '换填地基',
        'simple_soil_layer': '黏土',
        'witness': '杨勇',
        'structure_type': '——',
        'base_type': '沟槽基础',
        'base_elevation': '81.782～81.959',
        'table1_remark': '——',
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    if 'instruments' not in st.session_state:
        st.session_state.instruments = [
            {'name': '轻型动力触探仪', 'number': 'S8-172', 'calib_date': '2026-09-11', 'cert_number': 'ZJS8-1712025017'},
            {'name': '钢卷尺', 'number': 'S9-22', 'calib_date': '2027-02-04', 'cert_number': '2026CD02051152'},
        ]

    if 'geo_layers' not in st.session_state:
        st.session_state.geo_layers = [
            {'name': '第①层：填土（Q4ml）', 'description': '杂色，稍湿～湿，松散～稍密，主要成分为黏性土，含植物根系、碎石，偶见混凝土块等。全场均有分布，本层工程性质较差，建议挖除。层厚0.20～8.80m，层底埋深0.20～8.80m，层底标高77.82～105.52m。场地K1+140～K1+297.942里程范围内分布较厚填土，块径5cm~10cm，回填年限约10年，回填方式为机械回填，欠固结，均匀性较差。'},
            {'name': '第②层：粉质黏土（Q4al+pl）', 'description': '灰黑色，可塑，土质不匀，含砂颗粒及姜石，无摇振反应，干强度及韧性中等。该土层在场地内局部分布，工程性质一般，建议进行处理。层厚0.70～7.80m，层底埋深7.00～10.50m，层底高程76.33～82.10m。'},
            {'name': '第③层：粉砂岩（K1w）', 'description': '褐红色～青灰色，泥钙质胶结，层状结构，中厚层状构造，层理发育，清晰可见，岩体上部风化裂隙发育。岩石主要由长石、石英及黏土矿物组成，透水性差，断面含有少量灰白色、浅棕红色泥斑，夹灰白色粗砂岩。节理裂隙较发育。根据岩石风化程度不同，划分为两个亚层。'},
            {'name': '第③1层：强风化粉砂岩', 'description': '褐红色，风化裂隙较发育，可见原岩结构，色泽暗淡，岩芯风化呈短柱状，部分成块状，一般柱长4~8cm，岩芯采取率70%左右，钻进较快，饱水状态易崩解。岩体完整程度为较完整，属极软岩类别，岩体基本质量等级为Ⅴ类。层厚3.50～7.00m，层底埋深4.80～14.80m，层底高程71.69～100.32m。'},
            {'name': '第③2层：中风化粉砂岩', 'description': '褐红色～青灰色，岩芯呈长柱状，一般柱长10~25cm。岩体完整程度为较完整，岩石饱和状态单轴抗压强度标准值Rc=2.9MPa，属极软岩类别，岩体基本质量等级为Ⅴ类。岩芯采取率80%～90%，RQD值75%～85%左右，为易软化岩石，暴露地表后，遇水易风化。最大揭露厚度25.00m，层顶高程71.69～100.32m。'},
        ]

    if 'raw_data' not in st.session_state:
        st.session_state.raw_data = [
            {'point_id': '1#（YS7+10）', 'depth': '0.00~0.45', 'blows': '42、51'},
            {'point_id': '2#（YS7+20）', 'depth': '0.00~0.45', 'blows': '41、51'},
            {'point_id': '3#（YS7+30）', 'depth': '0.00~0.45', 'blows': '43、51'},
            {'point_id': '4#（YS8+10）', 'depth': '0.00~0.45', 'blows': '42、51'},
            {'point_id': '5#（YS8+20）', 'depth': '0.00~0.45', 'blows': '41、51'},
            {'point_id': '6#（YS8+30）', 'depth': '0.00~0.45', 'blows': '42、51'},
        ]

    if 'summary_data' not in st.session_state:
        st.session_state.summary_data = [
            {'soil_layer': '素填土', 'point_id': '1#（YS7+10）', 'elevation': '81.782', 'avg_blows': '42.0', 'bearing_capacity': '208'},
            {'soil_layer': '素填土', 'point_id': '2#（YS7+20）', 'elevation': '81.812', 'avg_blows': '41.0', 'bearing_capacity': '204'},
            {'soil_layer': '素填土', 'point_id': '3#（YS7+30）', 'elevation': '81.841', 'avg_blows': '43.0', 'bearing_capacity': '212'},
            {'soil_layer': '素填土', 'point_id': '4#（YS8+10）', 'elevation': '81.902', 'avg_blows': '42.0', 'bearing_capacity': '208'},
            {'soil_layer': '素填土', 'point_id': '5#（YS8+20）', 'elevation': '81.931', 'avg_blows': '41.0', 'bearing_capacity': '204'},
            {'soil_layer': '素填土', 'point_id': '6#（YS8+30）', 'elevation': '81.959', 'avg_blows': '42.0', 'bearing_capacity': '208'},
        ]

    if 'appendix_images' not in st.session_state:
        st.session_state.appendix_images = []

init_state()
st.title('📋 轻型动力触探检测报告生成工具')

# ===== 历史 =====
projects = _load_projects()
project_names = list(projects.keys())
col_psel, col_new, col_del = st.columns([4, 1, 1])
with col_psel:
    sel_project = st.selectbox('📂 选择项目', ['-- 新建/当前项目 --'] + project_names, key='_project_sel')
    if sel_project != '-- 新建/当前项目 --' and sel_project:
        if st.session_state.get('_last_loaded') != sel_project:
            _load_project(sel_project)
            st.session_state['_last_loaded'] = sel_project
            st.rerun()
with col_new:
    st.markdown('<br>', unsafe_allow_html=True)
    if st.button('💾 保存当前', key='save_proj', use_container_width=True):
        _save_current_project()
        st.success('已保存')
        st.rerun()
with col_del:
    st.markdown('<br>', unsafe_allow_html=True)
    if sel_project != '-- 新建/当前项目 --':
        if st.button('🗑️ 删除', key='del_proj', use_container_width=True):
            projects = _load_projects()
            if sel_project in projects:
                del projects[sel_project]
                _save_all_projects(projects)
                st.session_state['_last_loaded'] = ''
                st.rerun()

# ===== 一、封面信息 =====
with st.expander("封面信息", expanded=True):
    st.session_state.project_name = st.text_input('工程名称', st.session_state.project_name)
    st.session_state.report_number = st.text_input('报告编号', st.session_state.report_number)
    st.session_state.project_location = st.text_input('工程地点', st.session_state.project_location)
    st.session_state.client_name = st.text_input('委托单位', st.session_state.client_name)

    st.markdown('**检测日期**')
    ca, cb = st.columns([1, 4])
    with ca:
        dc = st.number_input('日期数量', 1, 99, st.session_state.date_count, key='_date_count')
        st.session_state.date_count = dc
    while len(st.session_state.test_dates) < dc:
        st.session_state.test_dates.append('')
    while len(st.session_state.test_dates) > dc:
        st.session_state.test_dates.pop()
    for i in range(dc):
        st.session_state.test_dates[i] = st.text_input(
            f'日期{i+1}（如2026年05月08日）', st.session_state.test_dates[i], key=f'td_{i}')

    # 报告日期自动计算
    first_date_raw = st.session_state.test_dates[0].strip() if st.session_state.test_dates else ''
    try:
        date_match = re.match(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', first_date_raw)
        if date_match:
            y, m, d = int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3))
            report_dt = datetime(y, m, d) + timedelta(days=2)
            auto_date = f'{report_dt.year}年{report_dt.month:02d}月{report_dt.day:02d}日'
            st.session_state.report_date = auto_date
    except:
        pass
    st.session_state.report_date = st.text_input('报告日期（检测日期+2天，可手动修改）', st.session_state.report_date)

# ===== 一、项目概况（内含首页检测信息）=====
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
with st.expander("一、项目概况", expanded=True):
    st.caption('以下内容输出到报告第3页"首页"和第5页"表1 项目概况"')

    st.markdown('**首页信息（第3页）**')
    st.session_state.testing_standards_page1 = st.text_input('检测依据', st.session_state.testing_standards_page1,
        help='填入首页表检测依据行的显示文本，如：JGJ 340-2015、DB42/T 169-2022')
    st.session_state.bearing_capacities = st.text_input(
        '设计承载力特征值（kPa）', st.session_state.bearing_capacities,
        help='自动添加≥前缀，顿号分隔')
    st.session_state.foundation_area = st.text_input('地基面积（㎡）', st.session_state.foundation_area)
    st.session_state.foundation_type = st.text_input('地基类型', st.session_state.foundation_type)
    st.session_state.soil_layer = st.text_input('土层', st.session_state.soil_layer)
    st.session_state.sample_count = st.text_input('抽检数量（如"6点"）', st.session_state.sample_count)
    st.session_state.total_depth = st.text_input('总进尺（m）', st.session_state.total_depth)
    st.session_state.test_depth_range = st.text_input('检测深度范围（如0.00～0.45）', st.session_state.test_depth_range)
    st.session_state.test_depth_meters = st.text_input('最大检测深度（第六章用，如0.45）', st.session_state.test_depth_meters)
    st.session_state.pile_range = st.text_input('检测位置（只填检测部位，如YS7-YS9雨水管道沟槽）', st.session_state.pile_range,
        help='模板已固定后缀"（具体点位见附图1）"，不需要重复填写')

    st.markdown('**检测结论**')
    st.caption('前半段"根据动力触探试验结果统计显示，参照JGJ 340-2015表8.4.9-1换算得出："已固定在模板中不动，只需填写后半段')
    st.session_state.test_conclusion = st.text_area(
        '结论后半段', st.session_state.test_conclusion, height=60,
        help='如：所检测6个点地基土承载力特征值均大于100kPa，符合设计要求。')

    st.markdown('---')
    st.markdown('**项目概况表（第5页 表1）**')

    st.session_state.construction_unit = st.text_input('建设单位', st.session_state.construction_unit)
    st.session_state.survey_company = st.text_input('勘察单位', st.session_state.survey_company)
    st.session_state.design_company = st.text_input('设计单位', st.session_state.design_company)
    st.session_state.construction_company = st.text_input('施工单位', st.session_state.construction_company)
    st.session_state.supervision_company = st.text_input('监理单位', st.session_state.supervision_company)
    st.session_state.witness = st.text_input('见证人', st.session_state.get('witness', '杨勇'))
    st.session_state.quality_station = st.text_input('质量监督站', st.session_state.quality_station)
    st.text_input('结构型式', value=st.session_state.get('structure_type', '——'), key='structure_type')
    st.text_input('基础类型', value=st.session_state.get('base_type', '沟槽基础'), key='base_type')
    st.text_input('基底高程（m）', value=st.session_state.get('base_elevation', '81.782～81.959'), key='base_elevation')
    st.text_input('备注（表1独立字段）', value=st.session_state.get('table1_remark', '——'), key='table1_remark',
        help='填入第5页表1底部的"备注"栏，通常填"——"或具体内容')

    st.caption('<span class="fixed-label">💡 以下为模板固定内容，不需要在工具中修改：检测性质=委托、检测目的=根据轻型动力触探击数判定地基土承载力特征值、检测项目=地基土承载力、检测方法=采用轻型(10kg)动力触探试验、证书编号=——</span>', unsafe_allow_html=True)

# ===== 检测单位信息 =====
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
with st.expander("检测单位信息", expanded=False):
    st.text_area(
        '检测单位信息（可粘贴替换）',
        value=st.session_state.get('test_unit_info',
            '检测单位：湖北建夷检验检测中心有限公司（盖章）\n'
            '地    址：湖北省宜昌市高新区汉宜大道205号\n'
            '邮    编：443000\n'
            '电    话：0717-7108205\n'
            '传    真：0717-6448611\n'
            '监督电话：0717-6448856'),
        height=160, key='test_unit_info')

# ===== 二、地质概况 =====
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
with st.expander("二、地质概况", expanded=True):
    geo_mode = st.radio('有无地勘单位',
        ['有地勘报告', '无地勘报告'],
        index=0 if st.session_state.geo_mode == '完整' else 1, horizontal=True,
        help='选择"无地勘报告"时，地质概况只输出一句话，且表2不生成')
    st.session_state.geo_mode = '完整' if geo_mode == '有地勘报告' else '简化'

    if st.session_state.geo_mode == '完整':
        st.text_area(
            '地质概况描述（将替换模板中红色文本，支持^{上标}和_{下标}标记）',
            value=st.session_state.get('geo_description',
                '由中国兵器工业北方勘察设计研究院有限公司提供的《岩土工程勘察报告》，'
                '该场地埋深30.00m深度范围内，场地土为第四系全新统填土（Q4^{ml}）、'
                '粉质黏土（Q4^{al+pl}）和白垩系下统五龙组粉砂岩（K1w）组成，'
                '从上至下共分为3个工程地质主层和2个工程地质亚层，场地岩土层概况见表2，'
                '本次检测YS7-YS9雨水管道沟槽，地基类型为天然地基，主要土层为素填土。'),
            height=120, key='geo_description')

        st.markdown("---")
        st.markdown('**表2 — 岩土层概况（逐条人工录入）**')
        layers = st.session_state.geo_layers
        to_del = []
        for i, layer in enumerate(layers):
            st.markdown(f'**第 {i+1} 层**')
            cols_top = st.columns([4, 1])
            with cols_top[0]:
                layers[i]['name'] = st.text_input('岩土名称', layer.get('name', ''), key=f'gn_{i}', label_visibility='collapsed', placeholder='岩土名称（如：填土（Q4ml））')
            with cols_top[1]:
                if st.button('✕ 删除', key=f'gdel_{i}'):
                    to_del.append(i)
            layers[i]['description'] = st.text_area('土层描述', layer.get('description', ''), key=f'gd_{i}', height=80, label_visibility='collapsed', placeholder='土层描述（如：杂色，稍湿～湿，松散～稍密...）')
            if i < len(layers) - 1:
                st.markdown("---")
        for i in sorted(to_del, reverse=True):
            layers.pop(i)
        st.markdown("")
        col_ga, col_gb = st.columns([1, 1])
        with col_ga:
            if st.button('+ 添加地层行', key='gadd'):
                layers.append({'name': '', 'description': ''})
                st.rerun()
        with col_gb:
            if st.button('🗑️ 清空表2', key='gclr'):
                st.session_state.geo_layers = []
                st.rerun()
    else:
        st.markdown('**无地勘报告模式下，地质概况仅输出以下一句话：**')
        st.caption('模板中地质概况段落将只显示检测范围+地基类型+土层信息')
        st.session_state.simple_pile_range = st.text_input('检测范围', st.session_state.simple_pile_range)
        st.session_state.simple_foundation_type = st.text_input('地基类型', st.session_state.simple_foundation_type)
        st.session_state.simple_soil_layer = st.text_input('主要土层', st.session_state.simple_soil_layer)

# ===== 三、检测依据 =====
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
with st.expander("三、检测依据", expanded=False):
    items = st.session_state.testing_items
    to_del_i = []
    for i, item in enumerate(items):
        c1, c2 = st.columns([8, 0.5])
        with c1:
            items[i] = st.text_input(f'第{i+1}条', item, key=f'ti_{i}')
        with c2:
            if i >= 2 and st.button('✕', key=f'tidel_{i}'):
                to_del_i.append(i)
    for i in sorted(to_del_i, reverse=True):
        items.pop(i)
    if st.button('+ 添加检测依据', key='tiadd'):
        items.append('')
        st.rerun()
    fixed_item_num = len(items) + 1
    st.caption(f'💡 固定末尾：{fixed_item_num}、本工程设计文件及相关要求。')

# ===== 四、现场检测及检测仪器 =====
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
with st.expander("四、现场检测及检测仪器", expanded=True):
    st.caption('人工逐条录入，支持历史记录保存')

    h1, h2, h3, h4, h5 = st.columns([2.5, 2, 2, 2.5, 0.5])
    with h1: st.markdown('**仪器名称**')
    with h2: st.markdown('**编号**')
    with h3: st.markdown('**校准日期**')
    with h4: st.markdown('**证书编号**')
    with h5: st.markdown('&nbsp;', unsafe_allow_html=True)

    instruments = st.session_state.instruments
    to_del_i = []
    for i in range(len(instruments)):
        inst = instruments[i]
        c1, c2, c3, c4, c5 = st.columns([2.5, 2, 2, 2.5, 0.5])
        with c1: inst['name'] = st.text_input('仪器名', inst['name'], key=f'in_{i}', label_visibility='collapsed')
        with c2: inst['number'] = st.text_input('编号', inst['number'], key=f'inum_{i}', label_visibility='collapsed')
        with c3: inst['calib_date'] = st.text_input('日期', inst['calib_date'], key=f'idate_{i}', label_visibility='collapsed')
        with c4: inst['cert_number'] = st.text_input('证书号', inst['cert_number'], key=f'icert_{i}', label_visibility='collapsed')
        with c5:
            if len(instruments) > 1 and st.button('✕', key=f'idel_inst_{i}'):
                to_del_i.append(i)
    for i in sorted(to_del_i, reverse=True):
        instruments.pop(i)
    if st.button('+ 添加仪器', key='iadd'):
        instruments.append({'name': '', 'number': '', 'calib_date': '', 'cert_number': ''})
        st.rerun()

# ===== 五、检测数据分析（表8 + 表9）=====
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
with st.expander("五、检测数据分析（表8 + 表9）", expanded=True):
    st.markdown("### 📊 表8 — N10动力触探原始数据")
    st.caption('点号和孔深人工录入，锤击数列从轻型计算表 Excel 粘贴解析')

    raw_data = st.session_state.raw_data

    # 锤击数粘贴
    blows_paste = st.text_area(
        '粘贴锤击数（只粘贴第3列，每行一个值，如 42、51）',
        placeholder='42、51\n41、51\n43、51\n42、51\n41、51\n42、51',
        height=80, key='blows_paste')

    c_parse, c_clr = st.columns([1, 1])
    with c_parse:
        if st.button('📋 解析锤击数', key='pr1', use_container_width=True):
            lines = blows_paste.strip().split('\n')
            new_blows = [l.strip() for l in lines if l.strip()]
            if new_blows:
                # 先清空旧数据，再按新数据行数重建
                st.session_state.raw_data = []
                raw_data = st.session_state.raw_data
                for idx, b in enumerate(new_blows):
                    raw_data.append({
                        'point_id': '',  # 点号留空，让用户手动填写
                        'depth': '',
                        'blows': b
                    })
                st.success(f'已解析 {len(new_blows)} 行锤击数，请手动填写点号和孔深')
                st.rerun()
    with c_clr:
        if st.button('🗑️ 清空表8', key='clr_r', use_container_width=True):
            st.session_state.raw_data = []
            st.session_state.summary_data = []  # 同步清空表9
            st.rerun()

    # 表8逐行编辑
    if raw_data:
        for j, rd_item in enumerate(raw_data):
            c_a, c_b, c_c, c_d = st.columns([2.5, 2, 2, 0.4])
            with c_a:
                raw_data[j]['point_id'] = st.text_input('点号', rd_item.get('point_id', ''),
                    key=f'ptid_{j}', label_visibility='collapsed', placeholder=f'点号{j+1}')
            with c_b:
                raw_data[j]['depth'] = st.text_input('孔深(m)', rd_item.get('depth', ''),
                    key=f'dpt_{j}', label_visibility='collapsed', placeholder='0.00~0.45')
            with c_c:
                raw_data[j]['blows'] = st.text_input('锤击数', rd_item.get('blows', ''),
                    key=f'blw_{j}', label_visibility='collapsed')
            with c_d:
                if st.button('✕', key=f'rdel_{j}'):
                    raw_data.pop(j)
                    st.rerun()
        col_r1, _ = st.columns([1, 3])
        with col_r1:
            if st.button('+ 添加行（表8）', key='radd'):
                raw_data.append({'point_id': '', 'depth': '', 'blows': ''})
                st.rerun()
    else:
        st.info('请粘贴锤击数后点"解析锤击数"，或点"添加行"逐条录入')

    st.markdown("---")
    st.markdown("### 📊 表9 — 地基土承载力确定")
    st.caption('土层和点号从表8自动带入（可手动修改），标高人工录入，平均值从Excel粘贴解析，承载力人工录入')

    sum_data = st.session_state.summary_data

    if sum_data or raw_data:
        # 与表8行数保持同步
        if len(sum_data) < len(raw_data):
            for j in range(len(sum_data), len(raw_data)):
                rd = raw_data[j] if j < len(raw_data) else {}
                sum_data.append({
                    'soil_layer': st.session_state.soil_layer,
                    'point_id': rd.get('point_id', ''),
                    'elevation': '',
                    'avg_blows': '',
                    'bearing_capacity': ''
                })
        # 多出的删除
        while len(sum_data) > max(len(raw_data), 1):
            sum_data.pop()

    # 平均值粘贴
    avg_paste = st.text_area(
        '粘贴平均值（只粘贴第3列，每行一个值 如 42.0）',
        placeholder='42.0\n41.0\n43.0\n42.0\n41.0\n42.0',
        height=80, key='avg_paste')

    cavg, cclr9 = st.columns([1, 1])
    with cavg:
        if st.button('📋 解析平均值', key='ps2', use_container_width=True):
            lines = avg_paste.strip().split('\n')
            new_avgs = [l.strip() for l in lines if l.strip()]
            if new_avgs:
                # 如果 sum_data 为空，先根据 raw_data 初始化
                if not sum_data and raw_data:
                    for j in range(len(raw_data)):
                        rd = raw_data[j]
                        sum_data.append({
                            'soil_layer': st.session_state.soil_layer,
                            'point_id': rd.get('point_id', ''),
                            'elevation': '',
                            'avg_blows': '',
                            'bearing_capacity': ''
                        })
                for idx, a in enumerate(new_avgs):
                    if idx < len(sum_data):
                        sum_data[idx]['avg_blows'] = a
                    else:
                        # 如果 sum_data 行数不够，添加新行
                        pt_id = raw_data[idx].get('point_id', '') if idx < len(raw_data) else ''
                        sum_data.append({
                            'soil_layer': st.session_state.soil_layer,
                            'point_id': pt_id,
                            'elevation': '',
                            'avg_blows': a,
                            'bearing_capacity': ''
                        })
                st.success(f'已更新 {len(new_avgs)} 行平均值')
                st.rerun()
    with cclr9:
        if st.button('🗑️ 清空表9', key='clr_s', use_container_width=True):
            st.session_state.summary_data = []
            st.session_state.raw_data = []  # 同时清空表8，避免自动同步重新填充
            st.rerun()

    if sum_data:
        for j, sd_item in enumerate(sum_data):
            c1s, c2s, c3s, c4s, c5s = st.columns([2, 2.5, 2, 2, 2])
            with c1s:
                sum_data[j]['soil_layer'] = st.text_input(
                    '土层', sd_item.get('soil_layer', st.session_state.soil_layer),
                    key=f'sl_{j}', label_visibility='collapsed')
            with c2s:
                pt = raw_data[j].get('point_id', '') if j < len(raw_data) else sd_item.get('point_id', '')
                sum_data[j]['point_id'] = st.text_input(
                    '点号', pt, key=f'spt_{j}', label_visibility='collapsed')
            with c3s:
                sum_data[j]['elevation'] = st.text_input(
                    '标高(m) ✏️', sd_item.get('elevation', ''),
                    key=f'sev_{j}', label_visibility='collapsed',
                    placeholder='人工录入')
            with c4s:
                sum_data[j]['avg_blows'] = st.text_input(
                    '平均值 📋', sd_item.get('avg_blows', ''),
                    key=f'sbl_{j}', label_visibility='collapsed',
                    placeholder='解析填充')
            with c5s:
                sum_data[j]['bearing_capacity'] = st.text_input(
                    '承载力(kPa) ✏️', sd_item.get('bearing_capacity', ''),
                    key=f'sbc_{j}', label_visibility='collapsed',
                    placeholder='人工录入')
    else:
        st.info('请先完成表8数据录入')

    # 查表工具
    st.markdown("---")
    st.markdown("### 📋 轻型动力触探查表（JGJ340-2015 表8.4.9-1）")

    for tbl_name, tbl_data in _REF_TABLES.items():
        with st.expander(f'{tbl_name} — 承载力查表', expanded=False):
            ref = tbl_data['ref']
            inputs = tbl_data['inputs']
            st.caption('前两列为标准参考值（N₁₀ / fₐₖ），第三列输入平均锤击数，第四列自动线性插值计算承载力')
            for i in range(len(ref)):
                n10, fak = ref[i]
                x_val = inputs[i] if i < len(inputs) else None
                key_avg = f'rt_{tbl_name}_avg_{i}'
                if key_avg not in st.session_state:
                    st.session_state[key_avg] = x_val
                c1, c2, c3, c4 = st.columns([1.2, 1.4, 1.4, 1.4])
                with c1:
                    st.caption(f'N₁₀={n10}')
                with c2:
                    st.caption(f'fₐₖ={fak} kPa')
                with c3:
                    val = st.number_input(
                        '平均锤击数', value=float(st.session_state[key_avg]) if st.session_state[key_avg] is not None else 0.0,
                        step=0.1, key=key_avg, label_visibility='collapsed', format='%.1f'
                    )
                with c4:
                    result = _trend_interp(ref, val) if val else '-'
                    st.caption(f'**{result} kPa**' if result != '-' else '-')

# ===== 六、结论与建议 =====
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
with st.expander("六、结论与建议", expanded=True):
    st.markdown('**检测结论**')
    st.caption('[模板固定前缀] 根据动力触探试验结果统计显示，参照JGJ 340-2015表8.4.9-1换算得出：')
    st.text_input('结论后半段', st.session_state.test_conclusion, disabled=True, key='conclusion_readonly',
        help='结论从项目概况页同步，修改请回"一、项目概况"')

    st.session_state.suggestion_on = st.checkbox('包含"建议"章节', value=st.session_state.suggestion_on)

    if st.session_state.suggestion_on:
        selected_type = st.radio(
            '建议类型',
            ['合格', '不合格'],
            index=0 if st.session_state.suggestion_type == 'qualified' else 1,
            horizontal=True
        )
        st.session_state.suggestion_type = 'qualified' if selected_type == '合格' else 'unqualified'

        if st.session_state.suggestion_type == 'qualified':
            st.success('✅ 2、基础施工过程中，望有关部门加强截排水及验槽工作。')
        else:
            st.error('⚠️ 2、建议对不满足设计要求的地基采取有效方式进行相应处理后再进行下一步施工。')

# ===== 七、附图 =====
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
with st.expander("七、附图（可上传多张图片）", expanded=True):
    imgs = st.session_state.appendix_images

    batch_files = st.file_uploader(
        '📁 批量上传图片', type=['png', 'jpg', 'jpeg', 'bmp'],
        accept_multiple_files=True, key='batch_imgs',
        help='一次选择多张图片，自动按顺序添加'
    )
    if batch_files:
        existing_paths = {img['path'] for img in imgs if img.get('path')}
        new_added = False
        for bf in batch_files:
            temp_path = os.path.join(OUTPUT_DIR, f'appendix_batch_{bf.name}')
            if temp_path in existing_paths:
                continue
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            with open(temp_path, 'wb') as f:
                f.write(bf.getbuffer())
            caption = os.path.splitext(bf.name)[0]
            imgs.append({'caption': caption, 'path': temp_path})
            new_added = True
        if new_added:
            st.rerun()

    to_del = []
    for i, img in enumerate(imgs):
        c1, c2, c3 = st.columns([3, 4, 0.5])
        with c1:
            imgs[i]['caption'] = st.text_input(f'附图{i+1} 标题', img.get('caption', ''), key=f'icap_{i}')
        with c2:
            if img.get('path') and os.path.exists(img.get('path', '')):
                st.caption(f'✅ {os.path.basename(img["path"])}')
            uploaded = st.file_uploader('选择图片', type=['png', 'jpg', 'jpeg', 'bmp'], key=f'ifile_{i}', label_visibility='collapsed')
            if uploaded:
                temp_path = os.path.join(OUTPUT_DIR, f'appendix_{i}_{uploaded.name}')
                os.makedirs(OUTPUT_DIR, exist_ok=True)
                with open(temp_path, 'wb') as f:
                    f.write(uploaded.getbuffer())
                imgs[i]['path'] = temp_path
        with c3:
            if st.button('✕', key=f'idel_img_{i}'):
                to_del.append(i)
    for i in sorted(to_del, reverse=True):
        imgs.pop(i)
    if st.button('+ 添加附图'):
        imgs.append({'caption': '', 'path': ''})
        st.rerun()

# ===== 生成按钮 =====
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
of = st.text_input('输出文件名（可选）', placeholder='留空自动命名')

if st.button('⬇ 生成报告', type='primary', use_container_width=True):
    with st.spinner('正在生成报告...'):
        try:
            test_date_str = '、'.join([d for d in st.session_state.test_dates if d.strip()])

            first_date_raw = st.session_state.test_dates[0].strip() if st.session_state.test_dates else ''
            auto_report_date = ''
            try:
                date_match = re.match(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', first_date_raw)
                if date_match:
                    y, m, d = int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3))
                    report_dt = datetime(y, m, d) + timedelta(days=2)
                    auto_report_date = f'{report_dt.year}年{report_dt.month:02d}月{report_dt.day:02d}日'
            except:
                pass

            report_date = st.session_state.report_date
            if auto_report_date and report_date == '2026年05月10日':
                report_date = auto_report_date

            # 表9数据：直接从 session_state 取，尊重用户手动编辑
            summary_data = st.session_state.summary_data

            # 表8数据：直接从 session_state 取
            raw_data = st.session_state.raw_data

            # 检测依据
            testing_items = st.session_state.get('testing_items', [])
            testing_item1 = testing_items[0] if len(testing_items) > 0 else ''
            testing_item2 = testing_items[1] if len(testing_items) > 1 else ''
            extra_items = testing_items[2:] if len(testing_items) > 2 else []
            fixed_item_num = len(testing_items) + 1

            data = {
                'project_name': st.session_state.project_name,
                'project_location': st.session_state.project_location,
                'client_name': st.session_state.client_name,
                'test_date': test_date_str,
                'report_date': report_date,
                'report_number': st.session_state.report_number,
                'bearing_capacities': st.session_state.bearing_capacities,
                'foundation_type': st.session_state.foundation_type,
                'soil_layer': st.session_state.soil_layer,
                'sample_count': st.session_state.sample_count,
                'total_depth': st.session_state.total_depth,
                'test_depth_range': st.session_state.test_depth_range,
                'test_depth_meters': st.session_state.test_depth_meters,
                'pile_range': st.session_state.pile_range,
                'test_conclusion': st.session_state.test_conclusion,
                'geo_mode': 'full' if st.session_state.geo_mode == '完整' else 'simple',
                'geo_description': st.session_state.get('geo_description', ''),
                'simple_pile_range': st.session_state.simple_pile_range,
                'simple_foundation_type': st.session_state.simple_foundation_type,
                'simple_soil_layer': st.session_state.simple_soil_layer,
                'foundation_area': st.session_state.foundation_area,
                'testing_standards_page1': st.session_state.testing_standards_page1,
                'testing_items': testing_items,
                'extra_testing_items': extra_items,
                'fixed_item_num': fixed_item_num,
                'test_unit_info': st.session_state.get('test_unit_info', ''),
                'project_units': {
                    'construction_unit': st.session_state.get('construction_unit', ''),
                    'survey': st.session_state.survey_company,
                    'design': st.session_state.design_company,
                    'construction': st.session_state.construction_company,
                    'supervision': st.session_state.supervision_company,
                    'quality_station': st.session_state.quality_station,
                },
                'geo_layers': st.session_state.geo_layers,
                'instruments': st.session_state.instruments,
                'raw_data': raw_data,
                'summary_data': summary_data,
                'suggestion_on': st.session_state.suggestion_on,
                'suggestion_type': st.session_state.suggestion_type,
                'images': st.session_state.appendix_images,
                'witness': st.session_state.get('witness', '杨勇'),
                'structure_type': st.session_state.get('structure_type', ''),
                'base_type': st.session_state.get('base_type', ''),
                'base_elevation': st.session_state.get('base_elevation', ''),
                'table1_remark': st.session_state.get('table1_remark', '——'),
            }

            output_path = get_output_path(of if of.strip() else None)
            fill_document(TEMPLATE_PATH, output_path, data)

            auto_filename = build_report_filename(
                st.session_state.report_number,
                st.session_state.project_name,
                st.session_state.suggestion_type
            )
            final_path = os.path.join(OUTPUT_DIR, auto_filename)
            if output_path != final_path:
                import time as _time
                for _retry in range(5):
                    try:
                        if os.path.exists(final_path):
                            os.remove(final_path)
                        os.rename(output_path, final_path)
                        output_path = final_path
                        break
                    except PermissionError:
                        if _retry < 4:
                            _time.sleep(0.5)
                        else:
                            try:
                                import subprocess as _sp
                                _sp.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'],
                                        capture_output=True, timeout=5)
                            except Exception:
                                pass
                            _time.sleep(1)
                            if os.path.exists(final_path):
                                os.remove(final_path)
                            os.rename(output_path, final_path)
                            output_path = final_path

            st.success('✅ 报告生成成功！')
            st.info(f'📁 输出路径：`{output_path}`')
            with open(output_path, 'rb') as f:
                st.download_button('📥 下载报告', f.read(),
                    file_name=os.path.basename(output_path),
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True)
        except Exception as e:
            st.error(f'❌ 生成失败：{e}')
            st.code(traceback.format_exc())

st.caption(f'📄 模板：`{TEMPLATE_PATH}` | 📁 输出：`{OUTPUT_DIR}`')
