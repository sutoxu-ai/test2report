"""
轻型动力触探检测报告自动生成工具 v1.0
"""
import streamlit as st
import os
import re
import traceback
from datetime import datetime, timedelta
from fill_engine import fill_document, refresh_toc

st.set_page_config(page_title="轻型动力触探检测报告生成", page_icon="📋", layout="wide")

st.markdown("""
<style>
    .stButton > button { font-weight: bold; }
    .field-sync { font-size: 11px; color: #4361ee; }
    .section-divider { border-top: 2px solid #e5e7eb; margin: 20px 0; }
    .stTextArea textarea { font-family: 'Consolas', monospace; font-size: 13px; }
</style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.docx")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")


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


PROJECTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "projects.json")


def _load_projects():
    """加载所有项目记录"""
    import json
    if os.path.exists(PROJECTS_FILE):
        try:
            with open(PROJECTS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}


def _save_all_projects(projects):
    """保存所有项目"""
    import json
    with open(PROJECTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(projects, f, ensure_ascii=False, indent=2)


def _get_all_keys():
    """获取所有需要保存的 session_state key"""
    return [
        'project_name', 'project_location', 'client_name', 'test_dates', 'date_count',
        'report_date', 'report_number', 'bearing_capacities', 'foundation_type',
        'soil_layer', 'sample_count', 'total_depth', 'test_depth_range',
        'test_depth_meters', 'pile_range', 'test_conclusion', 'geo_mode',
        'geo_description', 'simple_pile_range', 'simple_foundation_type', 'simple_soil_layer',
        'foundation_area', 'testing_standards_page1', 'testing_items',
        'test_nature', 'test_purpose', 'test_project', 'test_unit_info',
        'survey_company', 'design_company', 'construction_company', 'supervision_company',
        'construction_unit', 'quality_station', 'geo_layers', 'instruments',
        'raw_data', 'summary_data', 'suggestion_on', 'suggestion_type',
        'witness', 'certificate_no', 'structure_type', 'base_type',
        'base_elevation', 'test_method', 'remark',
    ]


def _save_current_project():
    """保存当前项目"""
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
    """加载指定项目到 session_state"""
    projects = _load_projects()
    if pname in projects:
        data = projects[pname]
        for k, v in data.items():
            st.session_state[k] = v
        return True
    return False


# 初始化 session_state
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
        'testing_standards_page1': 'JGJ340-2015、DB42/T169-2022',
        'testing_standards_item1': '《建筑地基检测技术规范》（JGJ 340-2015）',
        'testing_standards_chapter3': '《岩土工程勘察规程》（DB42/T 169-2022）',
        'survey_company': '中国兵器工业北方勘察设计研究院有限公司',
        'design_company': '宜昌市城市规划设计研究院有限公司',
        'construction_company': '宜昌建投园林有限公司',
        'supervision_company': '湖北虹源工程咨询有限公司',
        'project_manager': '',
        'quality_station': '宜昌市市政工程质量安全监督站',
        'elevation_range': '81.782～81.959',
        'date_count': 1,
        'simple_pile_range': 'K0+534~K1+160段过街污水管道基础',
        'simple_foundation_type': '换填地基',
        'simple_soil_layer': '黏土',
        # 新增字段
        'witness': '杨勇',
        'certificate_no': '——',
        'structure_type': '——',
        'base_type': '沟槽基础',
        'base_elevation': '81.782～81.959',
        'test_method': '采用轻型(10kg)动力触探试验',
        'remark': '——',
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
            {'name': '第①层：填土', 'description': '杂色，稍湿～湿，松散～稍密，主要成分为黏性土，含植物根系、碎石，偶见混凝土块等。全场均有分布，本层工程性质较差，建议挖除。层厚0.20～8.80m，层底埋深0.20～8.80m，层底标高77.82～105.52m。'},
            {'name': '第②层：粉质黏土', 'description': '灰黑色，可塑，土质不匀，含砂颗粒及姜石，无摇振反应，干强度及韧性中等。该土层在场地内局部分布，工程性质一般，建议进行处理。层厚0.70～7.80m，层底埋深7.00～10.50m，层底高程76.33～82.10m。'},
            {'name': '第③层：粉砂岩', 'description': '褐红色～青灰色，泥钙质胶结，层状结构，中厚层状构造，层理发育，清晰可见，岩体上部风化裂隙发育。岩石主要由长石、石英及黏土矿物组成，透水性差，断面含有少量灰白色、浅棕红色泥斑，夹灰白色粗砂岩。'},
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

# ===== 项目选择 =====
projects = _load_projects()
project_names = list(projects.keys())
col_psel, col_new, col_del = st.columns([3, 1, 1])
with col_psel:
    sel_project = st.selectbox('📂 选择项目（加载历史记录）', ['-- 新建/当前项目 --'] + project_names, key='_project_sel')
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
                st.success(f'已删除 {sel_project}')
                st.session_state['_last_loaded'] = ''
                st.rerun()

col1, col2 = st.columns([1, 1])

with col1:
    with st.expander("封面信息", expanded=True):
        st.session_state.project_name = st.text_input('工程名称 🔄', st.session_state.project_name)
        st.session_state.report_number = st.text_input('报告编号 🔄', st.session_state.report_number)
        st.session_state.project_location = st.text_input('工程地点 🔄', st.session_state.project_location)
        st.session_state.client_name = st.text_input('委托单位 🔄', st.session_state.client_name)

        st.markdown('**检测日期**（支持多个，顿号分隔）')
        ca, cb = st.columns([1, 3])
        with ca:
            dc = st.number_input('数量', 1, 99, st.session_state.date_count, key='_date_count')
            st.session_state.date_count = dc
        while len(st.session_state.test_dates) < dc:
            st.session_state.test_dates.append('')
        while len(st.session_state.test_dates) > dc:
            st.session_state.test_dates.pop()
        for i in range(dc):
            st.session_state.test_dates[i] = st.text_input(
                f'日期{i+1}（如2026年05月08日）', st.session_state.test_dates[i], key=f'td_{i}')

        # 报告日期 = 首个检测日期 + 2天，自动计算
        first_date_raw = st.session_state.test_dates[0].strip() if st.session_state.test_dates else ''
        auto_report_date = ''
        try:
            date_match = re.match(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', first_date_raw)
            if date_match:
                y, m, d = int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3))
                report_dt = datetime(y, m, d) + timedelta(days=2)
                auto_report_date = f'{report_dt.year}年{report_dt.month:02d}月{report_dt.day:02d}日'
                st.session_state.report_date = auto_report_date
        except:
            pass
        st.session_state.report_date = st.text_input('报告日期（检测日期+2天，可手动修改）', st.session_state.report_date)
        st.session_state.bearing_capacities = st.text_input(
            '设计承载力特征值（kPa）', st.session_state.bearing_capacities,
            help='自动添加≥前缀，顿号分隔，如 ≥120、≥130、≥150')
        st.session_state.foundation_area = st.text_input('地基面积（㎡）', st.session_state.foundation_area)
        st.session_state.testing_standards_page1 = st.text_input(
            '检测依据（首页）', st.session_state.testing_standards_page1)

with col2:
    with st.expander("首页检测信息", expanded=True):
        st.text_input('检测性质', value=st.session_state.get('test_nature', '委托'), key='test_nature')
        st.text_input('检测目的', value=st.session_state.get('test_purpose', '根据轻型动力触探击数判定地基土承载力特征值'), key='test_purpose')
        st.text_input('检测项目', value=st.session_state.get('test_project', '地基土承载力'), key='test_project')
        
        c_home1, c_home2 = st.columns(2)
        with c_home1:
            st.session_state.foundation_type = st.text_input('地基类型', st.session_state.foundation_type)
        with c_home2:
            st.session_state.soil_layer = st.text_input('土层', st.session_state.soil_layer)
        
        c_home3, c_home4 = st.columns(2)
        with c_home3:
            st.session_state.sample_count = st.text_input('抽检数量', st.session_state.sample_count)
        with c_home4:
            st.session_state.total_depth = st.text_input('总进尺（m）', st.session_state.total_depth)
        
        c_depth1, c_depth2 = st.columns(2)
        with c_depth1:
            st.session_state.test_depth_range = st.text_input(
                '检测深度范围（如 0.00～0.45）', st.session_state.test_depth_range)
        with c_depth2:
            st.session_state.test_depth_meters = st.text_input(
                '最大检测深度（第六章，如 0.45）', st.session_state.test_depth_meters)
        st.session_state.pile_range = st.text_input('检测桩号范围（自动填充到检测位置）', st.session_state.pile_range)
        st.session_state.test_conclusion = st.text_area(
            '检测结论', st.session_state.test_conclusion, height=80)
        st.text_input('备注', value=st.session_state.get('remark', '检测位置及数量由施工、监理、设计及建设等单位共同确定。'), key='remark')

        st.markdown("---")
        st.subheader("检测单位基本信息")
        st.text_area(
            '检测单位信息（可粘贴替换）',
            value=st.session_state.get('test_unit_info',
                '检测单位：湖北建夷检验检测中心有限公司（盖章）\n'
                '地    址：湖北省宜昌市高新区汉宜大道205号\n'
                '邮    编：443000\n'
                '电    话：0717-7108205\n'
                '传    真：0717-6448611\n'
                '监督电话：0717-6448856'),
            height=120, key='test_unit_info')

# ===== 二、地质概况 =====
with st.expander("二、地质概况", expanded=True):
    geo_mode = st.radio('报告类型',
        ['完整（有勘察报告）', '简化（无勘察报告）'],
        index=0 if st.session_state.geo_mode == '完整' else 1, horizontal=True)
    st.session_state.geo_mode = '完整' if '完整' in geo_mode else '简化'

    if st.session_state.geo_mode == '完整':
        st.text_area(
            '地质概况描述（替换默认内容）',
            value=st.session_state.get('geo_description',
                '由中国兵器工业北方勘察设计研究院有限公司提供的《岩土工程勘察报告》，'
                '该场地埋深30.00m深度范围内，场地土为第四系全新统填土（Q4ml）、'
                '粉质黏土（Q4al+pl）和白垩系下统五龙组粉砂岩（K1w）组成，'
                '从上至下共分为3个工程地质主层和2个工程地质亚层，场地岩土层概况见表2。'),
            height=90, key='geo_description')

    st.markdown("---")
    st.subheader("项目概况参数")
    
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        st.session_state.survey_company = st.text_input('勘察单位', st.session_state.survey_company)
        st.session_state.design_company = st.text_input('设计单位', st.session_state.design_company)
    with col_p2:
        st.session_state.construction_company = st.text_input('施工单位', st.session_state.construction_company)
        st.session_state.supervision_company = st.text_input('监理单位', st.session_state.supervision_company)
    with col_p3:
        st.text_input('建设单位', value=st.session_state.get('construction_unit', ''), key='construction_unit')
        st.session_state.quality_station = st.text_input('监督单位', st.session_state.quality_station)
    
    # 第二行：基础类型/结构型式/基底高程等
    col_p4, col_p5, col_p6 = st.columns(3)
    with col_p4:
        st.text_input('见证人', value=st.session_state.get('witness', '杨勇'), key='witness')
        st.text_input('基础类型', value=st.session_state.get('base_type', '沟槽基础'), key='base_type')
    with col_p5:
        st.text_input('结构型式', value=st.session_state.get('structure_type', '——'), key='structure_type')
        st.text_input('基底高程（m）', value=st.session_state.get('base_elevation', '81.782～81.959'), key='base_elevation')
    with col_p6:
        st.text_input('证书编号', value=st.session_state.get('certificate_no', '——'), key='certificate_no')
        st.text_input('检测方法', value=st.session_state.get('test_method', '采用轻型(10kg)动力触探试验'), key='test_method')
    
    # 简化模式额外字段
    if st.session_state.geo_mode == '简化':
        st.markdown("---")
        st.session_state.simple_pile_range = st.text_input('桩号范围', st.session_state.simple_pile_range)
        st.session_state.simple_foundation_type = st.text_input('地基类型', st.session_state.simple_foundation_type)
        st.session_state.simple_soil_layer = st.text_input('主要土层', st.session_state.simple_soil_layer)
    else:
        # 完整模式：地层描述表格（支持空格/Tab分隔粘贴）
        st.markdown("---")
        st.subheader('表3 — 岩土层描述（可增删行）')
        
        geo_paste = st.text_area(
            '粘贴表3数据（空格或Tab分隔，每行2列：岩土名称 岩土层描述）',
            placeholder='第①层：填土  杂色，稍湿～湿，松散～稍密...\n第②层：粉质黏土  灰黑色，可塑...',
            height=80, key='geo_paste')
        if st.button('📋 解析表3', key='pgeo'):
            lines = geo_paste.strip().split('\n')
            layers = []
            for line in lines:
                if not line.strip():
                    continue
                parts = re.split(r'\t+|\s{2,}', line.strip(), maxsplit=1)
                if len(parts) >= 2:
                    layers.append({'name': parts[0].strip(), 'description': parts[1].strip()})
                elif len(parts) == 1:
                    layers.append({'name': parts[0].strip(), 'description': ''})
            if layers:
                st.session_state.geo_layers = layers
                st.success(f'已加载 {len(layers)} 行')
                st.rerun()
        
        layers = st.session_state.geo_layers
        to_del = []
        for i, layer in enumerate(layers):
            c1, c2, c3, c4 = st.columns([0.3, 1.5, 4.5, 0.5])
            with c1:
                st.write(str(i + 1))
            with c2:
                layers[i]['name'] = st.text_input('岩土名称', layer.get('name', ''), key=f'gn_{i}', label_visibility='collapsed')
            with c3:
                layers[i]['description'] = st.text_area('描述', layer.get('description', ''), key=f'gd_{i}', height=60, label_visibility='collapsed')
            with c4:
                if st.button('✕', key=f'gdel_{i}'):
                    to_del.append(i)
        for i in sorted(to_del, reverse=True):
            layers.pop(i)
        col_ga, col_gb = st.columns([1, 1])
        with col_ga:
            if st.button('+ 添加地层行', key='gadd'):
                layers.append({'name': '', 'description': ''})
                st.rerun()
        with col_gb:
            if st.button('🗑️ 清空表3', key='gclr'):
                st.session_state.geo_layers = []
                st.rerun()

# ===== 检测依据（第三章）动态添加 =====
with st.expander("检测依据（第三章）", expanded=True):
    if 'testing_items' not in st.session_state:
        st.session_state.testing_items = [
            st.session_state.get('testing_standards_item1', '《建筑地基检测技术规范》（JGJ 340-2015）'),
            st.session_state.get('testing_standards_chapter3', '《岩土工程勘察规程》（DB42/T 169-2022）'),
        ]
    items = st.session_state.testing_items
    to_del_i = []
    for i, item in enumerate(items):
        c1, c2 = st.columns([5, 0.5])
        with c1:
            items[i] = st.text_input(f'第三章-第{i+1}条', item, key=f'ti_{i}', placeholder=f'第{i+1}条检测依据')
        with c2:
            if i >= 2 and st.button('✕', key=f'tidel_{i}'):
                to_del_i.append(i)
    for i in sorted(to_del_i, reverse=True):
        items.pop(i)
    if st.button('+ 添加检测依据', key='tiadd'):
        items.append('')
        st.rerun()
    fixed_item_num = st.number_input('固定末尾条编号', min_value=1, max_value=20, value=len(items)+1, key='fixed_item_num')
    st.caption(f'💡 固定末尾：{fixed_item_num}、本工程设计文件及相关要求。')

# ===== 三、现场检测及仪器 =====
with st.expander("三、现场检测及仪器", expanded=True):
    inst_paste = st.text_area(
        '粘贴仪器数据（4列，Tab/空格分隔）',
        placeholder='轻型动力触探仪\tS8-172\t2026-09-11\tZJS8-1712025017\n钢卷尺\tS9-22\t2027-02-04\t2026CD02051152',
        height=80, key='inst_paste')

    if st.button('📋 解析仪器数据', key='pi'):
        lines = inst_paste.strip().split('\n')
        instruments = []
        for line in lines:
            if not line.strip():
                continue
            parts = re.split(r'\t|\s{2,}', line.strip())
            if len(parts) >= 4:
                instruments.append({
                    'name': parts[0], 'number': parts[1],
                    'calib_date': parts[2], 'cert_number': parts[3],
                })
        if instruments:
            st.session_state.instruments = instruments
            st.success(f'解析了 {len(instruments)} 行')

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
            if len(instruments) > 1 and st.button('✕', key=f'idel_{i}'):
                to_del_i.append(i)
    for i in sorted(to_del_i, reverse=True):
        instruments.pop(i)
    if st.button('+ 添加仪器', key='iadd'):
        instruments.append({'name': '', 'number': '', 'calib_date': '', 'cert_number': ''})
        st.rerun()

# ===== 四、检测结果 =====
with st.expander("四、检测结果（表8原始数据 + 表9汇总）", expanded=True):
    st.markdown("### 📊 表8 — 原始数据")
    st.caption('前两列手输（点号、检测深度），第三列粘贴击数批量解析')
    
    raw_paste = st.text_area(
        '粘贴击数数据（仅第三列，每行一个击数，空格或Tab分隔）',
        placeholder='击数格式（每行3列）：\n1#（YS7+10）  0.00~0.45  42、51\n...',
        height=80, key='raw_paste')
    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button('📋 解析击数', key='pr1', use_container_width=True):
            lines = raw_paste.strip().split('\n')
            rd = []
            for line in lines:
                if not line.strip():
                    continue
                parts = re.split(r'\t+|\s{2,}', line.strip())
                if len(parts) >= 3:
                    rd.append({
                        'point_id': parts[0],
                        'depth': parts[1],
                        'blows': parts[2]
                    })
                elif len(parts) == 1:
                    # 只解析到击数列
                    rd.append({'point_id': '', 'depth': '', 'blows': parts[0]})
            if rd:
                st.session_state.raw_data = rd
                st.success(f'已加载 {len(rd)} 行击数数据')
                st.rerun()
    with c2:
        if st.button('🗑️ 清空表8', key='clr_r', use_container_width=True):
            st.session_state.raw_data = []
            st.rerun()
    
    # 表8：前两列手输，后显示表格
    raw_data = st.session_state.raw_data
    if raw_data:
        # 维持有效行数
        for j, rd_item in enumerate(raw_data):
            c_a, c_b, c_c = st.columns([2, 2, 2])
            with c_a:
                raw_data[j]['point_id'] = st.text_input('点号', rd_item.get('point_id', ''), key=f'ptid_{j}', label_visibility='collapsed')
            with c_b:
                raw_data[j]['depth'] = st.text_input('检测深度(m)', rd_item.get('depth', ''), key=f'dpt_{j}', label_visibility='collapsed')
            with c_c:
                raw_data[j]['blows'] = st.text_input('击数', rd_item.get('blows', ''), key=f'blw_{j}', label_visibility='collapsed')
        col_r1, col_r2 = st.columns([1, 1])
        with col_r1:
            if st.button('+ 添加行（表8）', key='radd'):
                raw_data.append({'point_id': '', 'depth': '', 'blows': ''})
                st.rerun()
        with col_r2:
            if st.button('- 删除最后行（表8）', key='rdel') and len(raw_data) > 0:
                raw_data.pop()
                st.rerun()
    else:
        st.info('请粘贴击数数据后点"解析击数"')
    
    st.markdown("---")
    st.markdown("### 📊 表9 — 结果汇总")
    st.caption('土层从首页自动填充，点号从表8自动填充，标高/平均击数/承载力可粘贴解析')
    
    sum_paste = st.text_area(
        '粘贴表9数据（仅后3列：标高 平均击数 承载力，空格或Tab分隔）',
        placeholder='81.782  42.0  208\n81.812  41.0  204',
        height=80, key='sum_paste')
    cs1, cs2 = st.columns([1, 1])
    with cs1:
        if st.button('📋 解析表9', key='ps2', use_container_width=True):
            lines = sum_paste.strip().split('\n')
            sd = []
            for line in lines:
                if not line.strip():
                    continue
                parts = re.split(r'\t+|\s{2,}', line.strip())
                # 自动从首页和表8补填土层和点号
                soil_val = st.session_state.soil_layer
                pt_id = ''
                if len(sd) < len(raw_data):
                    pt_id = raw_data[len(sd)].get('point_id', '') if len(sd) < len(raw_data) else ''
                if len(parts) >= 3:
                    sd.append({
                        'soil_layer': soil_val,
                        'point_id': pt_id,
                        'elevation': parts[0],
                        'avg_blows': parts[1],
                        'bearing_capacity': parts[2] if len(parts) > 2 else ''
                    })
                elif len(parts) >= 1:
                    sd.append({
                        'soil_layer': soil_val,
                        'point_id': pt_id,
                        'elevation': parts[0] if len(parts) > 0 else '',
                        'avg_blows': parts[1] if len(parts) > 1 else '',
                        'bearing_capacity': parts[2] if len(parts) > 2 else ''
                    })
            if sd:
                st.session_state.summary_data = sd
                st.success(f'已加载 {len(sd)} 行')
                st.rerun()
    with cs2:
        if st.button('🗑️ 清空表9', key='clr_s', use_container_width=True):
            st.session_state.summary_data = []
            st.rerun()
    
    sum_data = st.session_state.summary_data
    if sum_data:
        for j, sd_item in enumerate(sum_data):
            c1s, c2s, c3s, c4s, c5s = st.columns([2, 2, 2, 2, 2])
            with c1s:
                # 土层从首页自动填充，只读
                sum_data[j]['soil_layer'] = st.text_input('土层', st.session_state.soil_layer, key=f'sl_{j}', label_visibility='collapsed', disabled=True)
            with c2s:
                # 点号从表8自动填充，只读
                pt = ''
                if j < len(raw_data):
                    pt = raw_data[j].get('point_id', '')
                sum_data[j]['point_id'] = st.text_input('点号', pt, key=f'spt_{j}', label_visibility='collapsed', disabled=True)
            with c3s:
                sum_data[j]['elevation'] = st.text_input('标高(m)', sd_item.get('elevation', ''), key=f'sev_{j}', label_visibility='collapsed')
            with c4s:
                sum_data[j]['avg_blows'] = st.text_input('平均击数', sd_item.get('avg_blows', ''), key=f'sbl_{j}', label_visibility='collapsed')
            with c5s:
                sum_data[j]['bearing_capacity'] = st.text_input('承载力(kPa)', sd_item.get('bearing_capacity', ''), key=f'sbc_{j}', label_visibility='collapsed')
    else:
        st.info('请粘贴表9数据后点"解析表9"')

# ===== 五、结论与建议 =====
with st.expander("五、结论与建议", expanded=True):
    st.text_area('检测结论（自动同步首页）', st.session_state.test_conclusion, height=80, disabled=True)
    
    st.session_state.suggestion_on = st.checkbox('包含"建议"章节', value=st.session_state.suggestion_on)
    
    if st.session_state.suggestion_on:
        selected_type = st.radio(
            '建议类型',
            ['合格', '不合格'],
            index=0,
            horizontal=True
        )
        st.session_state.suggestion_type = 'qualified' if selected_type == '合格' else 'unqualified'
        
        if st.session_state.suggestion_type == 'qualified':
            st.success('✅ 当前选择：2、基础施工过程中，望有关部门加强截排水及验槽工作。')
        else:
            st.error('⚠️ 当前选择：2、建议对不满足设计要求的地基采取有效方式进行相应处理后再进行下一步施工。')

# ===== 六、附图 =====
with st.expander("六、附图（可上传多张图片）", expanded=True):
    imgs = st.session_state.appendix_images
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
            if st.button('✕', key=f'idel_{i}'):
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

            # 表9：直接从粘贴框重新解析，绕过 session_state.summary_data
            sum_paste_raw = st.session_state.get('sum_paste', '')
            if sum_paste_raw:
                lines = sum_paste_raw.strip().split('\n')
                sd = []
                for j, line in enumerate(lines):
                    if not line.strip():
                        continue
                    parts = re.split(r'\t+|\s{2,}', line.strip())
                    soil_val = st.session_state.soil_layer
                    pt_id = ''
                    if j < len(st.session_state.raw_data):
                        pt_id = st.session_state.raw_data[j].get('point_id', '')
                    if len(parts) >= 3:
                        sd.append({
                            'soil_layer': soil_val,
                            'point_id': pt_id,
                            'elevation': parts[0],
                            'avg_blows': parts[1],
                            'bearing_capacity': parts[2]
                        })
                summary_data = sd if sd else st.session_state.summary_data
            else:
                summary_data = st.session_state.summary_data

            # 表8：直接从粘贴框重新解析
            raw_paste_raw = st.session_state.get('raw_paste', '')
            if raw_paste_raw:
                lines = raw_paste_raw.strip().split('\n')
                rd = []
                for line in lines:
                    if not line.strip():
                        continue
                    parts = re.split(r'\t|\s{2,}', line.strip())
                    if len(parts) >= 3:
                        rd.append({
                            'point_id': parts[0],
                            'depth': parts[1],
                            'blows': parts[2] if len(parts) > 2 else ''
                        })
                raw_data = rd if rd else st.session_state.raw_data
            else:
                raw_data = st.session_state.raw_data




            # 检测依据动态列表
            testing_items = st.session_state.get('testing_items', [])
            testing_item1 = testing_items[0] if len(testing_items) > 0 else ''
            testing_item2 = testing_items[1] if len(testing_items) > 1 else ''
            extra_items = testing_items[2:] if len(testing_items) > 2 else []
            # 使用用户手动设置的 fixed_item_num，否则自动计算
            fixed_item_num = st.session_state.get('fixed_item_num', len(testing_items) + 1)

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
                'testing_standards_item1': testing_item1,
                'testing_standards_chapter3': testing_item2,
                'extra_testing_items': extra_items,
                'fixed_item_num': fixed_item_num,
                'test_nature': st.session_state.get('test_nature', '委托'),
                'test_purpose': st.session_state.get('test_purpose', ''),
                'test_project': st.session_state.get('test_project', ''),
                'test_unit_info': st.session_state.get('test_unit_info', ''),
                'project_units': {
                    'survey': st.session_state.survey_company,
                    'design': st.session_state.design_company,
                    'construction': st.session_state.construction_company,
                    'supervision': st.session_state.supervision_company,
                    'construction_unit': st.session_state.get('construction_unit', ''),
                    'quality_station': st.session_state.quality_station,
                },
                'geo_layers': st.session_state.geo_layers,
                'instruments': st.session_state.instruments,
                'raw_data': raw_data,
                'summary_data': summary_data,
                'suggestion_on': st.session_state.suggestion_on,
                'suggestion_type': st.session_state.suggestion_type,
                'images': st.session_state.appendix_images,
                'witness': st.session_state.get('witness', ''),
                'certificate_no': st.session_state.get('certificate_no', ''),
                'structure_type': st.session_state.get('structure_type', ''),
                'base_type': st.session_state.get('base_type', ''),
                'base_elevation': st.session_state.get('base_elevation', ''),
                'test_method': st.session_state.get('test_method', ''),
                'remark': st.session_state.get('remark', ''),
            }

            output_path = get_output_path(of if of.strip() else None)
            fill_document(TEMPLATE_PATH, output_path, data)
            # refresh_toc(output_path, data.get('report_number', ''))

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

            st.success(f'✅ 报告生成成功！')
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
