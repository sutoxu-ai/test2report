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
        'raw_paste_count': 0,
        'sum_paste_count': 0,
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

col1, col2 = st.columns([1, 1])

with col1:
    with st.expander("一、基本信息", expanded=True):
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

        st.session_state.report_date = st.text_input('报告日期 🔄', st.session_state.report_date)
        st.caption('💡 报告日期 = 首个检测日期 + 2天（可手动修改）')
        st.session_state.bearing_capacities = st.text_input(
            '设计承载力特征值（kPa）', st.session_state.bearing_capacities,
            help='自动添加≥前缀，顿号分隔，如 ≥120、≥130、≥150')
        st.session_state.foundation_area = st.text_input('地基面积（㎡）', st.session_state.foundation_area)
        st.session_state.testing_standards_page1 = st.text_input(
            '检测依据（首页）', st.session_state.testing_standards_page1)
        st.session_state.testing_standards_item1 = st.text_input(
            '检测依据（第三章-第1条）', st.session_state.testing_standards_item1)
        st.session_state.testing_standards_chapter3 = st.text_input(
            '检测依据（第三章-第2条）', st.session_state.testing_standards_chapter3)

with col2:
    with st.expander("首页检测信息", expanded=True):
        st.session_state.foundation_type = st.text_input('地基类型', st.session_state.foundation_type)
        st.session_state.soil_layer = st.text_input('土层描述', st.session_state.soil_layer)
        st.session_state.sample_count = st.text_input('抽检数量', st.session_state.sample_count)
        st.session_state.total_depth = st.text_input('总进尺（m）', st.session_state.total_depth)
        c_depth1, c_depth2 = st.columns(2)
        with c_depth1:
            st.session_state.test_depth_range = st.text_input(
                '检测深度范围（首页，如 0.00～0.45）', st.session_state.test_depth_range)
        with c_depth2:
            st.session_state.test_depth_meters = st.text_input(
                '最大检测深度（第六章，如 0.45）', st.session_state.test_depth_meters)
        st.session_state.pile_range = st.text_input('检测桩号范围', st.session_state.pile_range)
        st.session_state.test_conclusion = st.text_area(
            '检测结论 🔄', st.session_state.test_conclusion, height=80)
        
        st.markdown("---")
        st.subheader("项目概况表参数")
        
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            st.session_state.witness = st.text_input('见证人', st.session_state.witness)
            st.session_state.structure_type = st.text_input('结构型式', st.session_state.structure_type)
            st.session_state.base_type = st.text_input('基础类型', st.session_state.base_type)
        with col_w2:
            st.session_state.certificate_no = st.text_input('证书编号', st.session_state.certificate_no)
            st.session_state.base_elevation = st.text_input('基底高程（m）', st.session_state.base_elevation)
        
        st.session_state.test_method = st.text_input('检测方法', st.session_state.test_method)
        st.session_state.remark = st.text_input('备注', st.session_state.remark)

# ===== 二、地质概况 =====
with st.expander("二、地质概况", expanded=True):
    geo_mode = st.radio('报告类型',
        ['完整（有勘察报告）', '简化（无勘察报告）'],
        index=0 if st.session_state.geo_mode == '完整' else 1, horizontal=True)
    st.session_state.geo_mode = '完整' if '完整' in geo_mode else '简化'

    st.subheader('参建单位信息')
    uc1, uc2, uc3 = st.columns(3)
    with uc1:
        st.session_state.survey_company = st.text_input('勘察单位', st.session_state.survey_company)
        st.session_state.design_company = st.text_input('设计单位', st.session_state.design_company)
    with uc2:
        st.session_state.construction_company = st.text_input('施工单位', st.session_state.construction_company)
        st.session_state.supervision_company = st.text_input('监理单位', st.session_state.supervision_company)
    with uc3:
        st.session_state.project_manager = st.text_input('项目经理', st.session_state.project_manager)
        st.session_state.quality_station = st.text_input('监督单位', st.session_state.quality_station)

    if st.session_state.geo_mode == '完整':
        st.subheader('地层描述（表格内容，可增删行）')
        layers = st.session_state.geo_layers
        to_del = []
        for i, layer in enumerate(layers):
            c1, c2, c3, c4 = st.columns([0.3, 1.5, 4.5, 0.5])
            with c1:
                st.write(str(i + 1))
            with c2:
                layers[i]['name'] = st.text_input('名称', layer['name'], key=f'gn_{i}', label_visibility='collapsed')
            with c3:
                layers[i]['description'] = st.text_area('描述', layer['description'], key=f'gd_{i}', height=60, label_visibility='collapsed')
            with c4:
                if st.button('✕', key=f'gdel_{i}'):
                    to_del.append(i)
        for i in sorted(to_del, reverse=True):
            layers.pop(i)
        if st.button('+ 添加地层行'):
            layers.append({'name': '', 'description': ''})
            st.rerun()
    else:
        st.info('简化模式：仅保留一行描述')
        st.session_state.simple_pile_range = st.text_input('桩号范围', st.session_state.simple_pile_range)
        st.session_state.simple_foundation_type = st.text_input('地基类型', st.session_state.simple_foundation_type)
        st.session_state.simple_soil_layer = st.text_input('主要土层', st.session_state.simple_soil_layer)

# ===== 三、现场检测及仪器 =====
with st.expander("三、现场检测及仪器（固定2行）", expanded=True):
    inst_paste = st.text_area(
        '粘贴仪器数据（2行4列，Tab分隔）',
        placeholder='轻型动力触探仪\tS8-172\t2026-09-11\tZJS8-1712025017\n钢卷尺\tS9-22\t2027-02-04\t2026CD02051152',
        height=80, key='inst_paste')

    if st.button('📋 解析仪器数据', key='pi'):
        lines = inst_paste.strip().split('\n')
        instruments = []
        for line in lines[:2]:
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

    while len(st.session_state.instruments) < 2:
        st.session_state.instruments.append({'name': '', 'number': '', 'calib_date': '', 'cert_number': ''})

    h1, h2, h3, h4 = st.columns([3, 2, 2, 3])
    with h1: st.markdown('**仪器名称**')
    with h2: st.markdown('**编号**')
    with h3: st.markdown('**校准日期**')
    with h4: st.markdown('**证书编号**')

    for i in range(2):
        inst = st.session_state.instruments[i]
        c1, c2, c3, c4 = st.columns([3, 2, 2, 3])
        with c1: inst['name'] = st.text_input('仪器名', inst['name'], key=f'in_{i}', label_visibility='collapsed')
        with c2: inst['number'] = st.text_input('编号', inst['number'], key=f'inum_{i}', label_visibility='collapsed')
        with c3: inst['calib_date'] = st.text_input('日期', inst['calib_date'], key=f'idate_{i}', label_visibility='collapsed')
        with c4: inst['cert_number'] = st.text_input('证书号', inst['cert_number'], key=f'icert_{i}', label_visibility='collapsed')

# ===== 四、检测结果 =====
with st.expander("四、检测结果（支持增删行 + 批量粘贴）", expanded=True):
    tab1, tab2 = st.tabs(['表8 — 原始数据', '表9 — 结果汇总'])

    with tab1:
        st.caption('格式：点号 / 检测深度(m) / 击数(击/10cm)')
        raw_paste = st.text_area(
            '粘贴表8数据（Tab分隔，每行3列）',
            placeholder='1#（YS7+10）\t0.00~0.45\t42、51\n2#（YS7+20）\t0.00~0.45\t41、51',
            height=80, key='raw_paste')
        if st.button('📋 解析表8', key='pr1'):
            lines = raw_paste.strip().split('\n')
            rd = []
            for line in lines:
                if not line.strip():
                    continue
                parts = re.split(r'\t|\s{2,}', line.strip())
                if len(parts) >= 2:
                    rd.append({
                        'point_id': parts[0],
                        'depth': parts[1] if len(parts) > 1 else '',
                        'blows': parts[2] if len(parts) > 2 else ''
                    })
            if rd:
                st.session_state.raw_data = rd
                st.session_state.raw_paste_count = st.session_state.get('raw_paste_count', 0) + 1
                st.success(f'已替换为 {len(rd)} 行数据')
                st.rerun()
        raw_data = st.session_state.raw_data
        rk = st.session_state.get('raw_paste_count', 0)
        to_del = []
        for i, rd in enumerate(raw_data):
            c1, c2, c3, c4 = st.columns([2.5, 2, 2, 0.7])
            with c1: raw_data[i]['point_id'] = st.text_input('点号', rd['point_id'], key=f'rid_{rk}_{i}', label_visibility='collapsed')
            with c2: raw_data[i]['depth'] = st.text_input('深度', rd['depth'], key=f'rdep_{rk}_{i}', label_visibility='collapsed')
            with c3: raw_data[i]['blows'] = st.text_input('击数', rd['blows'], key=f'rbl_{rk}_{i}', label_visibility='collapsed')
            with c4:
                if st.button('✕', key=f'rdel_{i}'):
                    to_del.append(i)
        for i in sorted(to_del, reverse=True):
            raw_data.pop(i)
        if st.button('+ 添加行', key='ar1'):
            raw_data.append({'point_id': '', 'depth': '', 'blows': ''})
            st.rerun()

    with tab2:
        st.caption('格式：土层 / 点号 / 标高 / 平均击数 / 承载力')
        sum_paste = st.text_area(
            '粘贴表9数据（Tab分隔，每行5列）',
            placeholder='素填土\t1#（YS7+10）\t81.782\t42.0\t208',
            height=80, key='sum_paste')
        if st.button('📋 解析表9', key='ps2'):
            lines = sum_paste.strip().split('\n')
            sd = []
            for line in lines:
                if not line.strip():
                    continue
                parts = re.split(r'\t|\s{2,}', line.strip())
                if len(parts) >= 4:
                    sd.append({
                        'soil_layer': parts[0] if parts[0] else '素填土',
                        'point_id': parts[1] if len(parts) > 1 else '',
                        'elevation': parts[2] if len(parts) > 2 else '',
                        'avg_blows': parts[3] if len(parts) > 3 else '',
                        'bearing_capacity': parts[4] if len(parts) > 4 else ''
                    })
            if sd:
                st.session_state.summary_data = sd
                st.session_state.sum_paste_count = st.session_state.get('sum_paste_count', 0) + 1
                st.success(f'已替换为 {len(sd)} 行数据')
                st.rerun()
        sum_data = st.session_state.summary_data
        sk = st.session_state.get('sum_paste_count', 0)
        to_del = []
        for i, sd in enumerate(sum_data):
            c1, c2, c3, c4, c5, c6 = st.columns([1.5, 1.8, 1.3, 1.3, 1.3, 0.7])
            with c1: sum_data[i]['soil_layer'] = st.text_input('土层', sd.get('soil_layer', '素填土'), key=f'ss_{sk}_{i}', label_visibility='collapsed')
            with c2: sum_data[i]['point_id'] = st.text_input('点号', sd['point_id'], key=f'sid_{sk}_{i}', label_visibility='collapsed')
            with c3: sum_data[i]['elevation'] = st.text_input('标高', sd['elevation'], key=f'se_{sk}_{i}', label_visibility='collapsed')
            with c4: sum_data[i]['avg_blows'] = st.text_input('击数', sd['avg_blows'], key=f'sa_{sk}_{i}', label_visibility='collapsed')
            with c5: sum_data[i]['bearing_capacity'] = st.text_input('承载力', sd['bearing_capacity'], key=f'sb_{sk}_{i}', label_visibility='collapsed')
            with c6:
                if st.button('✕', key=f'sdel_{i}'):
                    to_del.append(i)
        for i in sorted(to_del, reverse=True):
            sum_data.pop(i)
        if st.button('+ 添加行', key='as2'):
            sum_data.append({'soil_layer': '素填土', 'point_id': '', 'elevation': '', 'avg_blows': '', 'bearing_capacity': ''})
            st.rerun()

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
                'simple_pile_range': st.session_state.simple_pile_range,
                'simple_foundation_type': st.session_state.simple_foundation_type,
                'simple_soil_layer': st.session_state.simple_soil_layer,
                'foundation_area': st.session_state.foundation_area,
                'testing_standards_page1': st.session_state.testing_standards_page1,
                'testing_standards_item1': st.session_state.testing_standards_item1,
                'testing_standards_chapter3': st.session_state.testing_standards_chapter3,
                'project_units': {
                    'survey': st.session_state.survey_company,
                    'design': st.session_state.design_company,
                    'construction': st.session_state.construction_company,
                    'supervision': st.session_state.supervision_company,
                    'manager': st.session_state.project_manager,
                    'quality_station': st.session_state.quality_station,
                },
                'geo_layers': st.session_state.geo_layers,
                'instruments': st.session_state.instruments,
                'raw_data': st.session_state.raw_data,
                'summary_data': st.session_state.summary_data,
                'suggestion_on': st.session_state.suggestion_on,
                'suggestion_type': st.session_state.suggestion_type,
                'images': st.session_state.appendix_images,
                # 新增字段
                'witness': st.session_state.witness,
                'certificate_no': st.session_state.certificate_no,
                'structure_type': st.session_state.structure_type,
                'base_type': st.session_state.base_type,
                'base_elevation': st.session_state.base_elevation,
                'test_method': st.session_state.test_method,
                'remark': st.session_state.remark,
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