import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import yaml
from datetime import datetime
import sqlite3
import json
from pathlib import Path
import logging
import traceback
from contextlib import contextmanager
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import base64

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# 数据库上下文管理器
@contextmanager
def get_db_connection():
    conn = None
    try:
        conn = sqlite3.connect('audit_data.db')
        yield conn
    except Exception as e:
        logging.error(f"数据库连接错误: {str(e)}")
        logging.error(traceback.format_exc())
        raise
    finally:
        if conn:
            conn.close()

# 初始化数据库
def init_db():
    """初始化数据库"""
    try:
        with get_db_connection() as conn:
            c = conn.cursor()
            c.execute('''
                CREATE TABLE IF NOT EXISTS audit_results (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT,
                    responses TEXT,
                    sub_responses TEXT
                )
            ''')
            conn.commit()
        logging.info("数据库初始化成功")
    except Exception as e:
        logging.error(f"数据库初始化失败: {str(e)}")
        logging.error(traceback.format_exc())
        raise

# 保存审核结果
def save_audit_results(responses, sub_responses):
    """保存审核结果"""
    try:
        with get_db_connection() as conn:
            c = conn.cursor()
            c.execute('''
                INSERT INTO audit_results (timestamp, responses, sub_responses)
                VALUES (?, ?, ?)
            ''', (datetime.now().isoformat(), 
                  json.dumps(responses), 
                  json.dumps(sub_responses)))
            conn.commit()
        logging.info("审核结果保存成功")
    except Exception as e:
        logging.error(f"保存审核结果失败: {str(e)}")
        logging.error(traceback.format_exc())
        raise

# 加载最近的审核结果
def load_latest_audit_results():
    """加载最近的审核结果"""
    try:
        with get_db_connection() as conn:
            c = conn.cursor()
            c.execute('''
                SELECT responses, sub_responses 
                FROM audit_results 
                ORDER BY timestamp DESC 
                LIMIT 1
            ''')
            result = c.fetchone()
            
            if result:
                return json.loads(result[0]), json.loads(result[1])
            return {}, {}
    except Exception as e:
        logging.error(f"加载审核结果失败: {str(e)}")
        logging.error(traceback.format_exc())
        return {}, {}

# 初始化会话状态
def init_session_state():
    """初始化会话状态"""
    if 'responses' not in st.session_state:
        st.session_state.responses = {}
    if 'sub_responses' not in st.session_state:
        st.session_state.sub_responses = {}
    if 'last_save_time' not in st.session_state:
        st.session_state.last_save_time = datetime.now()
    if 'force_refresh' not in st.session_state:
        st.session_state.force_refresh = False

# 加载审核问题
def load_audit_questions():
    """加载审核问题"""
    try:
        with open('audit_questions.yaml', 'r', encoding='utf-8') as file:
            questions = yaml.safe_load(file)
            logging.info("成功加载审核问题")
            return questions
    except Exception as e:
        logging.error(f"加载审核问题失败: {str(e)}")
        logging.error(traceback.format_exc())
        raise

# 计算合规分数
def calculate_compliance_score(responses, question_type, sub_responses=None):
    """计算合规分数"""
    try:
        if not responses:
            return 0
        
        if question_type == "XO":
            return 100 if responses == 4 else 0
        elif question_type == "PW":
            if not sub_responses:
                return 0
            sub_scores = [calculate_compliance_score(r, "XO") for r in sub_responses.values()]
            return sum(sub_scores) / len(sub_scores) if sub_scores else 0
        else:  # PJ类型
            return (responses / 4) * 100
    except Exception as e:
        logging.error(f"计算合规分数失败: {str(e)}")
        logging.error(traceback.format_exc())
        return 0

# 生成雷达图
def create_radar_chart(section_scores):
    """生成雷达图"""
    try:
        if not section_scores:
            return None
            
        categories = list(section_scores.keys())
        values = list(section_scores.values())
        
        fig = go.Figure()
        fig.add_trace(go.Scatterpolar(
            r=values,
            theta=categories,
            fill='toself',
            name='合规分数',
            line_color='#4CAF50'
        ))
        
        fig.update_layout(
            polar=dict(
                radialaxis=dict(
                    visible=True,
                    range=[0, 100],
                    tickfont=dict(size=12),
                    gridcolor='#f0f2f6'
                ),
                angularaxis=dict(
                    tickfont=dict(size=12),
                    gridcolor='#f0f2f6'
                ),
                bgcolor='white'
            ),
            showlegend=False,
            paper_bgcolor='white',
            plot_bgcolor='white',
            margin=dict(t=30, b=30, l=30, r=30)
        )
        return fig
    except Exception as e:
        logging.error(f"生成雷达图失败: {str(e)}")
        logging.error(traceback.format_exc())
        return None

def create_pdf_report(section_scores, audit_questions, responses, sub_responses):
    """生成PDF报告"""
    try:
        if not section_scores or not audit_questions:
            logging.error("生成PDF报告失败：缺少必要数据")
            return None
            
        # 使用系统字体
        font_dir = Path(__file__).parent / "fonts"
        simsun_path = font_dir / "simsun.ttc"
        simhei_path = font_dir / "simhei.ttf"

        try:
            pdfmetrics.registerFont(TTFont('SimSun', str(simsun_path)))
            pdfmetrics.registerFont(TTFont('SimHei', str(simhei_path)))
            main_font = 'SimSun'
            bold_font = 'SimHei'
        except Exception as e:
            # 字体注册失败，降级为系统默认字体
            main_font = 'Helvetica'
            bold_font = 'Helvetica-Bold'
        
        # 创建PDF文档
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, 
                              leftMargin=50,
                              rightMargin=50,
                              topMargin=50,
                              bottomMargin=50)
        styles = getSampleStyleSheet()
        
        # 创建自定义样式
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontName=bold_font,
            fontSize=28,
            spaceAfter=30,
            alignment=1,
            textColor=colors.HexColor('#2E4053')
        )
        
        heading2_style = ParagraphStyle(
            'CustomHeading2',
            parent=styles['Heading2'],
            fontName=bold_font,
            fontSize=20,
            spaceAfter=15,
            textColor=colors.HexColor('#2874A6')
        )
        
        heading3_style = ParagraphStyle(
            'CustomHeading3',
            parent=styles['Heading3'],
            fontName=bold_font,
            fontSize=16,
            spaceAfter=12,
            textColor=colors.HexColor('#3498DB')
        )
        
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontName=main_font,
            fontSize=12,
            spaceAfter=8,
            leading=16,
            textColor=colors.black
        )
        
        elements = []
        
        # 添加标题
        elements.append(Paragraph("ISO 55001 审核报告", title_style))
        elements.append(Paragraph(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", normal_style))
        elements.append(Spacer(1, 30))
        
        # 添加总体得分
        total_score = sum(section_scores.values()) / len(section_scores) if section_scores else 0
        score_style = ParagraphStyle(
            'ScoreStyle',
            parent=heading2_style,
            fontSize=24,
            textColor=colors.HexColor('#27AE60')
        )
        elements.append(Paragraph(f"总体合规分数：{total_score:.1f}%", score_style))
        elements.append(Spacer(1, 30))
        
        # 添加雷达图
        radar_chart = create_radar_chart(section_scores)
        if radar_chart:
            try:
                img_data = radar_chart.to_image(format="png")
                img = Image(io.BytesIO(img_data), width=6*inch, height=4*inch)
                elements.append(img)
                elements.append(Spacer(1, 30))
            except Exception as e:
                logging.error(f"添加雷达图到PDF失败: {str(e)}")
        
        # 添加各要素得分
        elements.append(Paragraph("各要素得分详情", heading2_style))
        elements.append(Spacer(1, 15))
        
        # 创建得分表格
        data = [['要素', '得分']]
        for section, score in section_scores.items():
            data.append([section, f"{score:.1f}%"])
        
        # 计算表格宽度
        col_widths = [doc.width/2.0, doc.width/2.0]
        
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E4053')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'SimHei'),
            ('FONTSIZE', (0, 0), (-1, 0), 14),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F8F9F9')),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'SimSun'),
            ('FONTSIZE', (0, 1), (-1, -1), 12),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#D5D8DC')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8F9F9')]),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 30))
        
        # 添加详细评估结果
        elements.append(Paragraph("详细评估结果", heading2_style))
        elements.append(Spacer(1, 15))
        
        for section, questions in audit_questions.items():
            elements.append(Paragraph(section, heading3_style))
            for q_id, question in questions.items():
                key = f"{section}_{q_id}"
                score = responses.get(key, 0)
                
                # 创建问题样式
                question_style = ParagraphStyle(
                    'QuestionStyle',
                    parent=normal_style,
                    fontSize=13,
                    textColor=colors.HexColor('#2C3E50'),
                    spaceAfter=5
                )
                
                # 创建得分样式
                score_style = ParagraphStyle(
                    'ScoreStyle',
                    parent=normal_style,
                    fontSize=12,
                    textColor=colors.HexColor('#E74C3C'),
                    spaceAfter=8
                )
                
                # 添加问题描述和得分
                elements.append(Paragraph(f"问题：{question['description']}", question_style))
                elements.append(Paragraph(f"类型：{question['type']}", normal_style))
                elements.append(Paragraph(f"得分：{score}", score_style))
                
                # 如果是多选题，添加子问题得分
                if question['type'] == "PW" and "sub_questions" in question:
                    elements.append(Paragraph("子问题得分：", normal_style))
                    for i, sub_q in enumerate(question['sub_questions'], 1):
                        sub_key = f"{key}_sub_{i}"
                        sub_score = sub_responses.get(sub_key, 0)
                        elements.append(Paragraph(f"- {sub_q}: {'是' if sub_score == 4 else '否'}", normal_style))
                
                elements.append(Spacer(1, 15))
        
        # 生成PDF
        doc.build(elements)
        buffer.seek(0)
        return buffer
    except Exception as e:
        logging.error(f"生成PDF报告失败: {str(e)}")
        logging.error(traceback.format_exc())
        return None

# 初始化数据库
init_db()

# 设置页面配置
st.set_page_config(
    page_title="ISO 55001 审核工具",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded"
)

# 加载外部CSS文件
with open('style.css') as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

def main():
    """主函数"""
    try:
        # 初始化会话状态
        init_session_state()
        
        # 加载审核问题
        try:
            audit_questions = load_audit_questions()
        except Exception as e:
            st.error(f"加载审核问题时出错: {str(e)}")
            return

        # 添加侧边栏
        with st.sidebar:
            st.title("ISO 55001 审核工具")
            st.markdown("---")
            
            # 添加保存和加载按钮
            col1, col2 = st.columns(2)
            with col1:
                if st.button("保存当前进度", key="save_button"):
                    try:
                        save_audit_results(st.session_state.responses, st.session_state.sub_responses)
                        st.session_state.last_save_time = datetime.now()
                        st.success("进度已保存！")
                    except Exception as e:
                        st.error(f"保存进度时出错: {str(e)}")
            
            with col2:
                if st.button("加载上次进度", key="load_button"):
                    try:
                        responses, sub_responses = load_latest_audit_results()
                        st.session_state.responses = responses
                        st.session_state.sub_responses = sub_responses
                        st.session_state.force_refresh = True
                        st.success("已加载上次保存的进度！")
                        st.rerun()
                    except Exception as e:
                        st.error(f"加载进度时出错: {str(e)}")
            
            # 显示上次保存时间
            st.markdown(f"上次保存时间：{st.session_state.last_save_time.strftime('%Y-%m-%d %H:%M:%S')}")
            
            st.markdown("---")
            st.markdown("""
            #### 问题类型说明
            - PJ：主观判断。问题的评分基于"专业判断"，审核员须依照评分原则判断其符合程度。审核员可基于判断，给出零分至满分。
            - XO：是否判断。问题的回答只有是或者否两种答案，"是"得满分，"否"不得分。任何活动要得分的话，其至少应到达"90%符合"，60%的相关人员理解相关的内容和要求，执行时间不少于三个月。除此之外任何其他情形打零分。
            - PW：多项选择。当问题含有几个组成部分时，可以得到每一部分得分，总和为最终得分。任何活动要得分的话，其至少应到达"90%符合"，60%的相关人员理解相关的内容和要求，执行时间不少于三个月。除此之外任何其他情形打零分。
            """)

        # 创建选项卡
        tabs = st.tabs(["审核评估", "结果分析", "报告导出"])
        
        # 审核评估标签页
        with tabs[0]:
            try:
                section_titles = {
                    "组织环境": "**组织环境（Context of the organization）**",
                    "领导力": "**领导力（Leadership）**",
                    "策划": "**策划（Planning）**",
                    "支持": "**支持（Support）**",
                    "运行": "**运行（Operation）**",
                    "绩效评价": "**绩效评价（Performance evaluation）**",
                    "改进": "**改进（Improvement）**"
                }
                
                for section, questions in audit_questions.items():
                    with st.expander(section_titles[section], expanded=True):
                        for q_id, question in questions.items():
                            key = f"{section}_{q_id}"
                            col1, col2 = st.columns([3, 1])
                            with col1:
                                type_class = {
                                    "PJ": "question-type-pj",
                                    "XO": "question-type-xo",
                                    "PW": "question-type-pw"
                                }.get(question["type"], "")
                                st.markdown(
                                    f'<span class="question-type {type_class}">{question["type"]}</span>'
                                    f'<span style="font-weight: bold;">{question["description"]}</span>',
                                    unsafe_allow_html=True
                                )
                            
                            with col2:
                                if question["type"] == "XO":
                                    # 是否题使用单选框
                                    current_value = st.session_state.responses.get(key, 0)
                                    st.session_state.responses[key] = st.radio(
                                        "评分",
                                        options=[0, 4],
                                        format_func=lambda x: "是" if x == 4 else "否",
                                        horizontal=True,
                                        key=f"radio_{section}_{q_id}",
                                        label_visibility="collapsed",
                                        index=1 if current_value == 4 else 0
                                    )
                                elif question["type"] == "PJ":
                                    # 主观判断题使用下拉框
                                    current_value = st.session_state.responses.get(key, 0)
                                    st.session_state.responses[key] = st.selectbox(
                                        "评分",
                                        options=[0, 1, 2, 3, 4],
                                        format_func=lambda x: {
                                            0: "未实施",
                                            1: "初步实施",
                                            2: "部分实施",
                                            3: "大部分实施",
                                            4: "完全实施"
                                        }[x],
                                        key=f"select_{section}_{q_id}",
                                        label_visibility="collapsed",
                                        index=current_value
                                    )
                                else:  # PW类型
                                    # 多选题使用复选框
                                    if "sub_questions" in question:
                                        sub_scores = []
                                        for i, sub_q in enumerate(question["sub_questions"], 1):
                                            sub_key = f"{key}_sub_{i}"
                                            if sub_key not in st.session_state.sub_responses:
                                                st.session_state.sub_responses[sub_key] = False
                                            current_value = st.session_state.sub_responses.get(sub_key, False)
                                            checked = st.checkbox(
                                                sub_q,
                                                value=current_value,
                                                key=f"checkbox_{section}_{q_id}_{i}_sub"
                                            )
                                            st.session_state.sub_responses[sub_key] = checked
                                            sub_scores.append(4 if checked else 0)
                                        st.session_state.responses[key] = sum(sub_scores) / len(sub_scores) if sub_scores else 0
                
                # 自动保存功能
                current_time = datetime.now()
                if (current_time - st.session_state.last_save_time).total_seconds() > 300:  # 每5分钟自动保存一次
                    try:
                        save_audit_results(st.session_state.responses, st.session_state.sub_responses)
                        st.session_state.last_save_time = current_time
                        st.toast("进度已自动保存", icon="💾")
                    except Exception as e:
                        logging.error(f"自动保存失败: {str(e)}")
            
            except Exception as e:
                st.error(f"渲染审核评估页面时出错: {str(e)}")
                logging.error(f"渲染审核评估页面失败: {str(e)}")
                logging.error(traceback.format_exc())

        # 结果分析标签页
        with tabs[1]:
            try:
                # 计算各部分得分
                section_scores = {}
                for section in audit_questions.keys():
                    section_responses = {k: v for k, v in st.session_state.responses.items() if k.startswith(section)}
                    section_sub_responses = {k: v for k, v in st.session_state.sub_responses.items() if k.startswith(section)}
                    
                    # 计算每个问题的得分
                    question_scores = []
                    for q_id, question in audit_questions[section].items():
                        key = f"{section}_{q_id}"
                        if key in section_responses:
                            score = calculate_compliance_score(
                                section_responses[key],
                                question["type"],
                                {k: v for k, v in section_sub_responses.items() if k.startswith(key)}
                            )
                            question_scores.append(score)
                    
                    # 计算要素平均分
                    section_scores[section] = sum(question_scores) / len(question_scores) if question_scores else 0
                
                # 显示总体合规分数
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    total_score = sum(section_scores.values()) / len(section_scores) if section_scores else 0
                    st.metric("审核量化打分", f"{total_score:.1f}%")
                
                # 显示雷达图
                radar_chart = create_radar_chart(section_scores)
                if radar_chart:
                    st.plotly_chart(radar_chart, use_container_width=True)
                else:
                    st.warning("无法生成雷达图")
                
                # 显示详细得分
                st.subheader("要素得分")
                cols = st.columns(3)
                for i, (section, score) in enumerate(section_scores.items()):
                    with cols[i % 3]:
                        st.metric(section, f"{score:.1f}%")
                        st.progress(score / 100)
            
            except Exception as e:
                st.error(f"渲染结果分析页面时出错: {str(e)}")
                logging.error(f"渲染结果分析页面失败: {str(e)}")
                logging.error(traceback.format_exc())

        # 报告导出标签页
        with tabs[2]:
            try:
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("生成Excel报告", key="generate_excel_report"):
                        with st.spinner("正在生成Excel报告..."):
                            # 创建报告数据
                            report_data = []
                            for section, questions in audit_questions.items():
                                for q_id, question in questions.items():
                                    key = f"{section}_{q_id}"
                                    score = st.session_state.responses.get(key, 0)
                                    
                                    # 获取子问题得分（如果是多选题）
                                    sub_scores = []
                                    if question["type"] == "PW" and "sub_questions" in question:
                                        for i, sub_q in enumerate(question["sub_questions"], 1):
                                            sub_key = f"{key}_sub_{i}"
                                            sub_score = st.session_state.sub_responses.get(sub_key, 0)
                                            sub_scores.append({
                                                "子问题": sub_q,
                                                "得分": "是" if sub_score == 4 else "否"
                                            })
                                    
                                    report_data.append({
                                        "要素": section,
                                        "问题类型": question["type"],
                                        "问题": question["description"],
                                        "得分": score,
                                        "评估结果": {
                                            0: "未实施",
                                            1: "初步实施",
                                            2: "部分实施",
                                            3: "大部分实施",
                                            4: "完全实施"
                                        }[round(score)] if question["type"] != "XO" else ("是" if score == 4 else "否"),
                                        "子问题得分": sub_scores if sub_scores else None
                                    })
                            
                            # 创建DataFrame并导出为Excel
                            df = pd.DataFrame(report_data)
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            filename = f"ISO55001_审核报告_{timestamp}.xlsx"
                            
                            try:
                                # 使用ExcelWriter来创建Excel文件
                                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                                    # 写入审核结果数据
                                    df.to_excel(writer, index=False, sheet_name='审核结果')
                                    
                                    # 创建雷达图数据工作表
                                    radar_data = pd.DataFrame({
                                        '要素': list(section_scores.keys()),
                                        '得分': list(section_scores.values())
                                    })
                                    radar_data.to_excel(writer, index=False, sheet_name='雷达图数据')
                                    
                                    # 获取工作簿和工作表
                                    workbook = writer.book
                                    worksheet = writer.sheets['雷达图数据']
                                    
                                    # 创建雷达图
                                    from openpyxl.chart import RadarChart, Reference
                                    
                                    # 创建雷达图对象
                                    chart = RadarChart()
                                    chart.style = 2
                                    chart.title = "要素得分雷达图"
                                    
                                    # 设置数据范围
                                    data = Reference(worksheet, min_col=2, min_row=1, max_row=len(radar_data) + 1)
                                    cats = Reference(worksheet, min_col=1, min_row=2, max_row=len(radar_data) + 1)
                                    
                                    # 添加数据到图表
                                    chart.add_data(data, titles_from_data=True)
                                    chart.set_categories(cats)
                                    
                                    # 设置图表大小
                                    chart.height = 15
                                    chart.width = 20
                                    
                                    # 将图表添加到工作表
                                    worksheet.add_chart(chart, "D2")
                                
                                # 提供下载链接
                                with open(filename, 'rb') as f:
                                    st.download_button(
                                        label="下载Excel报告",
                                        data=f,
                                        file_name=filename,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                st.success("Excel报告生成成功")
                            except Exception as e:
                                st.error(f"生成Excel报告时出错: {str(e)}")
                                logging.error(f"生成Excel报告失败: {str(e)}")
                                logging.error(traceback.format_exc())
                
                with col2:
                    if st.button("生成PDF报告", key="generate_pdf_report"):
                        with st.spinner("正在生成PDF报告..."):
                            pdf_buffer = create_pdf_report(section_scores, audit_questions, st.session_state.responses, st.session_state.sub_responses)
                            if pdf_buffer:
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                filename = f"ISO55001_审核报告_{timestamp}.pdf"
                                st.download_button(
                                    label="下载PDF报告",
                                    data=pdf_buffer,
                                    file_name=filename,
                                    mime="application/pdf"
                                )
                                st.success("PDF报告生成成功")
                            else:
                                st.error("生成PDF报告失败")
            
            except Exception as e:
                st.error(f"生成报告时出错: {str(e)}")
                logging.error(f"生成报告失败: {str(e)}")
                logging.error(traceback.format_exc())
    
    except Exception as e:
        st.error(f"应用运行出错: {str(e)}")
        logging.error(f"应用运行失败: {str(e)}")
        logging.error(traceback.format_exc())

if __name__ == "__main__":
    main() 