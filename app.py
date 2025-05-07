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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether, PageBreak
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
        conn = sqlite3.connect('assessment_data.db')
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
                CREATE TABLE IF NOT EXISTS assessment_results (
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

# 保存评估结果
def save_assessment_results(responses, sub_responses):
    """保存评估结果"""
    try:
        with get_db_connection() as conn:
            c = conn.cursor()
            c.execute('''
                INSERT INTO assessment_results (timestamp, responses, sub_responses)
                VALUES (?, ?, ?)
            ''', (datetime.now().isoformat(), 
                  json.dumps(responses), 
                  json.dumps(sub_responses)))
            conn.commit()
        logging.info("结果保存成功")
    except Exception as e:
        logging.error(f"保存结果失败: {str(e)}")
        logging.error(traceback.format_exc())
        raise

# 加载最近的评估结果
def load_latest_assessment_results():
    """加载最近的评估结果"""
    try:
        with get_db_connection() as conn:
            c = conn.cursor()
            c.execute('''
                SELECT responses, sub_responses 
                FROM assessment_results 
                ORDER BY timestamp DESC 
                LIMIT 1
            ''')
            result = c.fetchone()
            
            if result:
                return json.loads(result[0]), json.loads(result[1])
            return {}, {}
    except Exception as e:
        logging.error(f"加载结果失败: {str(e)}")
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
    if 'language' not in st.session_state:
        st.session_state.language = 'zh'  # 默认中文

# 加载评估问题
def load_questionnaire():
    """加载问题"""
    try:
        # 加载中文问题
        with open('questionnaire.yaml', 'r', encoding='utf-8') as file:
            questions_zh = yaml.safe_load(file)
        
        # 加载英文问题
        with open('questionnaire_en.yaml', 'r', encoding='utf-8') as file:
            questions_en = yaml.safe_load(file)
        
        logging.info("成功加载问题")
        
        # 合并中英文内容
        formatted_questions = {}
        
        # 创建英文键名到中文键名的映射
        section_mapping = {
            'organization_context': '组织环境',
            'leadership': '领导力',
            'planning': '策划',
            'support': '支持',
            'operation': '运行',
            'performance_evaluation': '绩效评价',
            'improvements': '改进'
        }
        
        # 遍历英文问题作为基准
        for section_en, section_data_en in questions_en.items():
            section_zh = section_mapping.get(section_en)
            if not section_zh:
                logging.warning(f"找不到章节 '{section_en}' 的中文映射")
                continue
                
            section_zh_data = questions_zh.get(section_zh, {})
            
            formatted_questions[section_en] = {
                'id': section_en,
                'name': {
                    'zh': section_zh,
                    'en': section_en.replace('_', ' ').title()
                },
                'questions': {}
            }
            
            for q_id, q_data_en in section_data_en.items():
                q_data_zh = section_zh_data.get(q_id, {})
                
                formatted_question = {
                    'type': q_data_en['type'],
                    'description': {
                        'en': q_data_en['description'],
                        'zh': q_data_zh.get('description', q_data_en['description'])  # 如果没有中文描述，使用英文
                    }
                }
                
                # 处理子问题
                if 'sub_questions' in q_data_en:
                    formatted_question['sub_questions'] = {
                        'en': q_data_en['sub_questions'],
                        'zh': q_data_zh.get('sub_questions', q_data_en['sub_questions'])  # 如果没有中文子问题，使用英文
                    }
                
                formatted_questions[section_en]['questions'][q_id] = formatted_question
        
        return formatted_questions
    except Exception as e:
        logging.error(f"加载问题失败: {str(e)}")
        logging.error(traceback.format_exc())
        raise

# 加载分值权重配置
def load_score_weights():
    with open('score_weights.yaml', 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)

score_weights_config = load_score_weights()

# 计算合规分数
def calculate_compliance_score(responses, question_type, sub_responses=None, score_weight=1):
    """计算合规分数"""
    try:
        if question_type == "PW":
            if not sub_responses:
                return 0
            # 直接用True/False算平均
            sub_scores = [100 if v else 0 for v in sub_responses.values()]
            base_score = sum(sub_scores) / len(sub_scores) if sub_scores else 0
        else:
            if not responses:
                return 0
            type_base = score_weights_config['question_type_base_scores']
            if question_type == "XO":
                base_score = type_base['XO'].get("yes" if responses == 4 else "no", 0)
            else:  # PJ类型
                base_score = type_base['PJ'].get(responses, 0)
        weighted_score = (base_score / 100) * score_weight
        return weighted_score
    except Exception as e:
        logging.error(f"计算合规分数失败: {str(e)}")
        logging.error(traceback.format_exc())
        return 0

# 计算总分（1000分制）
def calculate_total_score(section_scores):
    """计算1000分制总分"""
    try:
        if not section_scores:
            return 0
        # 直接返回所有章节得分之和
        return sum(section_scores.values())
    except Exception as e:
        logging.error(f"计算总分失败: {str(e)}")
        logging.error(traceback.format_exc())
        return 0

# 生成雷达图
def create_radar_chart(section_scores):
    """生成雷达图"""
    try:
        if not section_scores:
            return None
        section_weights = score_weights_config['section_weights']
        # 获取章节名称和分数
        categories = []
        values = []
        for section, score in section_scores.items():
            section_name = section.replace('_', ' ').title() if st.session_state.language == 'en' else {
                'organization_context': '组织环境',
                'leadership': '领导力',
                'planning': '策划',
                'support': '支持',
                'operation': '运行',
                'performance_evaluation': '绩效评价',
                'improvements': '改进'
            }.get(section, section)
            categories.append(section_name)
            max_score = section_weights.get(section, 100)
            # 满分显示100，否则按百分比
            value = 100 if abs(score - max_score) < 1e-6 else (score / max_score * 100 if max_score else 0)
            values.append(value)
        fig = go.Figure()
        fig.add_trace(go.Scatterpolar(
            r=values,
            theta=categories,
            fill='toself',
            name='Compliance Score' if st.session_state.language == 'en' else '合规分数',
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

def create_pdf_report(section_scores, questionnaire, responses, sub_responses):
    """生成PDF报告"""
    try:
        if not section_scores or not questionnaire:
            logging.error("生成PDF报告失败：缺少必要数据")
            return None
            
        # 使用系统字体
        font_dir = Path(__file__).parent / "fonts"
        simsun_path = font_dir / "simsun.ttc"
        simhei_path = font_dir / "simhei.ttf"

        try:
            pdfmetrics.registerFont(TTFont('SimSun', str(simsun_path)))
            pdfmetrics.registerFont(TTFont('SimHei', str(simhei_path)))
            main_font = 'SimSun' if st.session_state.language == 'zh' else 'Helvetica'
            bold_font = 'SimHei' if st.session_state.language == 'zh' else 'Helvetica-Bold'
        except Exception as e:
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
        title = "ISO 55013 Assessment Report" if st.session_state.language == 'en' else "ISO 55013 评估报告"
        # 组合第一页内容
        first_page_content = []
        first_page_content.append(Paragraph(title, title_style))
        # 总体评分
        total_score = sum(section_scores.values()) if section_scores else 0
        score_style = ParagraphStyle(
            'ScoreStyle',
            parent=heading2_style,
            fontSize=24,
            textColor=colors.HexColor('#27AE60')
        )
        overall_score = "Overall Score: " if st.session_state.language == 'en' else "总体评分："
        first_page_content.append(Paragraph(f"{overall_score}{total_score:.1f}", score_style))
        first_page_content.append(Spacer(1, 10))
        # 雷达图
        radar_chart = create_radar_chart(section_scores)
        if radar_chart:
            try:
                img_data = radar_chart.to_image(format="png")
                img = Image(io.BytesIO(img_data), width=6.8*inch, height=5.0*inch)  # 宽高
                first_page_content.append(img)
                first_page_content.append(Spacer(1, 10))
            except Exception as e:
                logging.error(f"添加雷达图到PDF失败: {str(e)}")
        # 各要素得分详情表格
        first_page_content.append(Paragraph(
            "Element Scores Detail" if st.session_state.language == 'en' else "各要素得分详情",
            heading2_style
        ))
        first_page_content.append(Spacer(1, 5))
        headers = ['Element', 'Score'] if st.session_state.language == 'en' else ['要素', '得分']
        data = [headers]
        for section, score in section_scores.items():
            section_id = questionnaire[section]['name']['zh'] if st.session_state.language == 'zh' else questionnaire[section].get('id', section)
            data.append([section_id, f"{score:.1f}"])
        col_widths = [doc.width/2.0, doc.width/2.0]
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E4053')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), bold_font),
            ('FONTSIZE', (0, 0), (-1, 0), 14),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F8F9F9')),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), main_font),
            ('FONTSIZE', (0, 1), (-1, -1), 12),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#D5D8DC')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8F9F9')]),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ('RIGHTPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]))
        first_page_content.append(table)
        first_page_content.append(Spacer(1, 10))
        elements.append(KeepTogether(first_page_content))
        elements.append(PageBreak())
        
        # 添加详细评估结果
        elements.append(Paragraph(
            "Detailed Assessment Results" if st.session_state.language == 'en' else "详细评估结果",
            heading2_style
        ))
        elements.append(Spacer(1, 15))
        
        score_labels = {
            'zh': {
                0: "未实施",
                1: "初步实施",
                2: "部分实施",
                3: "大部分实施",
                4: "完全实施"
            },
            'en': {
                0: "Not Implemented",
                1: "Initial Implementation",
                2: "Partial Implementation",
                3: "Mostly Implemented",
                4: "Fully Implemented"
            }
        }
        
        for section, section_data in questionnaire.items():
            section_id = section_data['name']['zh'] if st.session_state.language == 'zh' else section_data.get('id', section)
            elements.append(Paragraph(section_id, heading3_style))
            
            for q_id, question in section_data.get('questions', {}).items():
                key = f"{section}_{q_id}"
                score = responses.get(key, 0)
                question_weight = score_weights_config['question_weights'][section].get(q_id, 1)
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
                question_text = "Question: " if st.session_state.language == 'en' else "问题："
                type_text = "Type: " if st.session_state.language == 'en' else "类型："
                score_text = "Score: " if st.session_state.language == 'en' else "得分："
                weight_text = "Weight: " if st.session_state.language == 'en' else "分值权重："
                
                description = get_translated_text(question["description"], st.session_state.language)
                elements.append(Paragraph(f"{question_text}{description}", question_style))
                elements.append(Paragraph(f"{type_text}{question['type']}", normal_style))
                elements.append(Paragraph(f"{score_text}{score}", score_style))
                elements.append(Paragraph(f"{weight_text}{question_weight}", normal_style))
                
                # 如果是多选题，添加子问题得分
                if question['type'] == "PW" and "sub_questions" in question:
                    sub_questions = question["sub_questions"].get(st.session_state.language, [])
                    sub_scores_text = "Sub-question scores: " if st.session_state.language == 'en' else "子问题得分："
                    elements.append(Paragraph(sub_scores_text, normal_style))
                    sub_count = len(sub_questions)
                    question_weight = score_weights_config['question_weights'][section].get(q_id, 1)
                    for i, sub_q in enumerate(sub_questions, 1):
                        sub_key = f"{key}_sub_{i}"
                        sub_score = st.session_state.sub_responses.get(sub_key, False)
                        yes_text = "Yes" if st.session_state.language == 'en' else "是"
                        no_text = "No" if st.session_state.language == 'en' else "否"
                        # 子问题得分为题目权重/子项数
                        sub_score_value = question_weight / sub_count if sub_score else 0
                        if st.session_state.language == 'en':
                            elements.append(Paragraph(f"- {sub_q}: {yes_text if sub_score else no_text} ({sub_score_value:.1f})", normal_style))
                        else:
                            elements.append(Paragraph(f"- {sub_q}: {yes_text if sub_score else no_text}（{sub_score_value:.1f}）", normal_style))
                
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
    page_title="ISO 55013 评估工具",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded"
)

# 加载外部CSS文件
with open('style.css', encoding='utf-8') as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

def get_translated_text(text_dict, lang='zh'):
    """获取翻译文本"""
    if isinstance(text_dict, str):
        return text_dict
    if isinstance(text_dict, dict):
        return text_dict.get(lang, text_dict.get('zh', ''))
    return ''

def get_section_title(section_data, lang='zh'):
    """获取章节标题"""
    name = get_translated_text(section_data.get('name', section_data.get('id', '')), lang)
    if lang == 'zh':
        return f"**{name}**"
    else:
        return f"**{name.title()}**"

def main():
    """主函数"""
    try:
        # 初始化会话状态
        init_session_state()
        
        # 加载评估问题
        try:
            questionnaire = load_questionnaire()
        except Exception as e:
            st.error(f"加载评估问题时出错: {str(e)}")
            return

        # 添加侧边栏
        with st.sidebar:
            # 简化语言切换为两个按钮
            col1, col2 = st.columns(2)
            with col1:
                if st.button("中文", type="primary" if st.session_state.language == 'zh' else "secondary"):
                    st.session_state.language = 'zh'
                    st.rerun()
            with col2:
                if st.button("English", type="primary" if st.session_state.language == 'en' else "secondary"):
                    st.session_state.language = 'en'
                    st.rerun()
            
            # 添加标题
            st.title("ISO 55013 Assessment Toolkit" if st.session_state.language == 'en' else "ISO 55013 评估工具")
            
            st.markdown("---")
            
            # 添加保存和加载按钮
            col1, col2 = st.columns(2)
            with col1:
                save_text = "Save Progress" if st.session_state.language == 'en' else "保存当前进度"
                if st.button(save_text, key="save_button"):
                    try:
                        save_assessment_results(st.session_state.responses, st.session_state.sub_responses)
                        st.session_state.last_save_time = datetime.now()
                        st.success("Progress saved!" if st.session_state.language == 'en' else "进度已保存！")
                    except Exception as e:
                        st.error(f"Error saving progress: {str(e)}" if st.session_state.language == 'en' else f"保存进度时出错: {str(e)}")
            
            with col2:
                load_text = "Load Progress" if st.session_state.language == 'en' else "加载上次进度"
                if st.button(load_text, key="load_button"):
                    try:
                        responses, sub_responses = load_latest_assessment_results()
                        st.session_state.responses = responses
                        st.session_state.sub_responses = sub_responses
                        st.session_state.force_refresh = True
                        st.success("Last progress loaded!" if st.session_state.language == 'en' else "已加载上次保存的进度！")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error loading progress: {str(e)}" if st.session_state.language == 'en' else f"加载进度时出错: {str(e)}")
            
            # 显示上次保存时间
            last_save_text = "Last saved: " if st.session_state.language == 'en' else "上次保存时间："
            st.markdown(f"{last_save_text}{st.session_state.last_save_time.strftime('%Y-%m-%d %H:%M:%S')}")
            
            st.markdown("---")
            if st.session_state.language == 'en':
                st.markdown("""
                #### Question Types
                - PJ: Subjective Judgement. Questions are scored based on 'professional judgement' to judge the level of compliance in accordance with the scoring principles. Give a score from zero to full based on judgement.
                - XO: Judgemental or not. A question can only be answered with a yes or no answer, with a yes receiving full marks and a no receiving no marks. For any activity to be scored, it should be at least '90 per cent compliant', with 60 per cent of those involved understanding the content and requirements, and with an implementation time of at least three months. Anything else scores zero.
                - PW: Multiple choice. When the question contains several components, a score is given for each component and the total is the final score. For any activity to be awarded a score, it should be at least '90 per cent compliant', with 60 per cent of the relevant people understanding the content and requirements, and implementation time of at least three months. Anything else will be scored zero.
                """)
            else:
                st.markdown("""
                #### 问题类型说明
                - PJ：主观判断。问题的评分基于"专业判断"，依照评分原则判断其符合程度。可基于判断，给出零分至满分。
                - XO：是否判断。问题的回答只有是或者否两种答案，"是"得满分，"否"不得分。任何活动要得分的话，其至少应到达"90%符合"，60%的相关人员理解相关的内容和要求，执行时间不少于三个月。除此之外任何其他情形打零分。
                - PW：多项选择。当问题含有几个组成部分时，可以得到每一部分得分，总和为最终得分。任何活动要得分的话，其至少应到达"90%符合"，60%的相关人员理解相关的内容和要求，执行时间不少于三个月。除此之外任何其他情形打零分。
                """)

        # 创建选项卡
        tab_titles = ["System Assessment", "Result Analysis", "Report Export"] if st.session_state.language == 'en' else ["体系评估", "结果分析", "报告导出"]
        tabs = st.tabs(tab_titles)
        
        # 评估标签页
        with tabs[0]:
            try:
                for section, section_data in questionnaire.items():
                    with st.expander(get_section_title(section_data, st.session_state.language), expanded=True):
                        for q_id, question in section_data.get('questions', {}).items():
                            key = f"{section}_{q_id}"
                            col1, col2 = st.columns([3, 1])
                            with col1:
                                type_class = {
                                    "PJ": "question-type-pj",
                                    "XO": "question-type-xo",
                                    "PW": "question-type-pw"
                                }.get(question["type"], "")
                                
                                description = get_translated_text(question.get('description', {}), st.session_state.language)
                                st.markdown(
                                    f'<span class="question-type {type_class}">{question["type"]}</span>'
                                    f'<span style="font-weight: bold;">{description}</span>',
                                    unsafe_allow_html=True
                                )
                            
                            with col2:
                                if question["type"] == "XO":
                                    # 是否题使用单选框
                                    current_value = st.session_state.responses.get(key, 0)
                                    yes_no_options = {
                                        'zh': {0: "否", 4: "是"},
                                        'en': {0: "No", 4: "Yes"}
                                    }
                                    st.session_state.responses[key] = st.radio(
                                        "Score" if st.session_state.language == 'en' else "评分",
                                        options=[0, 4],
                                        format_func=lambda x: yes_no_options[st.session_state.language][x],
                                        horizontal=True,
                                        key=f"radio_{section}_{q_id}",
                                        label_visibility="collapsed",
                                        index=1 if current_value == 4 else 0
                                    )
                                elif question["type"] == "PJ":
                                    # 主观判断题使用下拉框
                                    current_value = st.session_state.responses.get(key, 0)
                                    score_labels = {
                                        'zh': {
                                            0: "未实施",
                                            1: "初步实施",
                                            2: "部分实施",
                                            3: "大部分实施",
                                            4: "完全实施"
                                        },
                                        'en': {
                                            0: "Not Implemented",
                                            1: "Initial Implementation",
                                            2: "Partial Implementation",
                                            3: "Mostly Implemented",
                                            4: "Fully Implemented"
                                        }
                                    }
                                    # 修正index越界问题
                                    index = 0
                                    try:
                                        index = int(current_value)
                                        if index not in [0, 1, 2, 3, 4]:
                                            index = 0
                                    except Exception:
                                        index = 0
                                    st.session_state.responses[key] = st.selectbox(
                                        "Score" if st.session_state.language == 'en' else "评分",
                                        options=[0, 1, 2, 3, 4],
                                        format_func=lambda x: score_labels[st.session_state.language][x],
                                        key=f"select_{section}_{q_id}",
                                        label_visibility="collapsed",
                                        index=index  # 保证index合法
                                    )
                                else:  # PW类型
                                    # 多选题使用复选框
                                    if "sub_questions" in question:
                                        sub_questions = question.get('sub_questions', {}).get(st.session_state.language, [])
                                        for i, sub_q in enumerate(sub_questions, 1):
                                            sub_key = f"{key}_sub_{i}"
                                            if sub_key not in st.session_state.sub_responses:
                                                st.session_state.sub_responses[sub_key] = False
                                            checked = st.checkbox(
                                                sub_q,
                                                value=st.session_state.sub_responses.get(sub_key, False),
                                                key=f"checkbox_{section}_{q_id}_{i}_sub"
                                            )
                                            st.session_state.sub_responses[sub_key] = checked
                                        # 不再赋值st.session_state.responses[key]
                
                # 自动保存功能
                current_time = datetime.now()
                if (current_time - st.session_state.last_save_time).total_seconds() > 300:  # 每5分钟自动保存一次
                    try:
                        save_assessment_results(st.session_state.responses, st.session_state.sub_responses)
                        st.session_state.last_save_time = current_time
                        auto_save_text = "Progress auto-saved" if st.session_state.language == 'en' else "进度已自动保存"
                        st.toast(auto_save_text, icon="💾")
                    except Exception as e:
                        logging.error(f"自动保存失败: {str(e)}")
            
            except Exception as e:
                st.error(f"渲染评估页面时出错: {str(e)}")
                logging.error(f"渲染评估页面失败: {str(e)}")
                logging.error(traceback.format_exc())

        # 结果分析标签页
        with tabs[1]:
            try:
                # 计算各部分得分
                section_scores = {}
                for section in questionnaire.keys():
                    section_responses = {k: v for k, v in st.session_state.responses.items() if k.startswith(section)}
                    section_sub_responses = {k: v for k, v in st.session_state.sub_responses.items() if k.startswith(section)}
                    
                    # 计算每个问题的得分
                    question_scores = []
                    questions = questionnaire[section].get('questions', {})
                    for q_id, question in questions.items():
                        key = f"{section}_{q_id}"
                        if question["type"] == "PW":
                            sub_keys = [k for k in section_sub_responses if k.startswith(key)]
                            sub_count = len(sub_keys)
                            weight = score_weights_config['question_weights'][section].get(q_id, 1)
                            selected_count = sum(1 for k in sub_keys if section_sub_responses[k])
                            score = (weight / sub_count) * selected_count if sub_count else 0
                        elif key in section_responses:
                            score = calculate_compliance_score(
                                section_responses[key],
                                question["type"],
                                None,
                                score_weight=score_weights_config['question_weights'][section].get(q_id, 1)
                            )
                        else:
                            score = 0
                        question_scores.append(score)
                    
                    # 计算要素总分（由平均分改为总和）
                    section_scores[section] = sum(question_scores) if question_scores else 0
                
                # 显示总体合规分数
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    # 只显示1000分制得分
                    total_score = calculate_total_score(section_scores)
                    st.metric(
                        "总体评分" if st.session_state.language == 'zh' else "Overall Score",
                        f"{total_score:.1f}/1000"
                    )
                
                # 显示雷达图
                radar_chart = create_radar_chart(section_scores)
                if radar_chart:
                    st.plotly_chart(radar_chart, use_container_width=True)
                else:
                    st.warning("Cannot generate radar chart" if st.session_state.language == 'en' else "无法生成雷达图")
                # 显示详细得分（去除百分比进度条，仅保留分数）
                st.subheader("Element Scores" if st.session_state.language == 'en' else "要素得分")
                cols = st.columns(3)
                for i, (section, score) in enumerate(section_scores.items()):
                    with cols[i % 3]:
                        section_name = questionnaire[section]['name'][st.session_state.language]
                        st.metric(section_name, f"{score:.1f}")
            
            except Exception as e:
                st.error(f"渲染结果分析页面时出错: {str(e)}")
                logging.error(f"渲染结果分析页面失败: {str(e)}")
                logging.error(traceback.format_exc())

        # 报告导出标签页
        with tabs[2]:
            try:
                col1, col2 = st.columns(2)
                with col1:
                    excel_btn_text = "Generate Excel Report" if st.session_state.language == 'en' else "生成Excel报告"
                    if st.button(excel_btn_text, key="generate_excel_report"):
                        with st.spinner("Generating Excel report..." if st.session_state.language == 'en' else "正在生成Excel报告..."):
                            # 创建报告数据
                            report_data = []
                            for section, section_data in questionnaire.items():
                                section_id = section_data.get('id', section)
                                for q_id, question in section_data.get('questions', {}).items():
                                    key = f"{section}_{q_id}"
                                    score = st.session_state.responses.get(key, 0)
                                    
                                    # 获取子问题得分（如果是多选题）
                                    sub_scores = []
                                    if question["type"] == "PW" and "sub_questions" in question:
                                        sub_questions = question["sub_questions"].get(st.session_state.language, [])
                                        for i, sub_q in enumerate(sub_questions, 1):
                                            sub_key = f"{key}_sub_{i}"
                                            sub_score = st.session_state.sub_responses.get(sub_key, 0)
                                            sub_scores.append({
                                                "sub_question": sub_q,
                                                "score": "Yes" if sub_score == 4 else "No" if st.session_state.language == 'en' else "是" if sub_score == 4 else "否"
                                            })
                                    
                                    score_labels = {
                                        'zh': {
                                            0: "未实施",
                                            1: "初步实施",
                                            2: "部分实施",
                                            3: "大部分实施",
                                            4: "完全实施"
                                        },
                                        'en': {
                                            0: "Not Implemented",
                                            1: "Initial Implementation",
                                            2: "Partial Implementation",
                                            3: "Mostly Implemented",
                                            4: "Fully Implemented"
                                        }
                                    }

                                    report_data.append({
                                        "element": section_id,
                                        "question_type": question["type"],
                                        "question": get_translated_text(question["description"], st.session_state.language),
                                        "score": score,
                                        "assessment": score_labels[st.session_state.language][round(score)] if question["type"] != "XO" else 
                                                    ("Yes" if score == 4 else "No" if st.session_state.language == 'en' else "是" if score == 4 else "否"),
                                        "weight": score_weights_config['question_weights'][section].get(q_id, 1),
                                        "sub_scores": sub_scores if sub_scores else None
                                    })
                            
                            # 创建DataFrame并导出为Excel
                            df = pd.DataFrame(report_data)
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            filename = f"ISO55013_{'Assessment_Report' if st.session_state.language == 'en' else '评估报告'}_{timestamp}.xlsx"
                            
                            try:
                                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                                    sheet_name = 'Assessment Results' if st.session_state.language == 'en' else '评估结果'
                                    df.to_excel(writer, index=False, sheet_name=sheet_name)
                                    
                                    # 创建雷达图数据工作表
                                    radar_data = pd.DataFrame({
                                        'Element' if st.session_state.language == 'en' else '要素': list(section_scores.keys()),
                                        'Score' if st.session_state.language == 'en' else '得分': [f"{v:.1f}" for v in section_scores.values()]
                                    })
                                    radar_sheet_name = 'Radar Chart Data' if st.session_state.language == 'en' else '雷达图数据'
                                    radar_data.to_excel(writer, index=False, sheet_name=radar_sheet_name)
                                
                                # 提供下载链接
                                with open(filename, 'rb') as f:
                                    download_label = "Download Excel Report" if st.session_state.language == 'en' else "下载Excel报告"
                                    st.download_button(
                                        label=download_label,
                                        data=f,
                                        file_name=filename,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                success_msg = "Excel report generated successfully" if st.session_state.language == 'en' else "Excel报告生成成功"
                                st.success(success_msg)
                            except Exception as e:
                                error_msg = f"Error generating Excel report: {str(e)}" if st.session_state.language == 'en' else f"生成Excel报告时出错: {str(e)}"
                                st.error(error_msg)
                                logging.error(error_msg)
                                logging.error(traceback.format_exc())
                
                with col2:
                    pdf_btn_text = "Generate PDF Report" if st.session_state.language == 'en' else "生成PDF报告"
                    if st.button(pdf_btn_text, key="generate_pdf_report"):
                        spinner_text = "Generating PDF report..." if st.session_state.language == 'en' else "正在生成PDF报告..."
                        with st.spinner(spinner_text):
                            pdf_buffer = create_pdf_report(section_scores, questionnaire, st.session_state.responses, st.session_state.sub_responses)
                            if pdf_buffer:
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                filename = f"ISO55013_{'Assessment_Report' if st.session_state.language == 'en' else '评估报告'}_{timestamp}.pdf"
                                download_label = "Download PDF Report" if st.session_state.language == 'en' else "下载PDF报告"
                                st.download_button(
                                    label=download_label,
                                    data=pdf_buffer,
                                    file_name=filename,
                                    mime="application/pdf"
                                )
                                success_msg = "PDF report generated successfully" if st.session_state.language == 'en' else "PDF报告生成成功"
                                st.success(success_msg)
                            else:
                                error_msg = "Failed to generate PDF report" if st.session_state.language == 'en' else "生成PDF报告失败"
                                st.error(error_msg)
            
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