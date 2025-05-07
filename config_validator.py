import yaml
import os

def validate_lang_config(lang_config, required_keys, lang_name):
    missing = [k for k in required_keys if k not in lang_config]
    if missing:
        raise ValueError(f"{lang_name} 缺少以下多语言键: {missing}")

def validate_score_weights(score_weights):
    # 检查主要结构
    for key in ['section_weights', 'question_weights', 'question_type_base_scores']:
        if key not in score_weights:
            raise ValueError(f"score_weights.yaml 缺少关键字段: {key}")
    # 可进一步校验每个 section/question 的结构和类型

if __name__ == "__main__":
    CONFIG_DIR = 'config'
    # 加载配置
    def load_yaml(filename):
        path = os.path.join(CONFIG_DIR, filename)
        with open(path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)

    lang_zh = load_yaml('lang_zh.yaml')
    lang_en = load_yaml('lang_en.yaml')
    score_weights = load_yaml('score_weights.yaml')

    lang_required_keys = [
        'app_title', 'section_scores', 'total_score', 'save_progress', 'load_progress',
        'progress_saved', 'progress_loaded', 'report_title', 'report_generate_time',
        'radar_chart', 'section_table', 'assessment_detail', 'export_pdf', 'export_excel',
        'zh_button', 'en_button', 'last_saved', 'system_assessment', 'result_analysis', 'report_export',
        'generate_excel_report', 'generating_excel_report', 'download_excel_report', 'excel_report_generated',
        'error_generating_excel_report', 'radar_chart_data', 'assessment_results',
        'generate_pdf_report', 'generating_pdf_report', 'download_pdf_report', 'pdf_report_generated',
        'failed_generating_pdf_report', 'overall_score', 'element_scores', 'cannot_generate_radar_chart',
        'score', 'question', 'type', 'weight', 'sub_question_scores', 'answer_yes', 'answer_no',
        'progress_auto_saved', 'last_progress_loaded', 'error_saving_progress', 'error_loading_progress',
        'element', 'element_scores_detail', 'detailed_assessment_results'
    ]
    try:
        validate_lang_config(lang_zh, lang_required_keys, "lang_zh.yaml")
        print("lang_zh.yaml 校验通过！")
        validate_lang_config(lang_en, lang_required_keys, "lang_en.yaml")
        print("lang_en.yaml 校验通过！")
        validate_score_weights(score_weights)
        print("score_weights.yaml 校验通过！")
    except Exception as e:
        print(f"配置校验失败: {e}") 