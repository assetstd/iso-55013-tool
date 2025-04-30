# ISO 55001 审核工具 / ISO 55001 Audit Tool

## 项目简介 / Project Introduction
本项目是基于Streamlit的ISO 55001资产管理体系审核问卷工具，支持中英文双语，灵活的分值权重配置，自动生成PDF/Excel报告，适用于企业自评或第三方审核。

This project is a Streamlit-based ISO 55001 asset management audit questionnaire tool, supporting both Chinese and English, flexible score weighting configuration, and automatic PDF/Excel report generation. Suitable for enterprise self-assessment or third-party audit.

## 主要功能 / Main Features
- 审核问卷支持PJ/XO/PW三种题型
- 分值权重灵活配置，1000分总分制
- 中英文双语切换，界面与报告均支持
- 自动生成美观的PDF/Excel报告，含雷达图
- 进度保存与恢复
- 详细的分数统计与要素分析

## 安装依赖 / Installation
```bash
pip install -r requirements.txt
```

## 运行方法 / How to Run
```bash
streamlit run app.py
```

## 分值权重配置说明 / Score Weight Configuration
- `score_weights.yaml` 文件用于配置各章节、各题目的分值权重。
- 章节权重总分为1000分，题目权重需与题库一一对应。
- 支持自定义权重，修改后无需更改主程序。

## 常见问题 / FAQ
- **Q: 如何切换语言？**
  A: 侧边栏点击"中文/English"按钮即可。
- **Q: 如何保存/加载进度？**
  A: 侧边栏有"保存当前进度/加载上次进度"按钮。
- **Q: 报告导出乱码？**
  A: 请确保`fonts`文件夹下有`simsun.ttc`和`simhei.ttf`字体文件。
- **Q: 权重配置后分数不对？**
  A: 请检查`score_weights.yaml`题目权重与题库题目数量是否一致。

## 联系方式 / Contact
如有问题或建议，请联系：
- 邮箱: your_email@example.com
- Issues: https://github.com/yourrepo/iso-55001-tool/issues

## 系统要求

- Python 3.8 或更高版本
- Windows/Linux/MacOS

## 安装步骤

1. 克隆仓库：
```bash
git clone https://github.com/yourusername/iso-55001-tool.git
cd iso-55001-tool
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

3. 运行应用：
```bash
streamlit run app.py
```

## 使用说明

1. 启动应用后，在浏览器中打开显示的地址
2. 选择界面语言（中文/English）
3. 按章节回答审核问题
4. 使用顶部的保存按钮保存进度
5. 在"结果分析"标签页查看评分情况
6. 在"报告导出"标签页生成审核报告

## 问题类型说明

- PJ（主观判断）：根据符合程度评分（0-4分）
- XO（是否判断）：是=4分，否=0分
- PW（多项选择）：多个子项的平均分

## 评分标准

- 90%符合度要求
- 60%相关人员理解内容和要求
- 执行时间不少于3个月

## 开发团队

[您的团队/组织名称]

## 许可证

[许可证类型]
