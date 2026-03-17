#  AI Word 智能排版助手

本工具利用大语言模型（LLM）根据自然语言指令自动排版 Word 文档。

## 功能特性
* 自然语言排版
* 全自动目录生成（含真实页码和域规则）
* 智能学术级页码隔离及分节

## 快速上手
1. pip install -r requirements.txt
2. 复制 .env.example 为 .env 并填写 OPENAI_API_KEY。
3. 运行 python web_ui.py 启动可视化界面。
