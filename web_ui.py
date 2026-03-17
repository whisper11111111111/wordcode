import os
import gradio as gr

# 配置环境变量，防止本机的网络代理（如梯子）拦截本地端口导致 503 错误
os.environ["NO_PROXY"] = "localhost,127.0.0.1,::1"

from src.ai_parser import parse_requirements, classify_paragraphs
from src.formatter import process_document, extract_paragraphs_text

DEFAULT_REQUIREMENTS = """【默认学术排版规范】
1. 论文大标题（封面题目）：二号、黑体，居中对齐；
2. 一级标题（如“1 文献综述”、“2 开题报告”）：三号、黑体；
3. 二级标题（如“1.1 发展趋势”）：四号、黑体；
4. 三级标题（如“1.1.1 图题标注”）：小四（12磅）、黑体；
5. 正文内容：小四（12磅）、中文字体为宋体，英文及数字字体为 Times New Roman，首行缩进两字符，1.5倍行距；
6. 封面表格及其他补充信息：四号、宋体（若填报双字姓名中间最好留有空格，如“李  明”）；
7. 目录格式：目录大标题“目  录”黑体二号居中；目录中的一级标题部分宋体小四加粗；二、三级标题部分宋体小四且不加粗。
8. 参考文献：内容为五号字体（中文宋体，英文Times New Roman），单倍行距，左对齐。悬挂缩进。正文引用处需统一处理为上标（在程序内部将自动实现）。"""

def format_word_document(input_file, requirements_text, generate_toc):
    if not input_file:
        return None, "错误：请先上传 Word 文档！"
    
    if not requirements_text or not requirements_text.strip():
        requirements_text = DEFAULT_REQUIREMENTS

    input_path = input_file.name
    # 构造输出路径（在原文档目录下临时生成）
    output_path = input_path.replace(".docx", "_formatted.docx")
    if output_path == input_path:
        output_path = input_path + "_formatted.docx"

    try:
        # 1. 提取结构化规范
        config = parse_requirements(requirements_text)
        if not config:
            return None, "❌ 解析格式要求失败，请检查 API 配置或网络（稍后重试）。"

        # 2. 提取文本并智能分类
        paragraphs_data = extract_paragraphs_text(input_path)
        if paragraphs_data:
            paragraph_types = classify_paragraphs(paragraphs_data)
        else:
            paragraph_types = {}

        # 3. 排版应用并保存
        process_document(input_path, output_path, config, paragraph_types, generate_toc)
        
        return output_path, "✅ 排版成功！请点击上方下载按钮获取文件。"
    
    except Exception as e:
        return None, f"❌ 排版处理时发生错误: {str(e)}"

# 构建 Gradio 网页界面
with gr.Blocks(title="AI Word 智能排版助手") as app:
    gr.Markdown("# 📝 AI Word 智能排版助手")
    gr.Markdown("告别枯燥的排版！上传你的纯文本 Word 文档（.docx），用大白话输入你的排版要求，AI 将自动分析文档结构并为你生成规范的文档。")
    
    with gr.Row():
        with gr.Column():
            file_in = gr.File(label="1. 拖拽或上传原始 Word 文档", file_types=[".docx"])
            req_in = gr.Textbox(
                label="2. 输入排版要求（口语化即可）", 
                lines=8, 
                value=DEFAULT_REQUIREMENTS,
                placeholder="在此输入您的自定义排版要求，如果为空则使用上面的默认规范..."
            )
            toc_checkbox = gr.Checkbox(label="根据一级标题、二级标题自动生成目录（放在文档开头处）", value=True)
            submit_btn = gr.Button("🚀 开始智能排版", variant="primary")
            
        with gr.Column():
            file_out = gr.File(label="🎉 下载排版完成的文档")
            msg_out = gr.Textbox(label="运行状态", interactive=False, lines=2)
            
    # 绑定事件
    submit_btn.click(
        fn=format_word_document, 
        inputs=[file_in, req_in, toc_checkbox],
        outputs=[file_out, msg_out]
    )

if __name__ == "__main__":
    print("正在启动 Web 界面... 请在浏览器中打开返回的链接")
    # 启动服务，inbrowser=True 会在运行后自动在默认浏览器中打开
    app.launch(server_name="127.0.0.1", server_port=7860, inbrowser=True)
