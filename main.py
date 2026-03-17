import argparse
import sys
import os

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.ai_parser import parse_requirements, classify_paragraphs
from src.formatter import process_document, extract_paragraphs_text

def main():
    parser = argparse.ArgumentParser(description="根据要求自动格式化 Word 文档")
    parser.add_argument("input_file", help="要处理的原始 Word 文档路径 (.docx)")
    parser.add_argument("requirements", help="排版格式要求（可以是自然语言文本，也可以是一个包含文本的 .txt 文件路径）")
    parser.add_argument("-o", "--output", help="输出的新文档路径，默认为 '原文件名_formatted.docx'")
    parser.add_argument("--toc", action="store_true", help="是否基于大纲自动生成目录")
    
    args = parser.parse_args()

    # 1. 解析输入参数
    input_path = args.input_file
    if not os.path.exists(input_path):
        print(f"找不到输入的Word文件: {input_path}")
        return

    req_text = args.requirements
    if os.path.exists(req_text):
        with open(req_text, "r", encoding="utf-8") as f:
            req_text = f.read()

    output_path = args.output
    if not output_path:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_formatted{ext}"

    # 2. 调用 AI 解析排版要求
    print("---------------------------------")
    print(f"正在理解排版规范...")
    print(f"自然语言输入: {req_text}")
    
    config = parse_requirements(req_text)
    
    if not config:
        print("解析格式要求失败，请检查 AI API 是否配置正确 (在 .env 文件中)。")
        # 降级：采用默认的规范要求配置 (基于通用学术排版标准)
        config = {
            "title": {"font_name_zh": "黑体", "size_pt": 22.0, "bold": True, "alignment": "center"},
            "heading_1": {"font_name_zh": "黑体", "size_pt": 16.0, "bold": True},
            "heading_2": {"font_name_zh": "黑体", "size_pt": 14.0, "bold": True},
            "heading_3": {"font_name_zh": "黑体", "size_pt": 12.0, "bold": True},
            "body": {
                "font_name_zh": "宋体", 
                "font_name_en": "Times New Roman", 
                "size_pt": 12.0, 
                "line_spacing_multiple": 1.5, 
                "first_line_indent": True
            }
        }
        print("⚠️ 尝试使用默认排版格式...")
    else:
        print(f"解析成功，生成的规范配置为: {config}")
        
    print("---------------------------------")
        
    # 3. 提取 Word 文本并让 AI 进行智能段落分类
    print("正在提取文档文本并进行 AI 智能段落分类...")
    paragraphs_data = extract_paragraphs_text(input_path)
    
    if paragraphs_data:
        paragraph_types = classify_paragraphs(paragraphs_data)
        print("AI 段落分类完成！")
    else:
        paragraph_types = {}
        print("文档由于没有有效文本，跳过分类。")
        
    print("---------------------------------")
        
    # 4. 处理 Word 文档
    print("开始对底层的 Word 文档进行排版引擎应用...")
    try:
        process_document(input_path, output_path, config, paragraph_types, generate_toc=args.toc)
        print("---------------------------------")
    except Exception as e:
        print(f"处理文档时出错: {e}")

if __name__ == "__main__":
    main()
