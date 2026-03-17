import os
import json
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

# 初始化 OpenAI 客户端 (需要配置正确的 API Key 和 Base URL)
# 这里假设您使用的是兼容 OpenAI API 的模型 (如 DeepSeek, 豆包，Kimi，或 OpenAI 本身)
client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY", "your-api-key"),
    base_url=os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
)

def parse_requirements(requirements_text):
    """
    使用大模型将自然语言的排版要求解析为结构化的 JSON 配置
    """
    prompt = f"""
    将以下关于Word文档格式排版的要求，转换为标准的JSON格式字典。
    我们将排版维度扩展为以下 7 种类型：
    1. "title" (文档总的大标题)
    2. "heading_1" (一、 1...等一级标题)
    3. "heading_2" (1.1, 2.1...等二级标题)
    4. "heading_3" (1.1.1...等三级标题)
    5. "body" (普通正文)
    6. "caption" (图表标注、图题、表题)
    7. "reference" (参考文献)
    
    对于每种类型，从下文中提取以下属性（如果完全没有提到则给 null）：
    - "font_name_zh": 中文字体名称 (如 "宋体", "黑体", "楷体")
    - "font_name_en": 英文/数字字体名称 (如 "Times New Roman")
    - "size_pt": 字号的实际磅数（必须严格换算：初号=42, 小初=36, 一号=26, 小一=24, 二号=22, 小二=18, 三号=16, 小三号=15, 四号=14, 小四号=12, 五号=10.5, 小五号=9）
    - "bold": 是否加粗 (true/false)
    - "alignment": 对齐方式 ("left", "center", "right", "justify")
    - "line_spacing_multiple": 行距倍数 (如果是例如"1.5倍行距"等相对值，填1.5)
    - "line_spacing_exact_pt": 固定行距磅数 (如果要求是"固定值20磅"等绝对值，填数字20)
    - "space_before_pt": 段前间距磅数 (如要求0.5行，以小四12磅为基准约为 6，则填 6)
    - "space_after_pt": 段后间距磅数 (同上)
    - "first_line_indent": 是否首行缩进 (true/false)

    要求原文：
    {requirements_text}

    只返回干净的 JSON 对象，包含上述7个key（若无要求则配空字典），不要任何解释和Markdown语法：
    """

    try:
        response = client.chat.completions.create(
            model=os.getenv("LLM_MODEL", "gpt-3.5-turbo"),
            messages=[
                {"role": "system", "content": "你是一个专业的学术论文排版格式解析助手，严格输出合法JSON字典。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1
        )
        result = response.choices[0].message.content.strip()
        if result.startswith("```json"):
            result = result[7:-3]
        elif result.startswith("```"):
            result = result[3:-3]
        return json.loads(result)
    except Exception as e:
        print(f"解析格式要求时出错: {e}")
        return None

def classify_paragraphs(paragraphs_data):
    """
    使用大模型分析长文本的每一段属于什么类型。
    paragraphs_data: [{"idx": 0, "text": "..."}, ...]
    """
    prompt_data = []
    for p in paragraphs_data:
        prompt_data.append({"idx": p["idx"], "text": p["text"][:100]})
        
    prompt = f"""
    你是一个专业的学术文档排版结构分析工具。你需要分析以下段落的类型。
    段落类型严格只能是以下七种之一：
    - "title": 文档的大标题
    - "heading_1": 一级标题 (通常为 1. 引言, 或者 一、概述等只有几个字的章节开头)
    - "heading_2": 二级标题 (通常带有 1.1 这种编号的短标题)
    - "heading_3": 三级标题 (通常带有 1.1.1 这种编号的短标题)
    - "body": 普通正文内容 (包含摘要正文。特征是长句子，常常包含标点符号)
    - "caption": 图表标注 (通常极短，如 "图1 示意图" 或 "表1 数据")
    - "reference": 参考文献的内容条目 (常带有 [1] [2] 或者按照作者姓名排列，经常在全文末尾)
    
    请返回一个标准的 JSON 字典，键为段落的 idx（字符串形式），值为段落对应的类型（字符串）。
    除了 JSON 数据外，不要输出任何其他内容。
    
    待分析的数据：
    {json.dumps(prompt_data, ensure_ascii=False)}
    """

    try:
        response = client.chat.completions.create(
            model=os.getenv("LLM_MODEL", "gpt-3.5-turbo"),
            messages=[
                {"role": "system", "content": "你是一个严谨的结构分析器，通过特征（数字编号、句首词、长度等）准确分类段落。只输出合法的JSON格式字典。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1
        )
        result = response.choices[0].message.content.strip()
        if result.startswith("```json"):
            result = result[7:-3]
        elif result.startswith("```"):
            result = result[3:-3]
        return json.loads(result)
    except Exception as e:
        print(f"分析段落类型时出错: {e}")
        return {}

if __name__ == "__main__":
    text = "大标题用二号黑体加粗居中。正文部分宋体小四，1.5倍行距，首行缩进两个字符。"
    print("测试解析结果: ")
    print(json.dumps(parse_requirements(text), indent=4, ensure_ascii=False))
