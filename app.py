import streamlit as st
import os
from docx import Document
from docx.shared import Pt
import json
import time
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import JsonOutputParser
from io import BytesIO

# --- Configuration ---
st.set_page_config(page_title="AI 智能教案生成器 (V2 Pro)", layout="wide", initial_sidebar_state="expanded")

# --- UI Components: Console Logger ---
class ConsoleLogger:
    def __init__(self):
        self.container = st.empty()
        self.logs = []

    def log(self, message, icon="🤖"):
        timestamp = time.strftime("%H:%M:%S")
        self.logs.append(f"`{timestamp}` {icon} {message}")
        with self.container.container():
            with st.expander("🖥️ AI 运行终端 (实时日志)", expanded=True):
                for log in self.logs[-5:]: # Show last 5 logs
                    st.markdown(log)
    
    def clear(self):
        self.container.empty()
        self.logs = []

# --- Logic: Smart Parsing V2 ---
def get_cell_text(cell):
    return cell.text.strip()

def look_around_for_context(table, r, c):
    """
    向上/向左查找，为通用标题寻找“父级上下文”。
    示例：如果单元格是“内容”（通用词），向左查找看到“课前”。
    返回：“上下文 > 单元格文本” 或 仅“单元格文本”
    """
    current_text = get_cell_text(table.cell(r, c))
    
    # 1. 向左查找 (同一行, c-1)
    if c > 0:
        left_text = get_cell_text(table.cell(r, c - 1))
        if left_text:
            return f"{left_text} > {current_text}"
            
    # 2. 向上查找 (r-1, 同一列) - 主要用于垂直合并的单元格
    if r > 0:
        up_text = get_cell_text(table.cell(r - 1, c))
        # 仅当上方文本是视觉合并或相关时使用 (启发式)
        # 如果没有确切的合并信息，这比较棘手，但我们可以尝试
        if up_text and up_text != current_text:
             return f"{up_text} > {current_text}"
    
    return current_text

def get_table_structure_v2(doc, logger=None):
    """
    V2 解析器：遍历所有表格，识别 Key（字段名）与 Target（填空位置）。
    针对 "教学过程" 等复杂表格，增强了上下文感知能力 (向左/向上查找)。
    """
    if logger: logger.log("开始扫描文档结构...", "📄")
    
    structure = []
    
    for t_idx, table in enumerate(doc.tables):
        rows = len(table.rows)
        cols = len(table.columns)
        
        processed_targets = set()

        for r in range(rows):
            for c in range(cols):
                try:
                    cell = table.cell(r, c)
                    text = cell.text.strip()
                    
                    if not text:
                        continue # 跳过空的 Key 单元格
                    
                    # 智能上下文 Key
                    # 如果文本很短/很通用 (如 "内容", "时间")，尝试追加上下文
                    full_key = text
                    if len(text) < 4 or text in ["内容", "学生活动", "教师活动", "设计意图"]:
                        full_key = look_around_for_context(table, r, c)
                    
                    target_coords = None
                    
                    # 策略 1: 向右看
                    if c + 1 < cols:
                        right_cell = table.cell(r, c + 1)
                        if not right_cell.text.strip() and (t_idx, r, c+1) not in processed_targets:
                            target_coords = (t_idx, r, c + 1)
                    
                    # 策略 2: 向下看 (如果向右没找到)
                    if target_coords is None and r + 1 < rows:
                         down_cell = table.cell(r + 1, c)
                         if not down_cell.text.strip() and (t_idx, r+1, c) not in processed_targets:
                             target_coords = (t_idx, r + 1, c)

                    if target_coords:
                        structure.append({
                            'key_text': full_key, # 使用上下文增强的 Key
                            'original_text': text,
                            'key_coords': (t_idx, r, c),
                            'target_coords': target_coords
                        })
                        processed_targets.add(target_coords)
                        
                except IndexError:
                    continue
    
    if logger: logger.log(f"文档扫描完成，共识别到 {len(structure)} 个填空点。", "✅")
    return structure

# --- Logic: Agentic Generation ---
def generate_deep_content(user_inputs, doc_keys, api_key, logger):
    """
    使用“思维链”方法生成内容。
    1. 研究/Key分析：搜索教学重点和解决措施。
    2. 生成：创建具体内容 (课前/课中/课后)。
    3. 映射：返回 JSON 格式结果。
    """
    llm = ChatOpenAI(
        model="deepseek-chat", 
        temperature=0.7,
        base_url="https://api.deepseek.com",
        openai_api_key=api_key
    )
    
    # 1. 研究阶段
    logger.log(f"正在分析课程主题: {user_inputs['课程大纲']}...", "🧠")
    logger.log("正在联网检索(模拟) 教学重点、难点及解决措施...", "🔍")
    
    # 2. 生成 Prompt
    keys_list = [item['key_text'] for item in doc_keys]
    
    system_prompt = """
    你是一位经验丰富的金牌讲师及教案编写专家。
    请根据【用户输入】的信息，为一份教案填充内容。
    
    关键要求：
    1. **教学重点与解决措施**：必须生成具体、专业的知识点和教学策略，绝不能留空。
    2. **教学过程（课前/课中/课后）**：
       - 请根据课程主题，自动设计 "课前预习任务"、"课中导入/讲授/练习"、"课后拓展" 的具体环节。
       - 识别文档Key中的上下文（如 "课前 > 内容"），填入对应的设计内容。
    3. **教案序号**：如果用户未填，请自动生成一个合理的序号（如 "No. 2024-01"）。
    4. **课程性质**：如果文档有此字段，根据课程内容自动判断（如 "理论课" 或 "理实一体"）。
    
    请输出一个纯 JSON 对象，格式为 {{ "文档里的Key": "你的建议内容" }}。
    """
    
    human_template = """
    【用户输入】: {user_inputs}
    
    【文档所有待填字段 (Keys)】: {keys_list}
    
    请开始编写，确保所有字段（尤其是教学过程和重点）都有丰富的内容。
    """
    
    prompt = ChatPromptTemplate.from_messages([
        ("system", system_prompt),
        ("human", human_template)
    ])
    
    chain = prompt | llm
    
    logger.log("正在撰写教案详细内容 (这可能需要 30-60 秒)...", "✍️")
    
    try:
        response = chain.invoke({
            "user_inputs": json.dumps(user_inputs, ensure_ascii=False),
            "keys_list": json.dumps(keys_list, ensure_ascii=False)
        })
        
        content = response.content
        # 稳健的 JSON 提取
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0]
        elif "```" in content:
            content = content.split("```")[1].split("```")[0]
            
        logger.log("AI 撰写完成！正在准备写入...", "✨")
        return json.loads(content)
        
    except Exception as e:
        logger.log(f"生成出错: {e}", "❌")
        st.error(f"Generate Error: {e}")
        return {}

def set_cell_text_preserving_style(cell, text):
    if not cell.paragraphs:
        cell.add_paragraph(text)
        return

    paragraph = cell.paragraphs[0]
    style_run = paragraph.runs[0] if paragraph.runs else None
    
    paragraph.clear()
    run = paragraph.add_run(text)
    
    if style_run:
        run.bold = style_run.bold
        run.italic = style_run.italic
        run.font.name = style_run.font.name
        if style_run.font.size:
            run.font.size = style_run.font.size

# --- Main App ---

def main():
    st.markdown("## 🤖 AI 智能教案生成器 (Pro)")
    
    # 0. Global Logger
    logger = ConsoleLogger()

    # 1. Sidebar Config
    with st.sidebar:
        st.header("⚙️ 1. 基础配置")
        api_key = st.text_input("DeepSeek API Key", type="password")
        
        st.header("📝 2. 课程基础信息")
        
        # New: Serial Number
        col1, col2 = st.columns(2)
        serial_no = col1.text_input("教案序号", "No. 01")
        time_val = col2.text_input("授课时间", "2024-03-20")

        dept = st.text_input("部门/院系", "信息工程学院")
        teacher = st.text_input("教师姓名", "张三")
        
        # New: Selectors for common fields
        course_type = st.selectbox("课程性质 (AI可覆盖)", ["理论课", "实践课", "理实一体化", "研讨课"])
        
        user_inputs = {
            "教案序号": serial_no,
            "时间": time_val,
            "部门": dept,
            "教师姓名": teacher,
            "课程性质": course_type
        }

        with st.expander("📚 更多课程细节 (选填)", expanded=False):
            user_inputs["课程名称"] = st.text_input("课程名称", "Python 程序设计")
            user_inputs["班级"] = st.text_input("班级", "23级计算机1班")
            user_inputs["地点"] = st.text_input("授课地点", "A305")
            user_inputs["授课学时"] = st.number_input("学时", 1, 4, 2)
            user_inputs["授课形式"] = st.selectbox("授课形式", ["线下面授", "线上直播", "混合式教学"])
            user_inputs["使用教材"] = st.text_input("使用教材", "《Python编程：从入门到实践》")
            user_inputs["考核方式"] = st.selectbox("考核方式", ["考查", "考试", "过程化考核"])

        st.header("🧠 3. 核心内容输入")
        topic_outline = st.text_area("本节课主题 & 大纲", height=250, 
                                     placeholder="输入本节课的主题，例如：\n主题：Python 循环结构\n1. while 循环\n2. fo 循环\n3. 案例实战")
        user_inputs["课程大纲"] = topic_outline

    # 2. Main Area
    uploaded_file = st.file_uploader("📂 上传 Word 教案模板 (.docx)", type=["docx"])

    if uploaded_file and st.button("🚀 开始生成", type="primary"):
        if not api_key:
            st.error("请先在左侧输入 DeepSeek API Key")
            return
        
        if not topic_outline:
            st.warning("请填写【课程主题 & 大纲】，否则 AI 无法生成内容。")
            return

        # Step 1: Parse
        doc = Document(uploaded_file)
        structure = get_table_structure_v2(doc, logger)
        
        if not structure:
            st.warning("未能识别到表格结构。请确保文档包含标准表格。")
            return

        # Step 2: Generate
        mapping = generate_deep_content(user_inputs, structure, api_key, logger)
        
        # Step 3: Fill
        logger.log("正在将内容写入文档...", "💾")
        fill_count = 0
        
        # Progress bar
        my_bar = st.progress(0)
        total_items = len(structure)
        
        for i, item in enumerate(structure):
            key = item['key_text']
            target_coords = item['target_coords']
            original_text = item['original_text']
            
            # Try to find match in generated mapping
            # Priority: Full Contextual Key -> Original Text -> Partial Match
            content = mapping.get(key) or mapping.get(original_text)
            
            if content:
                t_idx, r, c = target_coords
                target_cell = doc.tables[t_idx].cell(r, c)
                set_cell_text_preserving_style(target_cell, str(content))
                fill_count += 1
                if i % 5 == 0: # Log partially
                     logger.log(f"已填入: {key} -> {str(content)[:10]}...", "📝")
            
            my_bar.progress(min((i + 1) / total_items, 1.0))

        logger.log(f"🎉 全部完成！共填充 {fill_count} 个字段。", "✅")
        st.success(f"生成成功！")

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            label="⬇️ 下载生成的教案",
            data=buffer,
            file_name="generated_lesson_plan_v2.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    
if __name__ == "__main__":
    main()
