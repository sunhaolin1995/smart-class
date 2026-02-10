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
st.set_page_config(page_title="AI 智能教案生成器 (V11 Fixed)", layout="wide", initial_sidebar_state="expanded")

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

# --- Logic: Smart Parsing V10 ---
def get_cell_text(cell):
    return cell.text.strip()

def set_cell_text_preserving_style(cell, text):
    """保留原有格式写入文本"""
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

def get_table_structure_v10(doc, logger=None):
    """
    【V10 终极修补版】
    1. 增加抓取 "教学环节" 列。
    2. 解锁行数限制 (MAX_ROWS = 50)。
    3. 优化 Key 生成逻辑。
    """
    if logger: logger.log("正在执行全量深度扫描 (V10 包含课后修复)...", "🔍")
    
    structure = []
    processed_cell_ids = set() 
    processed_keys = set()     

    MAX_ROWS_PER_PHASE = 50 

    def is_instructional(text):
        return len(text) > 30 or any(k in text for k in ["思政案例", "确保思政", "比例可根据"])

    for t_idx, table in enumerate(doc.tables):
        rows = table.rows
        all_text = "".join([c.text for r in rows for c in r.cells])
        
        # --- 策略 A：教学过程矩阵表 (Index 1) ---
        if any(k in all_text for k in ["教师活动", "学生活动", "设计意图"]):
            if logger: logger.log(f"正在解析教学过程表，准备提取全部行...", "🎯")
            
            col_map = {}
            for r_idx in range(min(3, len(rows))):
                for c_idx in range(len(table.columns)):
                    txt = table.cell(r_idx, c_idx).text.strip()
                    if txt in ["教学环节", "教学内容", "教师活动", "学生活动", "设计意图"]:
                        col_map[c_idx] = txt

            current_phase = "教学过程"
            phase_counter = {} 

            for r in range(len(rows)):
                row_raw_text = "".join(list(dict.fromkeys([c.text.strip() for c in rows[r].cells])))
                
                if row_raw_text in ["课前", "课中", "课后", "巩固拓展"]:
                    current_phase = row_raw_text
                    phase_counter[current_phase] = 0
                    continue
                
                if phase_counter.get(current_phase, 0) >= MAX_ROWS_PER_PHASE:
                    continue

                row_has_vacancy = False
                for c_idx, col_name in col_map.items():
                    target_cell = table.cell(r, c_idx)
                    if not target_cell.text.strip() and target_cell.text.strip() != col_name:
                        if target_cell._tc not in processed_cell_ids:
                            row_has_vacancy = True
                            full_key = f"{current_phase} > {col_name}_行{r}"
                            
                            structure.append({
                                'key_text': full_key,
                                'original_text': col_name,
                                'target_coords': (t_idx, r, c_idx),
                                'is_teaching_process': True
                            })
                            processed_cell_ids.add(target_cell._tc)
                
                if row_has_vacancy:
                    phase_counter[current_phase] = phase_counter.get(current_phase, 0) + 1
            continue

        # --- 策略 B：通用信息表 ---
        for r in range(len(rows)):
            for c in range(len(table.columns)):
                cell = table.cell(r, c)
                text = cell.text.strip().replace("\n", "").replace(" ", "")
                
                if not text or is_instructional(text): continue
                if text in processed_keys and len(text) < 10: continue

                target = None
                if c + 1 < len(table.columns):
                    r_c = table.cell(r, c + 1)
                    if not r_c.text.strip(): target = (r, c + 1, r_c._tc)
                if not target and r + 1 < len(rows):
                    d_c = table.cell(r + 1, c)
                    if not d_c.text.strip(): target = (r + 1, c, d_c._tc)
                
                if target:
                    tr, tc, t_id = target
                    if t_id not in processed_cell_ids:
                        full_key = text
                        p_header = table.cell(r, 0).text.strip()
                        if p_header in ["学情分析", "教学目标", "教学资源", "教学反思"]:
                            if p_header != text: full_key = f"{p_header} > {text}"

                        structure.append({
                            'key_text': full_key,
                            'original_text': text,
                            'target_coords': (t_idx, tr, tc),
                            'is_teaching_process': False
                        })
                        processed_cell_ids.add(t_id)
                        processed_keys.add(text)

    return structure

# --- Logic: Agentic Generation ---
def generate_deep_content(user_inputs, doc_keys, api_key, logger):
    """
    Prompt 升级版：
    修复了 JSON 示例花括号未转义导致的 LangChain 报错。
    """
    llm = ChatOpenAI(
        model="deepseek-chat", 
        temperature=0.7,
        base_url="https://api.deepseek.com",
        openai_api_key=api_key
    )
    
    # 1. 研究阶段
    logger.log(f"正在深度分析: {user_inputs['课程大纲']}", "🧠")
    logger.log("正在挖掘思政融合点 & 教学解决措施...", "🔍")
    
    keys_list = [item['key_text'] for item in doc_keys]
    
    # 注意：这里的 JSON 示例已经改成了 {{ ... }}，这就是修复点！
    system_prompt = """
你是一位顶尖的职业教育/高等教育教案编写专家。你的任务是根据用户提供的基础信息，填满文档中所有的空缺字段。

## ⚠️ 最高优先级指令（必须严格执行）

1.  **用户输入优先**：
    -   如果 Key 是 "授课时间"、"授课地点"、"班级"、"教师姓名"，**必须直接使用【用户输入】中的对应值**，严禁自己编造或留空。

2.  **必须填满所有教学过程的格子**：
    -   你会收到像 "课中 > 教师活动_行10", "课中 > 教师活动_行11" 这样的大量 Key。
    -   **有多少个 Key，就必须输出多少条内容！** 严禁合并，严禁偷懒，严禁只写前几行。
    -   如果是 "课后" 环节，即使有很多行，也要分别填写（如：布置作业、预习下节、整理笔记等）。
    -   **"教学环节" 列**：请填入简短的步骤名称，如 "导入新课"、"案例分析"、"小组讨论"、"课堂总结"。

3.  **特殊字段内容要求**：
    -   **"课程思政融合点" / "素质目标"**：请务必进行“联网搜索式”创作，结合课程内容，填入具体的家国情怀、职业道德、工匠精神、科学思维等融合点。**绝对不能留空！**
    -   **"解决措施"**：每一个 "教学难点" 对应的地方，必须填入具体的 "解决措施"。

4.  **内容连贯性**：
    -   "课中" 的多行内容应构成一个完整的教学流。例如：_行8 是导入，_行9-15 是讲解，_行16-20 是练习。

## 输出格式
-   输出纯 JSON 对象：`{{ "Key的名字": "填充内容" }}`
-   不要输出 Markdown 代码块标记。
"""
    
    human_template = """
    【用户输入数据】: {user_inputs}
    
    【需要填充的所有 Key】: {keys_list}
    
    请开始生成。请记住：授课时间用用户输入的；思政点要具体；教学过程的每一行都要填满，不要遗漏课后环节。
    """
    
    prompt = ChatPromptTemplate.from_messages([
        ("system", system_prompt),
        ("human", human_template)
    ])
    
    chain = prompt | llm
    
    logger.log("AI 正在根据格子数量撰写全量教案 (内容较多，请耐心等待)...", "✍️")
    
    try:
        response = chain.invoke({
            "user_inputs": json.dumps(user_inputs, ensure_ascii=False),
            "keys_list": json.dumps(keys_list, ensure_ascii=False)
        })
        
        content = response.content
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0]
        elif "```" in content:
            content = content.split("```")[1].split("```")[0]
            
        result = json.loads(content)
        
        # --- 硬逻辑补丁 ---
        user_mapping = {
            "授课时间": user_inputs.get("时间"),
            "教案序号 > 授课时间": user_inputs.get("时间"),
            "授课地点": user_inputs.get("地点"),
            "教案序号 > 授课地点": user_inputs.get("地点"),
            "授课班级": user_inputs.get("班级"),
            "授课内容 > 授课班级": user_inputs.get("班级"),
            "教师姓名": user_inputs.get("教师姓名")
        }
        
        for k, v in user_mapping.items():
            if v:
                result[k] = v
                
        logger.log("AI 撰写完成！正在写入文档...", "✨")
        return result
        
    except Exception as e:
        logger.log(f"生成出错: {e}", "❌")
        st.error(f"Generate Error: {e}")
        return {}

# --- Main App ---

def main():
    st.markdown("## 🤖 AI 智能教案生成器 (V11 Pro)")
    
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
        structure = get_table_structure_v10(doc, logger)
        
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
            content = mapping.get(key) or mapping.get(original_text)
            
            if content:
                t_idx, r, c = target_coords
                target_cell = doc.tables[t_idx].cell(r, c)
                set_cell_text_preserving_style(target_cell, str(content))
                fill_count += 1
                if i % 10 == 0: 
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
            file_name="generated_lesson_plan_v11.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    
if __name__ == "__main__":
    main()