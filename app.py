import streamlit as st
import os
from docx import Document
from docx.shared import Pt
import json
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import JsonOutputParser
from io import BytesIO

# --- Configuration ---
st.set_page_config(page_title="AI 智能教案生成器", layout="wide")

# --- Helper Functions ---

def get_table_structure(doc):
    """
    Traverses all tables in the document to map "Keys" (labels) to "Targets" (empty cells).
    Returns a list of binding objects:
    {
        'key_text': str,
        'key_coords': (table_idx, row_idx, col_idx),
        'target_coords': (table_idx, row_idx, col_idx)
    }
    """
    structure = []
    
    for t_idx, table in enumerate(doc.tables):
        rows = len(table.rows)
        cols = len(table.columns)
        
        # We process cells to find "Label -> Empty Cell" relationships.
        # Simple Heuristic: 
        # 1. Look Right: If cell(r, c) has text and cell(r, c+1) is empty, map them.
        # 2. Look Down: If cell(r, c) has text and cell(r+1, c) is empty (and right wasn't a match), map them.
        
        processed_targets = set()

        for r in range(rows):
            for c in range(cols):
                try:
                    cell = table.cell(r, c)
                    text = cell.text.strip()
                    
                    if not text:
                        continue # Skip empty key cells
                    
                    # Potential Key found: `text`
                    
                    key_coords = (t_idx, r, c)
                    target_coords = None
                    
                    # Strategy 1: Look Right
                    if c + 1 < cols:
                        right_cell = table.cell(r, c + 1)
                        if not right_cell.text.strip() and (t_idx, r, c+1) not in processed_targets:
                            target_coords = (t_idx, r, c + 1)
                    
                    # Strategy 2: Look Down (only if Right didn't work)
                    if target_coords is None and r + 1 < rows:
                         down_cell = table.cell(r + 1, c)
                         if not down_cell.text.strip() and (t_idx, r+1, c) not in processed_targets:
                             target_coords = (t_idx, r + 1, c)

                    if target_coords:
                        structure.append({
                            'key_text': text,
                            'key_coords': key_coords,
                            'target_coords': target_coords
                        })
                        processed_targets.add(target_coords)
                        
                except IndexError:
                    continue
                    
    return structure

def generate_ai_content(user_inputs, doc_keys, api_key):
    """
    Uses LangChain to map user inputs to document keys and generate missing content.
    """
    if not api_key:
        st.error("请输入 OpenAI API Key")
        return {}

    llm = ChatOpenAI(
        model="deepseek-chat", 
        temperature=0.7,
        base_url="https://api.deepseek.com",
        openai_api_key=api_key
    )

    # Convert keys to a clean list of strings
    keys_list = [item['key_text'] for item in doc_keys]
    
    # Prompt Design
    system_prompt = """
    你是一个专业的教案编写助手。
    你的任务是将用户提供的【表单信息】填入到【文档结构列表】中。
    
    规则：
    1. 如果【文档结构列表】中的字段在【表单信息】中有直接对应（如姓名、课程名），直接填入。
    2. 如果需要生成内容（如“教学目标”、“学情分析”），请根据【表单信息】中的“课程大纲/主题”进行专业扩写。
    3. 如果某个字段无法生成且无信息，填入 "（空）" 或留白。
    4. 输出必须是 JSON 格式： {{ "文档字段名": "填入内容" }}
    """
    
    human_template = """
    【表单信息】: {user_inputs}
    
    【文档结构列表】: {keys_list}
    
    请输出 JSON 映射结果。
    """
    
    prompt = ChatPromptTemplate.from_messages([
        ("system", system_prompt),
        ("human", human_template)
    ])
    
    chain = prompt | llm
    
    try:
        response = chain.invoke({
            "user_inputs": json.dumps(user_inputs, ensure_ascii=False),
            "keys_list": json.dumps(keys_list, ensure_ascii=False)
        })
        
        # Parse JSON from content (Found robustly)
        content = response.content
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0]
        elif "```" in content:
            content = content.split("```")[1].split("```")[0]
            
        return json.loads(content)
        
    except Exception as e:
        st.error(f"AI 生成失败: {e}")
        return {}

def set_cell_text_preserving_style(cell, text):
    """
    Sets text in a cell while attempting to preserve the style of the first paragraph/run.
    """
    if not cell.paragraphs:
        cell.add_paragraph(text)
        return

    paragraph = cell.paragraphs[0]
    
    # Check if there's existing style/runs to copy
    style_run = None
    if paragraph.runs:
        style_run = paragraph.runs[0]
    
    # Clear existing content but keep the paragraph object
    paragraph.clear()
    
    # Add new run
    run = paragraph.add_run(text)
    
    # Copy basic styles if they existed
    if style_run:
        run.bold = style_run.bold
        run.italic = style_run.italic
        run.font.name = style_run.font.name
        if style_run.font.size:
            run.font.size = style_run.font.size
            
    # Fallback: Try to ensure Chinese font compatibility if needed (Optional)
    # run.font.element.rPr.rFonts.setall(qn('w:eastAsia'), 'SimSun') 

# --- Main App Interface ---

def main():
    st.title("📚 AI 智能教案生成器 (DeepSeek 版)")
    st.markdown("上传任意 Word 表格模板，AI 自动识别字段并填入教案内容。")

    with st.sidebar:
        st.header("1. 配置与输入")
        api_key = st.text_input("DeepSeek API Key", type="password")
        
        st.subheader("基本信息")
        dept = st.text_input("部门/院系", "信息工程学院")
        teacher = st.text_input("教师姓名", "张三")
        course = st.text_input("课程名称", "Python 程序设计")
        cls = st.text_input("班级", "23级计算机1班")
        time = st.text_input("授课时间", "2024-03-20")
        location = st.text_input("授课地点", "A305")
        
        st.subheader("核心内容")
        topic_outline = st.text_area("本节课主题与大纲", height=200, 
                                     placeholder="例如：\n主题：Python 循环结构\n1. while 循环语法\n2. for 循环语法\n3. break 与 continue\n4. 实战案例：猜数字游戏")
        
        user_inputs = {
            "部门": dept,
            "教师姓名": teacher,
            "课程名称": course,
            "班级": cls,
            "时间": time,
            "地点": location,
            "课程大纲": topic_outline
        }

    # Main Area
    uploaded_file = st.file_uploader("上传 Word 教案模板 (.docx)", type=["docx"])

    if uploaded_file and st.button("开始生成"):
        if not api_key:
            st.warning("请先在左侧输入 API Key")
            return

        with st.spinner("1/3 正在解析文档结构..."):
            # Load doc
            doc = Document(uploaded_file)
            structure = get_table_structure(doc)
            
            if not structure:
                st.error("未在文档中检测到有效的表格结构，请检查模板。")
                return
            
            # Show preview of detected keys (optional debugging)
            # st.write(f"检测到 {len(structure)} 个填空项: {[s['key_text'] for s in structure]}")

        with st.spinner("2/3 AI 正在生成教案内容..."):
            # Generate content
            mapping_result = generate_ai_content(user_inputs, structure, api_key)
            if not mapping_result:
                st.stop()

        with st.spinner("3/3 正在写入文档..."):
            # Fill content
            fill_count = 0
            for item in structure:
                key = item['key_text']
                target_coords = item['target_coords']
                
                # Fuzzy get (in case keys slightly mismatch or AI shortened them)
                # Here we assume exact match from the JSON Key to parsed Key
                content = mapping_result.get(key)
                
                if content:
                    t_idx, r, c = target_coords
                    target_cell = doc.tables[t_idx].cell(r, c)
                    set_cell_text_preserving_style(target_cell, str(content))
                    fill_count += 1
            
            st.success(f"生成完成！已填充 {fill_count} 个数据项。")
            
            # Save to buffer
            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.download_button(
                label="下载生成的教案",
                data=buffer,
                file_name="generated_lesson_plan.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

if __name__ == "__main__":
    main()
