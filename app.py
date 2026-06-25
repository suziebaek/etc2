import streamlit as st
import docx
from docx import Document
import io
import re

# Set page configuration
st.set_page_config(
    page_title="정기평가 양식 변환기",
    page_icon="📝",
    layout="centered"
)

# ==========================================
# 1. 정규식 기반 데이터 파싱 및 정제 함수들
# ==========================================

def parse_doc1_to_data(doc):
    """
    1번 문서(검수용)에서 텍스트를 추출하여 구조화된 문항 데이터 리스트를 만듭니다.
    """
    full_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    raw_questions = re.split(r'(\[Chapter\s+\d+:[^\]]+\])', full_text)
    
    questions_data = []
    current_meta = ""
    
    for part in raw_questions:
        part = part.strip()
        if not part:
            continue
        if part.startswith("[Chapter"):
            current_meta = part
        else:
            lines = [l.strip() for l in part.split('\n') if l.strip()]
            if not lines:
                continue
            
            q_text = lines[0]
            options = []
            box_text = ""
            
            for line in lines[1:]:
                if "The following table:" in line or ("____" in line and not line.startswith(('①','②','③','④','⑤'))):
                    box_text = line.replace("The following table:", "").strip()
                elif line.startswith(('①','②','③','④','⑤')):
                    options.append(line)
            
            questions_data.append({
                "meta": current_meta,
                "q_text": q_text,
                "box_text": box_text,
                "options": options
            })
    return questions_data

def parse_doc2_to_data(doc):
    """
    2번 문서(학생용)에서 텍스트를 추출하여 구조화된 문항 데이터 리스트를 만듭니다.
    """
    full_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    raw_questions = re.split(r'(\n\d+\.\s+)', "\n" + full_text)
    
    questions_data = []
    
    for i in range(1, len(raw_questions), 2):
        q_num_prefix = raw_questions[i].strip()
        q_body = raw_questions[i+1].strip()
        
        lines = [l.strip() for l in q_body.split('\n') if l.strip()]
        if not lines:
            continue
            
        q_text = q_num_prefix + " " + lines[0]
        options = []
        box_text = ""
        
        for line in lines[1:]:
            if "The following table:" in line or ("," in line and not line.startswith(('①','②','③','④','⑤'))):
                box_text = line.replace("The following table:", "").strip()
            elif line.startswith(('①','②','③','④','⑤')):
                options.append(line)
                
        questions_data.append({
            "meta": "[Chapter 01: Default Chapter, pg 10, 교재 연계, Fill in the blank]", # 임시 메타데이터 기본값
            "q_text": q_text,
            "box_text": box_text,
            "options": options
        })
    return questions_data

# ==========================================
# 2. 새로운 Word 문서 생성 함수들
# ==========================================

def create_doc1(questions_data):
    """데이터 리스트를 바탕으로 1번 형식(검수용) .docx를 생성합니다."""
    doc = Document()
    for q in questions_data:
        doc.add_paragraph(q["meta"])
        doc.add_paragraph(q["q_text"])
        if q["box_text"]:
            doc.add_paragraph("The following table:")
            doc.add_paragraph(q["box_text"])
        for opt in q["options"]:
            doc.add_paragraph(opt)
        doc.add_paragraph("")
    
    b_io = io.BytesIO()
    doc.save(b_io)
    return b_io.getvalue()

def create_doc2(questions_data):
    """데이터 리스트를 바탕으로 2번 형식(학생용) .docx를 생성합니다."""
    doc = Document()
    doc.add_paragraph("2026년 봄학기 THE OPEN 정기평가 (Level P)")
    doc.add_paragraph("[ PartⅡ 문법 | 20문항 20분 ]\n")
    
    for idx, q in enumerate(questions_data, start=1):
        clean_q_text = re.sub(r'^\d+\.\s*', '', q["q_text"])
        doc.add_paragraph(f"{idx}. {clean_q_text}")
        
        if q["box_text"]:
            doc.add_paragraph("The following table:")
            doc.add_paragraph(q["box_text"])
            
        for opt in q["options"]:
            doc.add_paragraph(opt)
        doc.add_paragraph("")
        
    b_io = io.BytesIO()
    doc.save(b_io)
    return b_io.getvalue()

# ==========================================
# 3. Streamlit 웹 인터페이스 UI
# ==========================================

st.title("📝 정기평가 양식 상호 변환기")
st.markdown("업로드하신 Word 파일의 양식을 자동으로 분석하여 **검수용(1번)** 또는 **학생용(2번)** 문서로 상호 변환해 드립니다.")

uploaded_file = st.file_uploader("변환할 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    
    # 문서 스타일 감지를 위해 상단 텍스트 일부 분석
    sample_text = "\n".join([p.text for p in doc.paragraphs[:5] if p.text.strip()])
    
    # 기본 인식 모드 설정
    if "[Chapter" in sample_text:
        default_index = 0  # 1번 -> 2번
        detected_text = "🔍 **문서 양식 감지 결과:** [출제자 검수용 양식(1번)]이 감지되었습니다."
    else:
        default_index = 1  # 2번 -> 1번
        detected_text = "🔍 **문서 양식 감지 결과:** [학생 시험지용 양식(2번)]이 감지되었습니다."
        
    st.info(detected_text)
    
    # 자동 인식을 보여주되 사용자가 원하면 수동 전환도 가능하도록 라디오 버튼 배치
    mode = st.radio(
        "변환 방향을 확인하거나 선택하세요:",
        ("1번 ➡️ 2번 (검수용 ➡️ 학생용)", "2번 ➡️ 1번 (학생용 ➡️ 검수용)"),
        index=default_index
    )
    
    if st.button("🚀 변환하기", use_container_width=True):
        with st.spinner("문서 구조를 파싱하여 양식을 변환하고 있습니다..."):
            try:
                if "1번 ➡️ 2번" in mode:
                    data = parse_doc1_to_data(doc)
                    out_bytes = create_doc2(data)
                    file_name = "변환_2번_학생용_정기평가.docx"
                else:
                    data = parse_doc2_to_data(doc)
                    out_bytes = create_doc1(data)
                    file_name = "변환_1번_검수용_정기평가.docx"
                
                st.success("🎉 양식 변환이 완료되었습니다!")
                st.download_button(
                    label="💾 변환된 Word 파일 다운로드",
                    data=out_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 변환 중 오류가 발생했습니다. 파일 내부 형식을 확인해주세요. (오류 내용: {str(e)})")
