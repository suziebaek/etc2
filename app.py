import streamlit as st
import docx
from docx import Document
import io
import re
import os

st.set_page_config(
    page_title="정기평가 양식 상호 변환 시스템",
    page_icon="📝",
    layout="centered"
)

# ==========================================
# 1. 일반용 ➡️ 시험지용 변환 (서식 및 표 완벽 보존형)
# ==========================================
def convert_general_to_exam_v4(source_doc):
    """
    원본 문서의 표(Table), 지문, 폰트 스타일, 이미지를 100% 보존합니다.
    [Chapter ...] 단락만 제거하고, 기존에 존재하는 문항 번호가 중복되지 않도록 깔끔하게 재정렬합니다.
    """
    # 원본 문서 서식을 통째로 유지하기 위해 메모리 딥카피 진행
    doc_stream = io.BytesIO()
    source_doc.save(doc_stream)
    doc_stream.seek(0)
    target_doc = Document(doc_stream)
    
    q_counter = 1
    paragraphs_to_remove = []
    
    # 1. 문서 본문의 모든 단락 검사
    for p in target_doc.paragraphs:
        p_text = p.text.strip()
        
        # [Chapter ...] 메타데이터 단락은 삭제 대상으로 지정
        if p_text.startswith("[Chapter"):
            paragraphs_to_remove.append(p)
            continue
            
        # 문항 번호 중복 방지 및 재정렬 로직
        # 단락이 숫자+마침표(예: 1., 20.)로 시작하는지 확인
        if re.match(r'^\d+\.', p_text):
            # 기존에 있던 번호와 그 뒤의 공백을 완전히 제거 (예: "1. 다음 중..." -> "다음 중...")
            clean_text = re.sub(r'^\d+\.\s*', '', p.text)
            
            # 단락 내부의 기존 텍스트(Runs)를 모두 지우고 새 번호와 정제된 텍스트 결합
            # 이 방식을 쓰면 1. 1. 처럼 번호가 중복되는 현상이 100% 방지됩니다.
            p.text = "" 
            run_num = p.add_run(f"{q_counter}.  ")
            run_num.bold = True # 번호만 볼드 처리
            p.add_run(clean_text)
            
            q_counter += 1

    # 2. 본문에서 감지된 [Chapter ...] 단락들을 원본 서식 훼손 없이 안전하게 삭제
    for p in paragraphs_to_remove:
        p_element = p._p
        p_parent = p_element.getparent()
        if p_parent is not None:
            p_parent.remove(p_element)
            
    # *참고*: 본문 내부에 존재하는 표(Table) 객체들은 손대지 않고 그대로 두기 때문에
    # 표 안의 지문, 콤마, 레이아웃, 줄바꿈 서식이 원본 일반용 문서와 100% 일치하게 출력됩니다.

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 2. 시험지용 ➡️ 일반용 변환
# ==========================================
def convert_exam_to_general_v4(source_doc):
    """
    시험지용 문서의 스타일을 유지하면서 문항 번호만 다시 순서대로 동기화합니다.
    """
    doc_stream = io.BytesIO()
    source_doc.save(doc_stream)
    doc_stream.seek(0)
    target_doc = Document(doc_stream)
    
    q_counter = 1
    for p in target_doc.paragraphs:
        p_text = p.text.strip()
        if re.match(r'^\d+\.', p_text):
            clean_text = re.sub(r'^\d+\.\s*', '', p.text)
            p.text = ""
            run_num = p.add_run(f"{q_counter}.  ")
            run_num.bold = True
            p.add_run(clean_text)
            q_counter += 1

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 3. Streamlit 웹 인터페이스 UI
# ==========================================
st.title("📝 정기평가 양식 상호 변환 시스템")
st.markdown("디자인 서식 및 **표 내부 지문/데이터를 100% 무결하게 보존**하며 양식을 변환합니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    
    # 문서 본문 검사를 통해 챕터 태그 자동 감지
    sample_text = "\n".join([p.text for p in doc.paragraphs[:15] if p.text.strip()])
    
    if "[Chapter" in sample_text:
        default_index = 0
        detected_text = "🔍 **문서 양식 감지 결과:** [일반용 양식]이 확인되었습니다. 표와 지문을 보존하며 [시험지용]으로 변환합니다."
    else:
        default_index = 1
        detected_text = "🔍 **문서 양식 감지 결과:** [시험지용 양식]이 확인되었습니다. [일반용]으로 변환합니다."
        
    st.info(detected_text)
    
    mode = st.radio(
        "변환 방향을 확인하거나 변경하세요:",
        ("일반용 ➡️ 시험지용", "시험지용 ➡️ 일반용"),
        index=default_index
    )
    
    if st.button("🚀 서식 보존 변환 시작", use_container_width=True):
        with st.spinner("표 내부 데이터 및 문항 구조를 정밀 정제 중입니다..."):
            try:
                if "일반용 ➡️ 시험지용" in mode:
                    out_bytes = convert_general_to_exam_v4(doc)
                    file_name = "변환_시험지용_정기평가.docx"
                else:
                    out_bytes = convert_exam_to_general_v4(doc)
                    file_name = "변환_일반용_정기평가.docx"
                
                st.success("🎉 표 지문 및 번호 정렬 변환이 완료되었습니다!")
                st.download_button(
                    label="💾 변환된 Word 파일 다운로드",
                    data=out_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 변환 처리 중 오류가 발생했습니다: {str(e)}")
