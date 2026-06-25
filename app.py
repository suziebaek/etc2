import streamlit as st
import docx
from docx import Document
import io
import re

st.set_page_config(
    page_title="정기평가 양식 상호 변환 시스템",
    page_icon="📝",
    layout="centered"
)

# ==========================================
# 1. 일반용 ➡️ 시험지용 변환 (소거형 - 표 내부 100% 동기화)
# ==========================================
def convert_general_to_exam_v5(source_doc):
    """
    원본 문서의 표(Table) 내부 지문, 쉼표, 폰트 스타일, 이미지를 절대 재생성하지 않고 100% 유지합니다.
    본문 스트림에서 불필요한 [Chapter ...] 단락만 소거하고 번호를 깨끗하게 재정렬합니다.
    """
    # 1. 원본 일반용 문서 구조를 바이너리 레벨에서 통째로 복제
    doc_stream = io.BytesIO()
    source_doc.save(doc_stream)
    doc_stream.seek(0)
    target_doc = Document(doc_stream)
    
    q_counter = 1
    paragraphs_to_remove = []
    
    # 2. 본문 단락 전체를 순회하며 조사
    for p in target_doc.paragraphs:
        p_text = p.text.strip()
        
        # [Chapter ...] 로 시작하는 출제자 메타데이터는 삭제 리스트에 등록
        if p_text.startswith("[Chapter"):
            paragraphs_to_remove.append(p)
            continue
            
        # 문항 번호 중복 방지 및 정렬 구조 다듬기 (1. 1. 현상 완벽 차단)
        if re.match(r'^\d+\.', p_text):
            # 단락 내부에서 기존 번호(예: '1.', '20.')와 공백을 정규식으로 완전히 도려냄
            clean_text = re.sub(r'^\d+\.\s*', '', p.text)
            
            # 단락 초기화 후 깨끗하게 정렬된 새 번호 프레임 빌드
            p.text = ""
            run_num = p.add_run(f"{q_counter}.  ")
            run_num.bold = True
            p.add_run(clean_text)
            
            q_counter += 1

    # 3. 본문 텍스트 흐름을 방해하던 [Chapter ...] XML 엘리먼트들을 원본 구조에서 안전하게 완전 삭제
    for p in paragraphs_to_remove:
        p_element = p._p
        p_parent = p_element.getparent()
        if p_parent is not None:
            p_parent.remove(p_element)
            
    # ★ 핵심: 본문 내부에 존재하는 모든 표(Table)는 수정하거나 새로 그리지 않고 
    # 원본 그대로 통째로 넘어가기 때문에 표 안의 지문과 구조가 일반용 문서와 무조건 100% 일치합니다.

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 2. 시험지용 ➡️ 일반용 변환
# ==========================================
def convert_exam_to_general_v5(source_doc):
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
st.markdown("원본 문서의 **표 내부 지문, 쉼표, 줄바꿈 및 스타일을 100% 복제**하여 양식을 변환합니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    sample_text = "\n".join([p.text for p in doc.paragraphs[:15] if p.text.strip()])
    
    if "[Chapter" in sample_text:
        default_index = 0
        detected_text = "🔍 **문서 양식 감지 결과:** [일반용 양식]이 확인되었습니다. 표 내부 지문 서식을 100% 유지하며 시험지용으로 정제합니다."
    else:
        default_index = 1
        detected_text = "🔍 **문서 양식 감지 결과:** [시험지용 양식]이 확인되었습니다. [일반용] 양식으로 변환합니다."
        
    st.info(detected_text)
    
    mode = st.radio(
        "변환 방향을 확인하거나 변경하세요:",
        ("일반용 ➡️ 시험지용", "시험지용 ➡️ 일반용"),
        index=default_index
    )
    
    if st.button("🚀 서식/표 무결성 변환 시작", use_container_width=True):
        with st.spinner("표 내부 데이터 구조와 문항 서식을 동기화하고 있습니다..."):
            try:
                if "일반용 ➡️ 시험지용" in mode:
                    out_bytes = convert_general_to_exam_v5(doc)
                    file_name = "변환_시험지용_정기평가.docx"
                else:
                    out_bytes = convert_exam_to_general_v5(doc)
                    file_name = "변환_일반용_정기평가.docx"
                
                st.success("🎉 표 지문 보존 및 중복 번호 교정 변환이 성공적으로 완료되었습니다!")
                st.download_button(
                    label="💾 변환된 무결성 Word 파일 다운로드",
                    data=out_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 변환 처리 중 기술적 오류가 발생했습니다: {str(e)}")
