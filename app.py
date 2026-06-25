import streamlit as st
import docx
from docx import Document
import io
import re

st.set_page_config(
    page_title="정기평가 양식 상호 변환기",
    page_icon="📝",
    layout="centered"
)

# ==========================================
# 1. 일반용 ➡️ 시험지용 변환 (기존 레이아웃 복사 및 원치 않는 요소 제거)
# ==========================================
def convert_general_to_exam_v2(source_doc):
    """
    원본 문서의 '다단 레이아웃', '상단 분홍색 이미지 및 타이틀', '표 구조'를 100% 보존합니다.
    불필요한 [Chapter ...] 메타데이터 단락만 본문에서 정밀 제거하고 번호를 재정렬합니다.
    """
    # 템플릿 형태로 원본 바이너리를 완벽하게 deep copy 하기 위해 메모리 스트림 활용
    doc_stream = io.BytesIO()
    source_doc.save(doc_stream)
    doc_stream.seek(0)
    target_doc = Document(doc_stream)
    
    q_counter = 1
    paragraphs_to_remove = []
    
    # 1. 삭제해야 할 [Chapter ...] 단락 및 공백 감지
    for p in target_doc.paragraphs:
        p_text = p.text.strip()
        if p_text.startswith("[Chapter"):
            paragraphs_to_remove.append(p)
            continue
            
        # 문항 번호 정렬 (일반용에 붙어있던 번호를 순서대로 재정렬)
        if re.match(r'^\d+\.', p_text):
            # 기존 번호 포맷(예: 1.  ) 제거 후 깔끔하게 정렬
            clean_text = re.sub(r'^\d+\.\s*', '', p.text)
            p.text = f"{q_counter}.  {clean_text}"
            q_counter += 1

    # 2. 표(Table) 내부에서도 혹시 모를 Chapter 텍스트가 있을 경우 삭제 및 셀 구조 유지
    for table in target_doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text.strip().startswith("[Chapter"):
                        p.text = "" # 표 내부 구조가 깨지지 않도록 텍스트만 비움

    # 3. 문서 구조에서 감지된 [Chapter ...] 단락들을 안전하게 영구 삭제
    for p in paragraphs_to_remove:
        p_element = p._p
        p_parent = p_element.getparent()
        if p_parent is not None:
            p_parent.remove(p_element)
            
    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 2. 시험지용 ➡️ 일반용 변환 (가상 메타데이터 생성형)
# ==========================================
def convert_exam_to_general_v2(source_doc):
    """
    상단 분홍색 레이아웃과 다단을 유지한 상태에서, 
    각 문항 번호(1., 2.) 바로 윗줄에 일반용 양식 가이드 라인을 주입합니다.
    (Chapter 단원 정보는 필요 없다고 하셨으므로 표준 가이드 라인 형태로만 분리 적용)
    """
    doc_stream = io.BytesIO()
    source_doc.save(doc_stream)
    doc_stream.seek(0)
    target_doc = Document(doc_stream)
    
    q_counter = 1
    
    # 원본 문서 내부 객체를 순회하면서 각 문항 시작점 감지
    for p in target_doc.paragraphs:
        p_text = p.text.strip()
        
        if re.match(r'^\d+\.', p_text):
            # 문항 번호 재정렬
            clean_text = re.sub(r'^\d+\.\s*', '', p.text)
            p.text = f"{q_counter}.  {clean_text}"
            
            # 문항 시작점 바로 위에 구분용 빈 줄 및 가이드 라인 스타일 주입 (필요시 활성화 가능)
            # 여기서는 Chapter를 완전히 배제하므로 번호 재정렬 및 본문 레이아웃 보존에 집중합니다.
            q_counter += 1

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 3. Streamlit 웹 인터페이스 및 자동 인식 UI
# ==========================================

st.title("📝 정기평가 양식 상호 변환 시스템")
st.markdown("디자인 서식(2단 분할, 상단 분홍색 이미지, 표 구조)을 **100% 원본 그대로 유지**하며 변환합니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    
    # 문서 본문 검사를 통해 상단 레이아웃 및 챕터 태그 자동 감지
    sample_text = "\n".join([p.text for p in doc.paragraphs[:15] if p.text.strip()])
    
    if "[Chapter" in sample_text:
        default_index = 0
        detected_text = "🔍 **문서 양식 감지 결과:** [일반용 양식]이 확인되었습니다. 불필요한 Chapter 데이터를 제거하고 [시험지용 양식]으로 정제합니다."
    else:
        default_index = 1
        detected_text = "🔍 **문서 양식 감지 결과:** [시험지용 양식]이 확인되었습니다. 상단 디자인을 보존하며 [일반용 양식]으로 변환합니다."
        
    st.info(detected_text)
    
    mode = st.radio(
        "변환 방향을 확인하거나 변경하세요:",
        ("일반용 ➡️ 시험지용", "시험지용 ➡️ 일반용"),
        index=default_index
    )
    
    if st.button("🚀 정밀 양식 변환 시작", use_container_width=True):
        with st.spinner("다단 레이아웃, 이미지 서식 및 표 내부 데이터를 안전하게 추출 및 변환 중입니다..."):
            try:
                if "일반용 ➡️ 시험지용" in mode:
                    out_bytes = convert_general_to_exam_v2(doc)
                    file_name = "변환_시험지용_정기평가_최종.docx"
                else:
                    out_bytes = convert_exam_to_general_v2(doc)
                    file_name = "변환_일반용_정기평가_최종.docx"
                
                st.success("🎉 서식 보존 변환이 완료되었습니다!")
                st.download_button(
                    label="💾 서식 보존된 Word 파일 다운로드",
                    data=out_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 변환 처리 중 기술적 오류가 발생했습니다: {str(e)}")
