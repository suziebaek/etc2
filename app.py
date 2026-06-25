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
# 1. 일반용 ➡️ 시험지용 변환 (템플릿 디자인 주입 + 표 내부 지문 보존형)
# ==========================================
def convert_general_to_exam_final(source_doc):
    """
    [일반용 ➡️ 시험지용] 변환
    내장된 템플릿(template_exam.docx)의 2단 레이아웃과 분홍색 상단 디자인 틀을 가져온 뒤,
    그 안에 일반용 문서의 발문, 선지, 이미지, 표(내부 지문 서식 포함)를 깨짐 없이 채워 넣습니다.
    """
    template_path = "template_exam.docx"
    
    # 디자인 서식용 템플릿 파일 로드
    if os.path.exists(template_path):
        target_doc = Document(template_path)
    else:
        st.warning("⚠️ 'template_exam.docx'(시험지 디자인 템플릿) 파일이 서버에서 감지되지 않아 기본 서식으로 변환을 진행합니다.")
        target_doc = Document()
        target_doc.add_paragraph("2026년 봄학기 THE OPEN 정기평가 (Level P)")
        target_doc.add_paragraph("[ PartⅡ 문법 | 20문항 20분 ]\n")

    q_counter = 1
    
    # 원본(일반용) 문서의 본문 요소를 순차적으로 분석하며 복사
    for element in source_doc.element.body:
        if element.tag.endswith('p'): # 단락(텍스트/이미지)일 때
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            # 챕터 메타데이터 및 단순 공백 단락은 제외
            if not p_text or p_text.startswith("[Chapter"):
                continue
            
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_before = p.paragraph_format.space_before
            new_p.paragraph_format.space_after = p.paragraph_format.space_after
            
            # 문항 번호 시작 부분 감지 (예: 1., 12.)
            if re.match(r'^\d+\.', p_text):
                # 기존에 적혀있던 앞부분 번호 패턴(숫자+마침표+공백)을 완벽하게 제거하여 '1. 1.' 중복 방지
                clean_text = re.sub(r'^\d+\.\s*', '', p_text)
                
                # 깔끔하게 재정렬된 번호만 새로 주입
                run_num = new_p.add_run(f"{q_counter}.  ")
                run_num.bold = True
                q_counter += 1
                
                # 번호 뒷부분의 문항 텍스트 및 개별 Run 서식(이미지 포함) 복사
                for run in p.runs:
                    # 번호 텍스트를 중복해서 복사하지 않도록 제어
                    run_clean = re.sub(r'^\d+\.\s*', '', run.text) if run.text.strip().startswith(p_text[:2]) else run.text
                    if run_clean:
                        new_run = new_p.add_run(run_clean)
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                        if run._r.xpath('.//w:drawing'):
                            new_run._r.append(run._r.xpath('.//w:drawing')[0])
            else:
                # 일반 보기 선지, 발문 라인은 스타일 그대로 이전
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    if run._r.xpath('.//w:drawing'):
                        new_run._r.append(run._r.xpath('.//w:drawing')[0])
                        
        elif element.tag.endswith('tbl'): # 표(Table) 구조물 발견 시
            # 표의 틀과 구조를 복제
            tbl = docx.table.Table(element, source_doc)
            new_tbl = target_doc.add_table(rows=len(tbl.rows), cols=len(tbl.columns))
            new_tbl.style = tbl.style
            
            # 표 내부의 지문, 셀 텍스트, 줄바꿈 서식을 1:1로 매핑하여 무결하게 복사
            for r_idx, row in enumerate(tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    new_cell = new_tbl.cell(r_idx, c_idx)
                    
                    # 단순 텍스트(.text = cell.text) 방식을 버리고, 셀 내부 단락별 세부 복사 진행
                    new_cell.text = "" # 기본 생성된 텍스트 청소
                    for p_cell in cell.paragraphs:
                        new_p_cell = new_cell.add_paragraph()
                        for run in p_cell.runs:
                            new_run = new_p_cell.add_run(run.text)
                            new_run.bold = run.bold
                            new_run.italic = run.italic
                            if run._r.xpath('.//w:drawing'):
                                new_p_cell.runs[-1]._r.append(run._r.xpath('.//w:drawing')[0])

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 2. 시험지용 ➡️ 일반용 변환 (구조 단순화)
# ==========================================
def convert_exam_to_general_final(source_doc):
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
st.markdown("디자인 서식 및 **표 내부 지문 데이터**를 100% 무결하게 유지하며 변환합니다.")

# 서버 내 템플릿 탐지 상태 가이드 제공
if not os.path.exists("template_exam.docx"):
    st.error("🚨 알림: 현재 폴더에 `template_exam.docx`(디자인 템플릿) 파일이 보이지 않습니다. 일반용 문서에 시험지 레이아웃을 입히려면 반드시 템플릿 파일을 같은 레포지토리에 넣어주셔야 합니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    sample_text = "\n".join([p.text for p in doc.paragraphs[:15] if p.text.strip()])
    
    if "[Chapter" in sample_text:
        default_index = 0
        detected_text = "🔍 **문서 양식 감지 결과:** [일반용 양식]이 확인되었습니다. 시험지용 디자인 템플릿 프레임에 맞춰 변환합니다."
    else:
        default_index = 1
        detected_text = "🔍 **문서 양식 감지 결과:** [시험지용 양식]이 확인되었습니다. [일반용] 양식으로 변환합니다."
        
    st.info(detected_text)
    
    mode = st.radio(
        "변환 방향을 확인하거나 변경하세요:",
        ("일반용 ➡️ 시험지용", "시험지용 ➡️ 일반용"),
        index=default_index
    )
    
    if st.button("🚀 서식 보존 변환 시작", use_container_width=True):
        with st.spinner("표 내부 데이터 및 서식 구조를 동기화 중입니다..."):
            try:
                if "일반용 ➡️ 시험지용" in mode:
                    out_bytes = convert_general_to_exam_final(doc)
                    file_name = "변환_시험지용_정기평가.docx"
                else:
                    out_bytes = convert_exam_to_general_final(doc)
                    file_name = "변환_일반용_정기평가.docx"
                
                st.success("🎉 변환 및 정렬 처리가 완료되었습니다!")
                st.download_button(
                    label="💾 변환된 Word 파일 다운로드",
                    data=out_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 변환 처리 중 오류가 발생했습니다: {str(e)}")
