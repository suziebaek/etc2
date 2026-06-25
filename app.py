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
# 1. 일반용 ➡️ 시험지용 변환 (디자인 + 지문 박스 테두리 보존형)
# ==========================================
def convert_general_to_exam_perfect_v2(source_doc):
    """
    템플릿(template_exam.docx)의 상단 분홍색 레이아웃과 2단 다단을 그대로 가져옵니다.
    일반용 문서의 본문을 주입하되, 지문을 감싸는 네모 상자(Table)의 테두리와 내부 서식을 완벽하게 복제합니다.
    """
    template_path = "template_exam.docx"
    
    if os.path.exists(template_path):
        target_doc = Document(template_path)
    else:
        st.error("🚨 서버에 'template_exam.docx' 파일이 없습니다! GitHub 레포지토리에 템플릿 파일을 업로드해주세요.")
        target_doc = Document()
        target_doc.add_paragraph("2026년 봄학기 THE OPEN 정기평가 (Level P)")
        target_doc.add_paragraph("[ PartⅡ 문법 | 20문항 20분 ]\n")

    q_counter = 1
    
    # 원본 문서의 구조적 바디 요소를 그대로 추적
    for element in source_doc.element.body:
        
        # [1] 일반 단락 처리 (발문, 선지, 이미지 등)
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            if not p_text or p_text.startswith("[Chapter"):
                continue
            
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_before = p.paragraph_format.space_before
            new_p.paragraph_format.space_after = p.paragraph_format.space_after
            new_p.paragraph_format.line_spacing = p.paragraph_format.line_spacing
            
            # 문항 번호 정렬 및 중복 생성(1. 1.) 제어
            if re.match(r'^\d+\.', p_text):
                clean_p_text = re.sub(r'^\d+\.\s*', '', p_text)
                
                run_num = new_p.add_run(f"{q_counter}.  ")
                run_num.bold = True
                q_counter += 1
                
                for run in p.runs:
                    run_clean = re.sub(r'^\d+\.\s*', '', run.text) if run.text.strip().startswith(p_text[:2]) else run.text
                    if run_clean:
                        new_run = new_p.add_run(run_clean)
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                        new_run.font.size = run.font.size
                        new_run.font.name = run.font.name
                        if run._r.xpath('.//w:drawing'):
                            new_run._r.append(run._r.xpath('.//w:drawing')[0])
            else:
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.font.size = run.font.size
                    new_run.font.name = run.font.name
                    if run._r.xpath('.//w:drawing'):
                        new_run._r.append(run._r.xpath('.//w:drawing')[0])
                        
        # [2] 지문 상자 처리 (Table 구조 발견 시 테두리 보존 매핑)
        elif element.tag.endswith('tbl'):
            src_tbl = docx.table.Table(element, source_doc)
            
            # 새 문서에 표를 생성할 때 원본 표의 XML 서식 속성(tblPr - 테두리 선 정보 포함)을 그대로 주입
            dst_tbl = target_doc.add_table(rows=len(src_tbl.rows), cols=len(src_tbl.columns))
            dst_tbl.style = src_tbl.style
            
            # 핵심: 원본 표가 가지고 있는 테두리, 배경색 등의 XML 스타일 속성을 그대로 복사해 오기
            try:
                dst_tbl._tbl.tblPr = src_tbl._tbl.tblPr
            except Exception:
                pass # 예외 발생 시 기본 서식 스타일 유지
            
            # 표 내부 셀 및 텍스트 데이터의 다중 문단(Paragraph) 서식 정밀 복사
            for r_idx, row in enumerate(src_tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    dst_cell = dst_tbl.cell(r_idx, c_idx)
                    
                    # 원본 셀의 속성(너비, 패딩 등) 복사
                    try:
                        dst_cell._tc.tcPr = cell._tc.tcPr
                    except Exception:
                        pass
                        
                    dst_cell.text = "" # 기본 자동 생성 공백 제거
                    
                    # 셀 안의 문단별로 서식 분할 복사 (밑줄, 폰트, 기호 유지)
                    for idx, p_cell in enumerate(cell.paragraphs):
                        if idx == 0:
                            dst_p_cell = dst_cell.paragraphs[0]
                        else:
                            dst_p_cell = dst_cell.add_paragraph()
                            
                        dst_p_cell.paragraph_format.space_before = p_cell.paragraph_format.space_before
                        dst_p_cell.paragraph_format.space_after = p_cell.paragraph_format.space_after
                        dst_p_cell.paragraph_format.alignment = p_cell.paragraph_format.alignment
                        
                        for run in p_cell.runs:
                            dst_run = dst_p_cell.add_run(run.text)
                            dst_run.bold = run.bold
                            dst_run.italic = run.italic
                            dst_run.underline = run.underline # 밑줄 속성 복사 추가
                            dst_run.font.size = run.font.size
                            dst_run.font.name = run.font.name
                            if run._r.xpath('.//w:drawing'):
                                dst_p_cell.runs[-1]._r.append(run._r.xpath('.//w:drawing')[0])

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 2. 시험지용 ➡️ 일반용 변환
# ==========================================
def convert_exam_to_general_perfect_v2(source_doc):
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
st.markdown("디자인 서식과 **지문 테두리 박스(네모 상자) 속성**을 완벽하게 병합하여 변환합니다.")

if not os.path.exists("template_exam.docx"):
    st.error("🚨 [설정 에러] 현재 폴더에 `template_exam.docx` 디자인 템플릿 파일이 보이지 않습니다. 시험지 레이아웃을 입히려면 반드시 배경 틀 파일이 GitHub 같은 폴더에 업로드되어 있어야 합니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    sample_text = "\n".join([p.text for p in doc.paragraphs[:15] if p.text.strip()])
    
    if "[Chapter" in sample_text:
        default_index = 0
        detected_text = "🔍 **문서 양식 감지 결과:** [일반용 양식]이 확인되었습니다. 디자인 틀과 지문 네모 상자를 모두 유지하며 시험지용으로 변환합니다."
    else:
        default_index = 1
        detected_text = "🔍 **문서 양식 감지 결과:** [시험지용 양식]이 확인되었습니다. [일반용] 양식으로 변환합니다."
        
    st.info(detected_text)
    
    mode = st.radio(
        "변환 방향을 확인하거나 변경하세요:",
        ("일반용 ➡️ 시험지용", "시험지용 ➡️ 일반용"),
        index=default_index
    )
    
    if st.button("🚀 디자인 및 지문 박스 복원 변환 시작", use_container_width=True):
        with st.spinner("지문 상자의 테두리 및 텍스트 서식을 동기화하고 있습니다..."):
            try:
                if "일반용 ➡️ 시험지용" in mode:
                    out_bytes = convert_general_to_exam_perfect_v2(doc)
                    file_name = "변환_시험지용_정기평가.docx"
                else:
                    out_bytes = convert_exam_to_general_perfect_v2(doc)
                    file_name = "변환_일반용_정기평가.docx"
                
                st.success("🎉 변환 처리가 완료되었습니다!")
                st.download_button(
                    label="💾 변환된 Word 파일 다운로드",
                    data=out_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 변환 처리 중 오류가 발생했습니다: {str(e)}")
