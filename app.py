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
# 1. 일반용 ➡️ 시험지용 변환 (템플릿 기반 안전 제어)
# ==========================================
def convert_general_to_exam_v3(source_doc):
    template_path = "template_exam.docx"
    
    # 템플릿 파일 존재 여부 확인
    if os.path.exists(template_path):
        target_doc = Document(template_path)
        is_template_used = True
    else:
        target_doc = Document()
        is_template_used = False
        # 템플릿이 없을 때는 임시 헤더 추가
        target_doc.add_paragraph("2026년 봄학기 THE OPEN 정기평가 (Level P)")
        target_doc.add_paragraph("[ PartⅡ 문법 | 20문항 20분 ]\n")

    q_counter = 1
    
    for element in source_doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            if not p_text or p_text.startswith("[Chapter"):
                continue
            
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_before = p.paragraph_format.space_before
            new_p.paragraph_format.space_after = p.paragraph_format.space_after
            
            if re.match(r'^\d+\.', p_text):
                clean_text = re.sub(r'^\d+\.\s*', '', p_text)
                run_num = new_p.add_run(f"{q_counter}.  ")
                run_num.bold = True
                q_counter += 1
                
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    if run._r.xpath('.//w:drawing'):
                        new_run._r.append(run._r.xpath('.//w:drawing')[0])
            else:
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    if run._r.xpath('.//w:drawing'):
                        new_run._r.append(run._r.xpath('.//w:drawing')[0])
                        
        elif element.tag.endswith('tbl'):
            tbl = docx.table.Table(element, source_doc)
            new_tbl = target_doc.add_table(rows=len(tbl.rows), cols=len(tbl.columns))
            new_tbl.style = tbl.style
            for r_idx, row in enumerate(tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    new_cell = new_tbl.cell(r_idx, c_idx)
                    new_cell.text = cell.text
                    for p_cell in cell.paragraphs:
                        for run in p_cell.runs:
                            if run._r.xpath('.//w:drawing'):
                                new_cell.paragraphs[0].add_run()._r.append(run._r.xpath('.//w:drawing')[0])

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue(), is_template_used

# ==========================================
# 2. 시험지용 ➡️ 일반용 변환
# ==========================================
def convert_exam_to_general_v3(source_doc):
    doc_stream = io.BytesIO()
    source_doc.save(doc_stream)
    doc_stream.seek(0)
    target_doc = Document(doc_stream)
    
    q_counter = 1
    for p in target_doc.paragraphs:
        p_text = p.text.strip()
        if re.match(r'^\d+\.', p_text):
            clean_text = re.sub(r'^\d+\.\s*', '', p.text)
            p.text = f"{q_counter}.  {clean_text}"
            q_counter += 1

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()

# ==========================================
# 3. Streamlit 웹 인터페이스 UI
# ==========================================
st.title("📝 정기평가 양식 상호 변환 시스템")
st.markdown("디자인 서식(2단 분할, 상단 디자인)을 유지하며 변환을 진행합니다.")

# 서버에 템플릿 파일이 존재하는지 화면에 표시 (체크용)
if not os.path.exists("template_exam.docx"):
    st.warning("⚠️ 현재 서버에 `template_exam.docx`(시험지 디자인 템플릿) 파일이 업로드되지 않았습니다. 이 상태로 일반용을 넣으시면 디자인 레이아웃이 적용되지 않습니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    sample_text = "\n".join([p.text for p in doc.paragraphs[:15] if p.text.strip()])
    
    if "[Chapter" in sample_text:
        default_index = 0
        detected_text = "🔍 **문서 양식 감지 결과:** [일반용 양식]이 확인되었습니다."
    else:
        default_index = 1
        detected_text = "🔍 **문서 양식 감지 결과:** [시험지용 양식]이 확인되었습니다."
        
    st.info(detected_text)
    
    mode = st.radio(
        "변환 방향을 확인하거나 변경하세요:",
        ("일반용 ➡️ 시험지용", "시험지용 ➡️ 일반용"),
        index=default_index
    )
    
    if st.button("🚀 정밀 양식 변환 시작", use_container_width=True):
        with st.spinner("구조를 분석하여 변환 중입니다..."):
            try:
                if "일반용 ➡️ 시험지용" in mode:
                    out_bytes, used = convert_general_to_exam_v3(doc)
                    file_name = "변환_시험지용_정기평가.docx"
                    if not used:
                        st.error("⚠️ 디자인 템플릿 파일이 없어 기본 텍스트 포맷으로 다운로드됩니다. 레포지토리에 template_exam.docx를 추가해 주세요.")
                else:
                    out_bytes = convert_exam_to_general_v3(doc)
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
                st.error(f"⚠️ 변환 중 오류가 발생했습니다: {str(e)}")
