import streamlit as st
import docx
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re
import os

st.set_page_config(
    page_title="정기평가 양식 상호 변환 시스템",
    page_icon="📝",
    layout="centered"
)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """지문 상자 내부 여백(패딩)을 주어 글자가 답답해 보이지 않게 합니다."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_borders(cell):
    """지문 상자에 깔끔한 검은색 단선 테두리를 입힙니다."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')  # 선 두께
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000') # 검은색 테두리
        tcBorders.append(border)
    tcPr.append(tcBorders)

# ==========================================
# 1. 일반용 ➡️ 시험지용 변환 (지문 상자 자동 생성형)
# ==========================================
def convert_general_to_exam_forced_box(source_doc):
    """
    template_exam.docx의 디자인 서식을 유지하면서,
    한글 발문과 보기 선지(①~⑤) 사이에 있는 영어 지문을 무조건 '검은색 네모 상자' 안에 넣어서 생성합니다.
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
    
    # 순차 처리를 위해 원본 문서의 모든 텍스트 단락 수집 (표 내부 포함 추출)
    raw_paragraphs = []
    for element in source_doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            if p.text.strip():
                raw_paragraphs.append(p)
        elif element.tag.endswith('tbl'):
            tbl = docx.table.Table(element, source_doc)
            for row in tbl.rows:
                for cell in row.cells:
                    for p_cell in cell.paragraphs:
                        if p_cell.text.strip():
                            raw_paragraphs.append(p_cell)

    # 지문 텍스트를 모으기 위한 상태 변수
    pending_passage_runs = []
    
    for idx, p in enumerate(raw_paragraphs):
        p_text = p.text.strip()
        
        if p_text.startswith("[Chapter"):
            continue
            
        # 1. 새로운 문항의 시작 감지 (예: 1. 다음 빈칸에...)
        if re.match(r'^\d+\.', p_text):
            # 이전에 처리하던 지문 더미가 있다면 상자 없이 일반 출력 (안전장치)
            if pending_passage_runs:
                new_p = target_doc.add_paragraph()
                for r in pending_passage_runs:
                    new_r = new_p.add_run(r.text)
                    new_r.bold = r.bold; new_r.italic = r.italic
                pending_passage_runs = []
                
            clean_p_text = re.sub(r'^\d+\.\s*', '', p_text)
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_before = docx.shared.Pt(10)
            
            run_num = new_p.add_run(f"{q_counter}.  ")
            run_num.bold = True
            q_counter += 1
            
            # 발문 내용 복사
            for run in p.runs:
                run_clean = re.sub(r'^\d+\.\s*', '', run.text) if run.text.strip().startswith(p_text[:2]) else run.text
                if run_clean:
                    new_run = new_p.add_run(run_clean)
                    new_run.bold = run.bold; new_run.italic = run.italic
                    
        # 2. 보기 선지 시작 감지 (①, ②, ③, ④, ⑤ 또는 정답 세션)
        elif re.match(r'^[①②③④⑤]|^정답:', p_text):
            # 선지가 나오기 직전에 쌓여있던 pending_passage_runs가 바로 "발문과 선지 사이의 지문"입니다!
            if pending_passage_runs:
                # 무조건 검은색 테두리 표(1행 1열)를 생성하여 집어넣음
                box_table = target_doc.add_table(rows=1, cols=1)
                box_table.autofit = False
                cell = box_table.cell(0, 0)
                set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
                set_cell_borders(cell)
                
                # 상자 내부 문단 생성 및 지문 내용 채우기
                box_p = cell.paragraphs[0]
                box_p.paragraph_format.space_after = docx.shared.Pt(0)
                box_p.paragraph_format.line_spacing = 1.2
                
                for r in pending_passage_runs:
                    new_r = box_p.add_run(r.text)
                    new_r.bold = r.bold
                    new_r.italic = r.italic
                    new_r.font.size = docx.shared.Pt(10.5)
                    new_r.font.name = '맑은 고딕'
                
                # 지문 박스 뒤 공백 확보
                target_doc.add_paragraph().paragraph_format.space_before = docx.shared.Pt(6)
                pending_passage_runs = []
                
            # 선지 문단 출력
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_after = docx.shared.Pt(3)
            for run in p.runs:
                new_run = new_p.add_run(run.text)
                new_run.bold = run.bold; new_run.italic = run.italic
                
        # 3. 발문도 아니고 선지도 아닌 중간에 낀 본문 = 지문 데이터로 수집
        else:
            # 테이블 가이드 텍스트는 제외
            if "The following table:" in p_text:
                continue
            for run in p.runs:
                pending_passage_runs.append(run)

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 2. 시험지용 ➡️ 일반용 변환
# ==========================================
def convert_exam_to_general_forced_box(source_doc):
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
st.markdown("한글 지시문과 보기 선지 사이에 있는 모든 영어 지문을 **네모 상자(표)안에 강제 삽입**하여 변환합니다.")

if not os.path.exists("template_exam.docx"):
    st.error("🚨 [설정 에러] 현재 폴더에 `template_exam.docx` 디자인 템플릿 파일이 보이지 않습니다. 시험지 레이아웃을 입히려면 반드시 배경 틀 파일이 GitHub 같은 폴더에 업로드되어 있어야 합니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    sample_text = "\n".join([p.text for p in doc.paragraphs[:15] if p.text.strip()])
    
    if "[Chapter" in sample_text:
        default_index = 0
        detected_text = "🔍 **문서 양식 감지 결과:** [일반용 양식]이 확인되었습니다. 디자인 틀 유지 및 지문 상자 강제 생성 모드로 변환합니다."
    else:
        default_index = 1
        detected_text = "🔍 **문서 양식 감지 결과:** [시험지용 양식]이 확인되었습니다. [일반용] 양식으로 변환합니다."
        
    st.info(detected_text)
    
    mode = st.radio(
        "변환 방향을 확인하거나 변경하세요:",
        ("일반용 ➡️ 시험지용", "시험지용 ➡️ 일반용"),
        index=default_index
    )
    
    if st.button("🚀 디자인 및 지문 박스 강제 생성 변환", use_container_width=True):
        with st.spinner("지문을 추출하여 네모 상자 표 서식을 빌드하고 있습니다..."):
            try:
                if "일반용 ➡️ 시험지용" in mode:
                    out_bytes = convert_general_to_exam_forced_box(doc)
                    file_name = "변환_시험지용_정기평가.docx"
                else:
                    out_bytes = convert_exam_to_general_forced_box(doc)
                    file_name = "변환_일반용_정기평가.docx"
                
                st.success("🎉 변환 및 지문 상자 매핑 처리가 완료되었습니다!")
                st.download_button(
                    label="💾 변환된 Word 파일 다운로드",
                    data=out_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 변환 처리 중 오류가 발생했습니다: {str(e)}")
