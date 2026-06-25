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

# ==========================================
# [유틸리티 함수] 메타데이터 필터링 및 표 서식 제어
# ==========================================
def should_skip_paragraph(text):
    """
    본문과 표 내부를 막론하고, 제거해야 할 메타데이터 태그(Collocation 등)를 
    형식과 관계없이 키워드 기반으로 완벽하게 감지하여 거릅니다.
    """
    t = text.strip()
    if not t:
        return True
    
    # 제외할 핵심 키워드 목록 (대소문자 무시)
    low_t = t.lower()
    keywords = [
        "chapter", "collocation", "word arrangement", "fill in the blank", 
        "교재 연계", "객관-간접", "sentence transformation", "correct sentence", "pg "
    ]
    
    for kw in keywords:
        if kw in low_t:
            return True
            
    return False

def set_table_width_100_percent(table):
    """표가 2단 레이아웃의 단 너비(100%)를 절대 넘지 않도록 강제 제한합니다."""
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)
        
    # 기존 고정 너비 속성이 있다면 모두 제거
    for child in list(tblPr):
        if child.tag.endswith('tblW'):
            tblPr.remove(child)
            
    # 부모 컨테이너 너비의 100%를 뜻하는 5000 pct 속성 주입
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    tblPr.append(tblW)

def set_cell_margins(cell, top=140, bottom=140, left=160, right=160):
    """지문 상자 내부 여백(패딩) 설정"""
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
    """지문 상자 검은색 단선 테두리 명확하게 설정"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '6') 
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tcBorders.append(border)
    tcPr.append(tcBorders)


# ==========================================
# 1. 일반용 ➡️ 시험지용 변환 로직 (최종 개선본)
# ==========================================
def convert_general_to_exam_integrated(source_doc):
    template_path = "template_exam.docx"
    target_doc = Document(template_path) if os.path.exists(template_path) else Document()
    
    q_counter = 1
    
    for element in source_doc.element.body:
        # [1] 일반 단락 처리
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            # 메타데이터 및 불필요 항목 필터링
            if should_skip_paragraph(p_text):
                continue
            if "The following table:" in p_text:
                continue
            
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_before = p.paragraph_format.space_before
            new_p.paragraph_format.space_after = p.paragraph_format.space_after
            
            # 문항 번호 자동 정형화 (중복 방지)
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
                        new_run.underline = run.underline
            else:
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                        
        # [2] 표 처리 (가로폭 단 맞춤 및 표 내부 Collocation 제거)
        elif element.tag.endswith('tbl'):
            src_tbl = docx.table.Table(element, source_doc)
            
            # 표 내부에 필터링 되지 않는 '진짜 유효한 지문'이 존재하는지 사전 검사
            has_valid_content = False
            for row in src_tbl.rows:
                for cell in row.cells:
                    for p_cell in cell.paragraphs:
                        if p_cell.text.strip() and not should_skip_paragraph(p_cell.text):
                            has_valid_content = True
                            break
            
            # 만약 표 내부에 Collocation 같은 태그만 들어있고 실속이 없다면 표 자체를 패스
            if not has_valid_content:
                continue
            
            dst_tbl = target_doc.add_table(rows=len(src_tbl.rows), cols=len(src_tbl.columns))
            dst_tbl.style = 'Table Grid'
            
            # 💡 대책 1: 표 전체의 너비를 단 가로폭의 100%로 완전 고정
            set_table_width_100_percent(dst_tbl)
            
            for r_idx, row in enumerate(src_tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    dst_cell = dst_tbl.cell(r_idx, c_idx)
                    
                    set_cell_margins(dst_cell)
                    set_cell_borders(dst_cell)
                    
                    # 💡 대책 2: 셀 내부 고정 너비 속성도 초기화 후 100% pct로 강제 일치
                    tcPr = dst_cell._tc.get_or_add_tcPr()
                    for child in list(tcPr):
                        if child.tag.endswith('tcW'):
                            tcPr.remove(child)
                    tcW = OxmlElement('w:tcW')
                    tcW.set(qn('w:w'), '5000')
                    tcW.set(qn('w:type'), 'pct')
                    tcPr.append(tcW)
                    
                    dst_cell.text = "" 
                    
                    is_first_p = True
                    for p_cell in cell.paragraphs:
                        # 💡 대책 3: 표 안의 문단들도 하나씩 검사하여 Collocation 태그 라인은 삽입 안 함
                        if should_skip_paragraph(p_cell.text):
                            continue
                            
                        if is_first_p:
                            dst_p_cell = dst_cell.paragraphs[0]
                            is_first_p = False
                        else:
                            dst_p_cell = dst_cell.add_paragraph()
                            
                        dst_p_cell.paragraph_format.line_spacing = 1.2
                        dst_p_cell.paragraph_format.space_after = docx.shared.Pt(2)
                        
                        for run in p_cell.runs:
                            dst_run = dst_p_cell.add_run(run.text)
                            dst_run.bold = run.bold
                            dst_run.italic = run.italic
                            dst_run.underline = run.underline

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 2. 시험지용 ➡️ 일반용 변환 로직
# ==========================================
def convert_exam_to_general_integrated(source_doc):
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
# 3. Streamlit 웹 인터페이스
# ==========================================
st.title("📝 정기평가 양식 상호 변환 시스템")
st.markdown("표 너비를 **시험지 단 길이에 100% 매립**하고, 표 안팎의 **Collocation 태그를 완벽하게 지운** 최종본입니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    
    if st.button("🚀 최종 최적화 변환 시작", use_container_width=True):
        with st.spinner("지문 상자 폭 조정 및 불필요 태그 소거 작업 중..."):
            try:
                out_bytes = convert_general_to_exam_integrated(doc)
                st.success("🎉 변환이 완벽하게 완료되었습니다!")
                st.download_button(
                    label="💾 완성된 시험지 다운로드",
                    data=out_bytes,
                    file_name="최종최적화_시험지.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 에러 발생: {str(e)}")
