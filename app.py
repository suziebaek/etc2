import streamlit as st
import docx
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re
import os

st.set_page_config(page_title="정기평가 양식 변환기", page_icon="📝", layout="centered")

# ==========================================
# [유틸리티 함수] 지문 상자 서식 및 너비 강제 고정
# ==========================================
def set_table_width_100_percent(table):
    """표가 단(Column)의 가로폭을 넘지 못하도록 너비를 100%로 강제 매핑합니다."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    
    # 기존 너비 설정 제거 후 100% (5000 pct) 설정
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
    """지문 상자 검은색 단선 테두리 설정"""
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
# 1. 일반용 ➡️ 시험지용 변환 로직
# ==========================================
def convert_general_to_exam_integrated(source_doc):
    template_path = "template_exam.docx"
    target_doc = Document(template_path) if os.path.exists(template_path) else Document()
    
    q_counter = 1
    
    for element in source_doc.element.body:
        # [단락 처리]
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            # 💡 [필터링 추가] 빈 줄, Chapter, Collocation, pg 등 불필요한 메타데이터 태그 삭제
            if not p_text:
                continue
            if re.match(r'^\[.*(?:Chapter|pg|교재|객관|Collocation|Word Arrangement|Fill in the blank).*\]$', p_text):
                continue
            if "The following table:" in p_text:
                continue
            
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_before = p.paragraph_format.space_before
            new_p.paragraph_format.space_after = p.paragraph_format.space_after
            
            # 문항 번호 정형화
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
                        
        # [표 처리: 가로폭 짤림 원천 차단]
        elif element.tag.endswith('tbl'):
            src_tbl = docx.table.Table(element, source_doc)
            
            dst_tbl = target_doc.add_table(rows=len(src_tbl.rows), cols=len(src_tbl.columns))
            dst_tbl.style = 'Table Grid'
            
            # 💡 [핵심 해결책] 표 너비를 2단 레이아웃의 단 너비(100%)에 딱 맞게 강제 제한
            set_table_width_100_percent(dst_tbl)
            
            for r_idx, row in enumerate(src_tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    dst_cell = dst_tbl.cell(r_idx, c_idx)
                    
                    set_cell_margins(dst_cell)
                    set_cell_borders(dst_cell)
                    
                    dst_cell.text = "" 
                    
                    for p_idx, p_cell in enumerate(cell.paragraphs):
                        if p_idx == 0:
                            dst_p_cell = dst_cell.paragraphs[0]
                        else:
                            dst_p_cell = dst_cell.add_paragraph()
                            
                        dst_p_cell.paragraph_format.line_spacing = 1.2
                        
                        for run in p_cell.runs:
                            dst_run = dst_p_cell.add_run(run.text)
                            dst_run.bold = run.bold
                            dst_run.italic = run.italic
                            dst_run.underline = run.underline

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()

# ==========================================
# Streamlit UI
# ==========================================
st.title("📝 정기평가 양식 상호 변환 시스템")
st.markdown("표 너비 짤림 현상(100% 강제 맞춤) 및 불필요한 태그(Collocation 등) 제거가 적용된 버전입니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    
    if st.button("🚀 최적화 변환 시작 (일반용 ➡️ 시험지용)", use_container_width=True):
        with st.spinner("단락을 다듬고 지문 상자의 너비를 맞추는 중입니다..."):
            out_bytes = convert_general_to_exam_integrated(doc)
            st.success("🎉 변환 완료! 표 짤림과 불필요한 태그가 모두 제거되었습니다.")
            st.download_button(
                label="💾 완성된 시험지 다운로드",
                data=out_bytes,
                file_name="변환완료_시험지.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
