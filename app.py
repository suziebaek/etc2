import streamlit as st
import docx
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor, Inches, Pt
import io
import re
import os

st.set_page_config(
    page_title="정기평가 양식 상호 변환 시스템",
    page_icon="📝",
    layout="centered"
)

# ==========================================
# [유틸리티 함수] 글꼴, 색상, 표 너비 제어
# ==========================================
def apply_custom_style(run, font_name="Noto Sans KR", color_rgb=None):
    """
    텍스트 Run에 Noto Sans KR 글꼴(한글 깨짐 방지 XML 포함) 및 
    지정된 색상(정답용 빨간색 등)을 정밀하게 적용합니다.
    """
    run.font.name = font_name
    if color_rgb:
        run.font.color.rgb = color_rgb
        
    # 영문/한글 글꼴 설정을 위한 워드 내부 XML 제어
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)

def should_skip_paragraph(text):
    """제거해야 할 메타데이터 태그(Collocation 등)를 완벽하게 걸러냅니다."""
    t = text.strip()
    if not t or "The following table:" in t:
        return True
    
    low_t = t.lower()
    keywords = [
        "chapter", "collocation", "word arrangement", "fill in the blank", 
        "교재 연계", "객관-간접", "sentence transformation", "correct sentence", "pg "
    ]
    for kw in keywords:
        if kw in low_t:
            return True
            
    return False

def set_table_width_to_column(table):
    """표가 2단 레이아웃의 단 너비에 한 치의 오차도 없이 딱 맞추어 들어가도록 고정합니다."""
    table.allow_autofit = False
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)
        
    for child in list(tblPr):
        if child.tag.endswith('tblW'):
            tblPr.remove(child)
            
    # 단 너비의 100%를 채우도록 설정 (5000 pct)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    tblPr.append(tblW)

def set_cell_properties(cell):
    """셀 내부 여백과 검은색 단선 테두리를 깔끔하게 처리합니다."""
    tcPr = cell._tc.get_or_add_tcPr()
    
    # 셀 너비 속성 초기화 후 단 너비 연동
    for child in list(tcPr):
        if child.tag.endswith('tcW'):
            tcPr.remove(child)
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), '5000')
    tcW.set(qn('w:type'), 'pct')
    tcPr.append(tcW)
    
    # 내부 여백 패딩 설정 (상하좌우 미세 조정)
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', 140), ('bottom', 140), ('left', 160), ('right', 160)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

    # 검은색 단선 테두리 설정
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
# 1. 일반용 ➡️ 시험지용 변환 핵심 엔진
# ==========================================
def convert_general_to_exam_integrated(source_doc):
    template_path = "template_exam.docx"
    target_doc = Document(template_path) if os.path.exists(template_path) else Document()
    
    q_counter = 1
    red_color = RGBColor(255, 0, 0) # 정답 표기용 빨간색 크로마 수치
    
    for element in source_doc.element.body:
        # [1] 일반 문단 및 보기 선지 처리
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            if should_skip_paragraph(p_text):
                continue
            
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_before = p.paragraph_format.space_before
            new_p.paragraph_format.space_after = p.paragraph_format.space_after
            
            # 정답 검출 로직 (문단이 '정답:'으로 시작하면 빨간색 글씨 처리)
            is_answer = p_text.startswith("정답:") or "정답:" in p_text
            current_color = red_color if is_answer else None
            
            if re.match(r'^\d+\.', p_text):
                clean_p_text = re.sub(r'^\d+\.\s*', '', p_text)
                run_num = new_p.add_run(f"{q_counter}.  ")
                run_num.bold = True
                apply_custom_style(run_num, font_name="Noto Sans KR", color_rgb=current_color)
                q_counter += 1
                
                for run in p.runs:
                    run_clean = re.sub(r'^\d+\.\s*', '', run.text) if run.text.strip().startswith(p_text[:2]) else run.text
                    if run_clean:
                        new_run = new_p.add_run(run_clean)
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                        new_run.underline = run.underline
                        apply_custom_style(new_run, font_name="Noto Sans KR", color_rgb=current_color)
            else:
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                    apply_custom_style(new_run, font_name="Noto Sans KR", color_rgb=current_color)
                        
        # [2] 지문 표(Table) 처리 - 단폭 맞춤 기술 고도화 적용
        elif element.tag.endswith('tbl'):
            src_tbl = docx.table.Table(element, source_doc)
            
            # 유효한 알맹이 지문이 들어있는지 사전 체크
            has_valid_content = False
            for row in src_tbl.rows:
                for cell in row.cells:
                    for p_cell in cell.paragraphs:
                        if p_cell.text.strip() and not should_skip_paragraph(p_cell.text):
                            has_valid_content = True
            
            if not has_valid_content:
                continue
            
            dst_tbl = target_doc.add_table(rows=len(src_tbl.rows), cols=len(src_tbl.columns))
            dst_tbl.style = 'Table Grid'
            set_table_width_to_column(dst_tbl) # 표 너비를 단폭에 강제 동기화
            
            for r_idx, row in enumerate(src_tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    dst_cell = dst_tbl.cell(r_idx, c_idx)
                    set_cell_properties(dst_cell)
                    dst_cell.text = "" 
                    
                    is_first_p = True
                    for p_cell in cell.paragraphs:
                        if should_skip_paragraph(p_cell.text):
                            continue
                            
                        if is_first_p:
                            dst_p_cell = dst_cell.paragraphs[0]
                            is_first_p = False
                        else:
                            dst_p_cell = dst_cell.add_paragraph()
                            
                        dst_p_cell.paragraph_format.line_spacing = 1.2
                        dst_p_cell.paragraph_format.space_after = Pt(2)
                        
                        # 표 내부 정답 코드 하이라이트 검출
                        cell_p_text = p_cell.text.strip()
                        is_cell_answer = cell_p_text.startswith("정답:") or "정답:" in cell_p_text
                        cell_color = red_color if is_cell_answer else None
                        
                        for run in p_cell.runs:
                            dst_run = dst_p_cell.add_run(run.text)
                            dst_run.bold = run.bold
                            dst_run.italic = run.italic
                            dst_run.underline = run.underline
                            apply_custom_style(dst_run, font_name="Noto Sans KR", color_rgb=cell_color)

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 2. 시험지용 ➡️ 일반용 데이터 복원 로직
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
            
        # 복원 시에도 글꼴 속성 재정비
        for run in p.runs:
            apply_custom_style(run, font_name="Noto Sans KR")

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 3. Streamlit 웹 대시보드
# ==========================================
st.title("📝 정기평가 양식 최고 고도화 시스템")
st.markdown("가로 너비 **100% 단 맞춤** / **Noto Sans KR 글꼴 강제 지정** / **정답지 빨간색 하이라이트** 통합 수정본입니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    
    if st.button("🚀 맞춤 디자인 및 정답 하이라이트 변환 시작", use_container_width=True):
        with st.spinner("글꼴 치환 및 지문 상자 가로 폭 비율을 계산 정렬하는 중..."):
            try:
                out_bytes = convert_general_to_exam_integrated(doc)
                st.success("🎉 요구사항이 완벽하게 처리된 최종 파일이 빌드되었습니다!")
                st.download_button(
                    label="💾 완성본 시험지 다운로드",
                    data=out_bytes,
                    file_name="정기평가_최종형_출력문서.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 시스템 오류가 발생했습니다: {str(e)}")
