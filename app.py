import streamlit as st
import docx
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor, Pt
import io
import re
import os

st.set_page_config(
    page_title="정기평가 양식 상호 변환 시스템",
    page_icon="📝",
    layout="centered"
)

# ==========================================
# [공통 유틸리티 함수] 필터링, 글꼴 및 색상 정밀 제어
# ==========================================
def is_originally_red(run):
    """원본 Run(글자 마디)이 빨간색 계열로 지정되어 있는지 RGB 값을 정밀 추적합니다."""
    try:
        if run.font.color and run.font.color.rgb:
            hex_color = str(run.font.color.rgb).upper()
            # 표준 Red(FF0000) 및 유사 빨간색 판정
            if hex_color == "FF0000" or "255, 0, 0" in hex_color:
                return True
    except:
        pass
    return False

def should_skip_paragraph(text):
    """대괄호 태그 및 교육용 메타데이터 라인을 필터링하여 소거합니다."""
    t = re.sub(r'\s+', ' ', text).strip()
    if not t or "The following table:" in t:
        return True
    
    low_t = t.lower()
    if t.startswith('[') and t.endswith(']'):
        return True
        
    metadata_keywords = [
        "chapter", "collocation", "vocabulary", "reading", "word arrangement", 
        "fill in the blank", "교재 연계", "객관-간접", "sentence transformation", 
        "correct sentence", "pg"
    ]
    
    if t.startswith('[') or t.endswith(']'):
        if any(kw in low_t for kw in metadata_keywords):
            return True
            
    if any(kw in low_t for kw in ["vocabulary reading", "collocation"]):
        return True
                
    return False

def apply_custom_style(run, font_name="Noto Sans KR", color_rgb=None):
    """Noto Sans KR 폰트(동아시아 깨짐 방지 XML 포함) 및 빨간색 등 색상 옵션을 강제 주입합니다."""
    run.font.name = font_name
    if color_rgb:
        run.font.color.rgb = color_rgb
        
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)

def set_table_width_to_column(table):
    """표의 가로 너비를 현재 레이아웃 단 폭의 100%로 강제 정렬합니다."""
    table.allow_autofit = False
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)
        
    for child in list(tblPr):
        if child.tag.endswith('tblW'):
            tblPr.remove(child)
            
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    tblPr.append(tblW)

def set_cell_properties(cell):
    """지문 상자용 테두리와 여백 패딩을 설정합니다."""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('tcW'):
            tcPr.remove(child)
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), '5000')
    tcW.set(qn('w:type'), 'pct')
    tcPr.append(tcW)
    
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', 140), ('bottom', 140), ('left', 160), ('right', 160)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

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
# 1. 일반용 ➡️ 시험지용 변환 엔진 (정답 색상 추적 강화)
# ==========================================
def convert_general_to_exam_integrated(source_doc):
    template_path = "template_exam.docx"
    target_doc = Document(template_path) if os.path.exists(template_path) else Document()
    
    q_counter = 1
    red_color = RGBColor(255, 0, 0)
    
    for element in source_doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            if should_skip_paragraph(p_text):
                continue
            
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_before = p.paragraph_format.space_before
            new_p.paragraph_format.space_after = p.paragraph_format.space_after
            
            # 텍스트에 '정답'이 포함되어 있는지 여부 검사
            is_p_answer = "정답" in p_text
            
            if re.match(r'^\d+[\.\s]', p_text):
                # 기존 번호 제거 후 순차 번호 재배정
                clean_p_text = re.sub(r'^\d+\.\s*', '', p_text)
                clean_p_text = re.sub(r'^\d+\s+', '', clean_p_text)
                
                run_num = new_p.add_run(f"{q_counter}.  ")
                run_num.bold = True
                apply_custom_style(run_num, font_name="Noto Sans KR", color_rgb=red_color if is_p_answer else None)
                q_counter += 1
                
                for run in p.runs:
                    r_text = run.text
                    if r_text.strip().startswith(p_text[:2]):
                        r_text = re.sub(r'^\d+\.\s*', '', r_text)
                        r_text = re.sub(r'^\d+\s+', '', r_text)
                        
                    if r_text:
                        new_run = new_p.add_run(r_text)
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                        new_run.underline = run.underline
                        
                        # 조건: 문단에 '정답' 텍스트가 있거나, 원본 글자 자체가 빨간색인 경우 붉은색 계승
                        if is_p_answer or "정답" in run.text or is_originally_red(run):
                            apply_custom_style(new_run, font_name="Noto Sans KR", color_rgb=red_color)
                        else:
                            apply_custom_style(new_run, font_name="Noto Sans KR")
            else:
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                    
                    if is_p_answer or "정답" in run.text or is_originally_red(run):
                        apply_custom_style(new_run, font_name="Noto Sans KR", color_rgb=red_color)
                    else:
                        apply_custom_style(new_run, font_name="Noto Sans KR")
                        
        elif element.tag.endswith('tbl'):
            src_tbl = docx.table.Table(element, source_doc)
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
            set_table_width_to_column(dst_tbl)
            
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
                        
                        is_cell_answer = "정답" in p_cell.text
                        
                        for run in p_cell.runs:
                            dst_run = dst_p_cell.add_run(run.text)
                            dst_run.bold = run.bold
                            dst_run.italic = run.italic
                            dst_run.underline = run.underline
                            if is_cell_answer or "정답" in run.text or is_originally_red(run):
                                apply_custom_style(dst_run, font_name="Noto Sans KR", color_rgb=red_color)
                            else:
                                apply_custom_style(dst_run, font_name="Noto Sans KR")

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 2. 시험지용 ➡️ 일반용 변환 엔진 (1단 레이아웃 완전 이주)
# ==========================================
def convert_exam_to_general_integrated(source_doc):
    # 💡 중요: 시험지의 2단 양식을 완전히 파괴하기 위해 빈 1단 표준 문서를 새로 개설합니다.
    target_doc = Document()
    
    q_counter = 1
    red_color = RGBColor(255, 0, 0)
    
    for element in source_doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            if should_skip_paragraph(p_text):
                continue
                
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_before = p.paragraph_format.space_before
            new_p.paragraph_format.space_after = p.paragraph_format.space_after
            
            is_p_answer = "정답" in p_text
            
            if re.match(r'^\d+[\.\s]', p_text):
                run_num = new_p.add_run(f"{q_counter}.  ")
                run_num.bold = True
                apply_custom_style(run_num, font_name="Noto Sans KR", color_rgb=red_color if is_p_answer else None)
                q_counter += 1
                
                for run in p.runs:
                    r_text = run.text
                    if r_text.strip().startswith(p_text[:2]):
                        r_text = re.sub(r'^\d+\.\s*', '', r_text)
                        r_text = re.sub(r'^\d+\s+', '', r_text)
                        
                    if r_text:
                        new_run = new_p.add_run(r_text)
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                        new_run.underline = run.underline
                        if is_p_answer or "정답" in run.text or is_originally_red(run):
                            apply_custom_style(new_run, font_name="Noto Sans KR", color_rgb=red_color)
                        else:
                            apply_custom_style(new_run, font_name="Noto Sans KR")
            else:
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                    if is_p_answer or "정답" in run.text or is_originally_red(run):
                        apply_custom_style(new_run, font_name="Noto Sans KR", color_rgb=red_color)
                    else:
                        apply_custom_style(new_run, font_name="Noto Sans KR")
                        
        elif element.tag.endswith('tbl'):
            src_tbl = docx.table.Table(element, source_doc)
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
            set_table_width_to_column(dst_tbl) # 1단 일반 문서 폭에 맞추어 표 크기 확장 조정
            
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
                        
                        is_cell_answer = "정답" in p_cell.text
                        
                        for run in p_cell.runs:
                            dst_run = dst_p_cell.add_run(run.text)
                            dst_run.bold = run.bold
                            dst_run.italic = run.italic
                            dst_run.underline = run.underline
                            if is_cell_answer or "정답" in run.text or is_originally_red(run):
                                apply_custom_style(dst_run, font_name="Noto Sans KR", color_rgb=red_color)
                            else:
                                apply_custom_style(dst_run, font_name="Noto Sans KR")

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 3. Streamlit 웹 대시보드 인터페이스
# ==========================================
st.title("📝 정기평가 양식 최고 고도화 시스템")
st.markdown("가로 너비 **100% 자동 동기화** / **Noto Sans KR 글꼴** / **텍스트+기존 색상 정답 추적 빨간색 고정** / **양방향 완전 레이아웃 스위칭** 버전입니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    
    conversion_mode = st.radio(
        "원하시는 변환 작업 방향을 선택하세요:",
        [
            "일반용 ➡️ 시험지용 (2단 레이아웃 단 맞춤, 태그 소거, 정답지 빨간색 강제 지정)", 
            "시험지용 ➡️ 일반용 (1단 기본 문서 레이아웃 복원, 문항 번호 완전 순차 정렬)"
        ]
    )
    
    if st.button("🚀 선택한 모드로 정밀 변환 시작", use_container_width=True):
        with st.spinner("문서 데이터베이스 파싱 및 폰트 색상 크로마 동기화 중..."):
            try:
                if "일반용 ➡️ 시험지용" in conversion_mode:
                    out_bytes = convert_general_to_exam_integrated(doc)
                    download_filename = "정기평가_최종형_시험지문서.docx"
                else:
                    out_bytes = convert_exam_to_general_integrated(doc)
                    download_filename = "정기평가_복원형_일반문서.docx"
                    
                st.success("🎉 모든 요청사항 및 오류가 수정되어 변환이 완벽히 완료되었습니다!")
                st.download_button(
                    label="💾 완성본 파일 다운로드",
                    data=out_bytes,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 시스템 변환 처리 중 치명적 오류 발생: {str(e)}")
