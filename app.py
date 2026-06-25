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
# [유틸리티 함수] 지문 상자 스타일 및 여백/테두리 정의
# ==========================================
def set_cell_margins(cell, top=140, bottom=140, left=150, right=150):
    """지문 상자 내부 여백(패딩)을 설정하여 글자가 답답해 보이지 않게 합니다."""
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
    """지문 상자에 깔끔하고 명확한 검은색 단선 테두리를 입힙니다."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '6')  # 선 두께 설정
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tcBorders.append(border)
    tcPr.append(tcBorders)


# ==========================================
# 1. 일반용 ➡️ 시험지용 변환 (최종 통합본)
# ==========================================
def convert_general_to_exam_integrated(source_doc):
    """
    템플릿의 분홍 배너와 2단 레이아웃을 가져와 본문을 주입합니다.
    영어 지문 상자(표)가 대형 화면 기준으로 고정되어 잘리는 현상을 
    2단 가로폭 강제 매핑 및 자동 줄바꿈(Autofit) 기술로 완벽하게 해결합니다.
    """
    template_path = "template_exam.docx"
    
    if os.path.exists(template_path):
        target_doc = Document(template_path)
    else:
        st.error("🚨 서버에 'template_exam.docx' 파일이 없습니다! GitHub 레포지토리에 파일이 있어야 디자인이 적용됩니다.")
        target_doc = Document()
        target_doc.add_paragraph("2026년 봄학기 THE OPEN 정기평가 (Level P)")
        target_doc.add_paragraph("[ PartⅡ 문법 | 20문항 20분 ]\n")

    q_counter = 1
    
    # 원본 문서의 body 요소를 순서대로 정밀 순회
    for element in source_doc.element.body:
        
        # [단락 처리: 발문, 선지, 이미지 등]
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            # 메타데이터 패스
            if not p_text or p_text.startswith("[Chapter"):
                continue
                
            # 테이블 가이드 안내 텍스트 스킵 처리
            if "The following table:" in p_text:
                continue
            
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_before = p.paragraph_format.space_before
            new_p.paragraph_format.space_after = p.paragraph_format.space_after
            new_p.paragraph_format.line_spacing = p.paragraph_format.line_spacing
            
            # 문항 번호 중복(1. 1.) 원천 차단 및 정렬
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
                        if run.font.size: new_run.font.size = run.font.size
                        if run.font.name: new_run.font.name = run.font.name
                        if run._r.xpath('.//w:drawing'):
                            new_run._r.append(run._r.xpath('.//w:drawing')[0])
            else:
                # 일반 선지 및 보기 텍스트 이식
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                    if run.font.size: new_run.font.size = run.font.size
                    if run.font.name: new_run.font.name = run.font.name
                    if run._r.xpath('.//w:drawing'):
                        new_run._r.append(run._r.xpath('.//w:drawing')[0])
                        
        # [표 처리: 영어 지문 상자 복사 및 가로폭 잘림 방지 극대화]
        elif element.tag.endswith('tbl'):
            src_tbl = docx.table.Table(element, source_doc)
            
            # 새 표 생성 후 자동 너비 맞춤 속성 부여
            dst_tbl = target_doc.add_table(rows=len(src_tbl.rows), cols=len(src_tbl.columns))
            dst_tbl.allow_autofit = True  # 자동 줄바꿈 활성화
            dst_tbl.style = 'Table Grid'   # 기본 틀 스타일 고정
            
            # 내부 셀 데이터 매핑 및 너비 제한 설정
            for r_idx, row in enumerate(src_tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    dst_cell = dst_tbl.cell(r_idx, c_idx)
                    
                    # 💡 핵심 대책: 2단 서식 안에서 표가 늘어나 잘리지 않게 물리적 최대 가로폭 고정 (약 3.1 인치)
                    dst_cell.width = docx.shared.Inches(3.1)
                    
                    # 여백과 테두리 선 새로 빌드
                    set_cell_margins(dst_cell, top=140, bottom=140, left=160, right=160)
                    set_cell_borders(dst_cell)
                    
                    dst_cell.text = ""  # 자동 생성 더미 텍스트 삭제
                    
                    # 표 내부 문단 및 폰트, 밑줄 서식 무결성 이식
                    for p_idx, p_cell in enumerate(cell.paragraphs):
                        if p_idx == 0:
                            dst_p_cell = dst_cell.paragraphs[0]
                        else:
                            dst_p_cell = dst_cell.add_paragraph()
                            
                        dst_p_cell.paragraph_format.space_before = p_cell.paragraph_format.space_before
                        dst_p_cell.paragraph_format.space_after = p_cell.paragraph_format.space_after
                        dst_p_cell.paragraph_format.line_spacing = 1.2  # 보기 좋은 지문 행간 설정
                        
                        for run in p_cell.runs:
                            dst_run = dst_p_cell.add_run(run.text)
                            dst_run.bold = run.bold
                            dst_run.italic = run.italic
                            dst_run.underline = run.underline  # 지문 내 밑줄 완벽 보존
                            if run.font.size: dst_run.font.size = run.font.size
                            if run.font.name: dst_run.font.name = run.font.name
                            if run._r.xpath('.//w:drawing'):
                                dst_p_cell.runs[-1]._r.append(run._r.xpath('.//w:drawing')[0])

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


# ==========================================
# 2. 시험지용 ➡️ 일반용 변환 (순수 데이터 복원)
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
# 3. Streamlit 웹 인터페이스 UI 정의
# ==========================================
st.title("📝 정기평가 양식 상호 변환 시스템")
st.markdown("상단 디자인 배너 보존 및 지문 상자의 **우측 잘림 현상(너비 오버플로우)**을 수정한 최종 통합 버전입니다.")

if not os.path.exists("template_exam.docx"):
    st.error("🚨 [설정 안내] 현재 서버 폴더에 `template_exam.docx` 디자인 배경 틀이 존재하지 않습니다. 일반용 양식을 변경하기 전 반드시 템플릿 파일을 업로드해 주세요.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    sample_text = "\n".join([p.text for p in doc.paragraphs[:15] if p.text.strip()])
    
    # 문서 스타일 자동 감지 및 스마트 기본값 세팅
    if "[Chapter" in sample_text or "The following table:" in sample_text:
        default_index = 0
        detected_text = "🔍 **문서 양식 감지 결과:** [일반용 양식]이 확인되었습니다. 2단 맞춤형 지문 상자를 생성하며 시험지용으로 변환합니다."
    else:
        default_index = 1
        detected_text = "🔍 **문서 양식 감지 결과:** [시험지용 양식]이 확인되었습니다. [일반용] 서식으로 데이터 복원을 수행합니다."
        
    st.info(detected_text)
    
    mode = st.radio(
        "변환 방향을 확인하거나 변경하세요:",
        ("일반용 ➡️ 시험지용", "시험지용 ➡️ 일반용"),
        index=default_index
    )
    
    if st.button("🚀 통합형 양식 변환 및 너비 최적화 시작", use_container_width=True):
        with st.spinner("2단 단락 구조 및 지문 박스 폭 가로 정렬을 제어하는 중입니다..."):
            try:
                if "일반용 ➡️ 시험지용" in mode:
                    out_bytes = convert_general_to_exam_integrated(doc)
                    file_name = "변환_시험지용_정기평가.docx"
                else:
                    out_bytes = convert_exam_to_general_integrated(doc)
                    file_name = "변환_일반용_정기평가.docx"
                
                st.success("🎉 변환 및 지문 상자 가로 가이드 정렬이 모두 완료되었습니다!")
                st.download_button(
                    label="💾 최적화된 Word 파일 다운로드",
                    data=out_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 처리 도중 에러가 발생했습니다: {str(e)}")
