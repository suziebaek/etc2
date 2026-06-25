import streamlit as st
import docx
from docx import Document
import io
import re

st.set_page_config(
    page_title="정기평가 양식 변환기",
    page_icon="📝",
    layout="centered"
)

# ==========================================
# 1. 문서 파싱 및 상호 변환 핵심 로직 (구조 유지형)
# ==========================================

def convert_general_to_exam(source_doc):
    """
    일반용(1번) -> 시험지용(2번) 변환
    [Chapter ...] 메타데이터 단락만 골라서 삭제하고, 문항 번호를 순서대로 재정렬합니다.
    """
    target_doc = Document()
    
    # 시험지용 타이틀 헤더 추가
    target_doc.add_paragraph("2026년 봄학기 THE OPEN 정기평가 (Level P)")
    target_doc.add_paragraph("[ PartⅡ 문법 | 20문항 20분 ]\n")
    
    q_counter = 1
    
    # 원본 문서의 모든 요소(단락, 표 등)를 순회하며 복사
    for element in source_doc.element.body:
        if element.tag.endswith('p'): # 단락일 때
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            if not p_text:
                continue
                
            # [Chapter ...] 로 시작하는 메타데이터 라인은 시험지용에서 제외
            if p_text.startswith("[Chapter"):
                continue
            
            # 문항 시작 부분 처리 (번호 재정렬)
            # 숫자 뒤에 마침표가 오는 패턴(예: 1., 19.)을 찾아 현재 번호로 교체
            if re.match(r'^\d+\.', p_text):
                clean_text = re.sub(r'^\d+\.\s*', '', p_text)
                new_p = target_doc.add_paragraph()
                new_p.paragraph_format.space_before = p.paragraph_format.space_before
                new_p.paragraph_format.space_after = p.paragraph_format.space_after
                
                # 이미지 및 스타일 유지를 위해 run 단위 복사
                run_num = new_p.add_run(f"{q_counter}. ")
                run_num.bold = True
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    # 내부 이미지 xml 구조 복사
                    if run._r.xpath('.//w:drawing'):
                        new_run._r.append(run._r.xpath('.//w:drawing')[0])
                q_counter += 1
            else:
                # 일반 발문, 보기, 선지 등 내용 그대로 복사 (이미지 포함)
                new_p = target_doc.add_paragraph()
                new_p.paragraph_format.space_before = p.paragraph_format.space_before
                new_p.paragraph_format.space_after = p.paragraph_format.space_after
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    if run._r.xpath('.//w:drawing'):
                        new_run._r.append(run._r.xpath('.//w:drawing')[0])
                        
        elif element.tag.endswith('tbl'): # 표(Table) 구조물 발견 시
            # 표 구조를 깨뜨리지 않고 새 문서에 그대로 삽입
            tbl = docx.table.Table(element, source_doc)
            new_tbl = target_doc.add_table(rows=len(tbl.rows), cols=len(tbl.columns))
            new_tbl.style = tbl.style
            
            for r_idx, row in enumerate(tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    new_cell = new_tbl.cell(r_idx, c_idx)
                    new_cell.text = cell.text
                    # 표 내부 단락의 이미지나 서식도 복사
                    for p in cell.paragraphs:
                        for run in p.runs:
                            if run._r.xpath('.//w:drawing'):
                                new_cell.paragraphs[0].add_run()._r.append(run._r.xpath('.//w:drawing')[0])

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()


def convert_exam_to_general(source_doc):
    """
    시험지용(2번) -> 일반용(1번) 변환
    각 문항 위에 가상의 [Chapter 01: 임시 메타데이터] 템플릿을 자동으로 채워 넣어 줍니다.
    """
    target_doc = Document()
    q_counter = 1
    
    for element in source_doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            if not p_text:
                continue
                
            # 상단 헤더 타이틀 제거
            if "정기평가" in p_text or "PartⅡ" in p_text:
                continue
                
            if re.match(r'^\d+\.', p_text):
                # 문항 번호 앞에 메타데이터 양식 삽입
                target_doc.add_paragraph(f"[Chapter {q_counter:02d}: 연계 단원 확인 필요, pg 00, 객관식 유형]")
                
                clean_text = re.sub(r'^\d+\.\s*', '', p_text)
                new_p = target_doc.add_paragraph()
                run_num = new_p.add_run(f"{q_counter}. ")
                run_num.bold = True
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    if run._r.xpath('.//w:drawing'):
                        new_run._r.append(run._r.xpath('.//w:drawing')[0])
                q_counter += 1
            else:
                new_p = target_doc.add_paragraph()
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
                    for p in cell.paragraphs:
                        for run in p.runs:
                            if run._r.xpath('.//w:drawing'):
                                new_cell.paragraphs[0].add_run()._r.append(run._r.xpath('.//w:drawing')[0])

    b_io = io.BytesIO()
    target_doc.save(b_io)
    return b_io.getvalue()

# ==========================================
# 2. Streamlit 웹 인터페이스 UI (용어 변경 반영)
# ==========================================

st.title("📝 정기평가 양식 상호 변환기")
st.markdown("업로드하신 Word 파일의 양식을 분석하여 **일반용** 또는 **시험지용** 문서로 깨짐 없이 상호 변환합니다.")

uploaded_file = st.file_uploader("변환할 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    
    # 문서 상단 5단락 분석하여 자동 인식
    sample_text = "\n".join([p.text for p in doc.paragraphs[:5] if p.text.strip()])
    
    if "[Chapter" in sample_text:
        default_index = 0
        detected_text = "🔍 **문서 양식 감지 결과:** [일반용 양식]이 감지되었습니다."
    else:
        default_index = 1
        detected_text = "🔍 **문서 양식 감지 결과:** [시험지용 양식]이 감지되었습니다."
        
    st.info(detected_text)
    
    mode = st.radio(
        "변환 방향을 확인하거나 선택하세요:",
        ("일반용 ➡️ 시험지용", "시험지용 ➡️ 일반용"),
        index=default_index
    )
    
    if st.button("🚀 변환하기", use_container_width=True):
        with st.spinner("이미지, 표, 문항 구조를 유지하며 안전하게 변환 중입니다..."):
            try:
                if "일반용 ➡️ 시험지용" in mode:
                    out_bytes = convert_general_to_exam(doc)
                    file_name = "변환_시험지용_정기평가.docx"
                else:
                    out_bytes = convert_exam_to_general(doc)
                    file_name = "변환_일반용_정기평가.docx"
                
                st.success("🎉 양식 변환이 성공적으로 완료되었습니다!")
                st.download_button(
                    label="💾 변환된 Word 파일 다운로드",
                    data=out_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 변환 중 오류가 발생했습니다. 파일 형식을 다시 확인해주세요. (에러: {str(e)})")
