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

def convert_general_to_exam_v3(source_doc):
    """
    [일반용 ➡️ 시험지용] 변환 핵심 로직
    프로그램에 내장된 '디자인_템플릿.docx'를 기반으로, 일반용 문서의 알맹이만 이식합니다.
    """
    template_path = "template_exam.docx"
    
    if os.path.exists(template_path):
        # 디자인과 다단 서식이 이미 완성되어 있는 템플릿 파일을 로드
        target_doc = Document(template_path)
    else:
        # 만약 템플릿 파일이 없을 경우 에러 메시지 출력 후 빈 문서로 대체
        st.warning("⚠️ 'template_exam.docx'(시험지 디자인 템플릿) 파일을 찾을 수 없어 기본 서식으로 변환합니다.")
        target_doc = Document()
        target_doc.add_paragraph("2026년 봄학기 THE OPEN 정기평가 (Level P)")
        target_doc.add_paragraph("[ PartⅡ 문법 | 20문항 20분 ]\n")

    q_counter = 1
    
    # 일반용 문서의 바디 요소를 순회하며 템플릿 문서에 데이터 주입
    for element in source_doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            if not p_text or p_text.startswith("[Chapter"):
                continue
            
            new_p = target_doc.add_paragraph()
            new_p.paragraph_format.space_before = p.paragraph_format.space_before
            new_p.paragraph_format.space_after = p.paragraph_format.space_after
            
            # 문항 번호 시작점 감지 시 번호 재정렬
            if re.match(r'^\d+\.', p_text):
                clean_text = re.sub(r'^\d+\.\s*', '', p_text)
                run_num = new_p.add_run(f"{q_counter}.  ")
                run_num.bold = True
                q_counter += 1
                
                # 나머지 문항 텍스트 및 서식, 이미지 복사
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    if run._r.xpath('.//w:drawing'):
                        new_run._r.append(run._r.xpath('.//w:drawing')[0])
            else:
                # 보기문, 선지 등 그대로 복사
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    if run._r.xpath('.//w:drawing'):
                        new_run._r.append(run._r.xpath('.//w:drawing')[0])
                        
        elif element.tag.endswith('tbl'):
            # 표 구조물 이식
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
    return b_io.getvalue()

# (이하 시험지용 ➡️ 일반용 로직 및 Streamlit UI 구현부는동일하며, 함수 호출만 convert_general_to_exam_v3로 매핑)
