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
# 1. 일반용 ➡️ 시험지용 변환 (디자인 템플릿 + 표/지문 무결성 이식)
# ==========================================
def convert_general_to_exam_perfect(source_doc):
    """
    내장된 템플릿(template_exam.docx)의 2단 레이아웃과 상단 분홍색 디자인 틀을 가져옵니다.
    그 후 원본 일반용 문서의 발문, 선지, 이미지, '표 객체 자체'를 무결하게 복사하여 주입합니다.
    """
    template_path = "template_exam.docx"
    
    # 1. 상단 분홍색 그림 및 2단 다단 서식이 잡혀있는 템플릿 로드
    if os.path.exists(template_path):
        target_doc = Document(template_path)
    else:
        st.error("🚨 서버에 'template_exam.docx' 파일이 없습니다! GitHub 레포지토리에 템플릿 파일을 업로드해주세요.")
        target_doc = Document()
        target_doc.add_paragraph("2026년 봄학기 THE OPEN 정기평가 (Level P)")
        target_doc.add_paragraph("[ PartⅡ 문법 | 20문항 20분 ]\n")

    q_counter = 1
    
    # 2. 원본(일반용) 문서의 body 엘리먼트를 순차적으로 순회하며 이식
    for element in source_doc.element.body:
        
        # [단락 처리: 발문, 선지, 보기 텍스트 및 이미지]
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, source_doc)
            p_text = p.text.strip()
            
            # 공백 단락이나 [Chapter ...] 메타데이터는 복사하지 않고 패스
            if not p_text or p_text.startswith("[Chapter"):
                continue
            
            new_p = target_doc.add_paragraph()
            # 원본 단락의 간격 스타일 유지
            new_p.paragraph_format.space_before = p.paragraph_format.space_before
            new_p.paragraph_format.space_after = p.paragraph_format.space_after
            new_p.paragraph_format.line_spacing = p.paragraph_format.line_spacing
            
            # 문항 번호 시작점 감지 (예: 1. 다음 중...)
            if re.match(r'^\d+\.', p_text):
                # 기존 텍스트에서 '1.' 또는 '1. ' 형태의 앞부분 번호를 완전히 도려내어 중복 방지
                clean_p_text = re.sub(r'^\d+\.\s*', '', p_text)
                
                # 정렬된 새 번호만 볼드로 깔끔하게 추가
                run_num = new_p.add_run(f"{q_counter}.  ")
                run_num.bold = True
                q_counter += 1
                
                # 번호 뒷부분의 문항 본문 내용(Runs)을 서식/이미지 깨짐 없이 이전
                for run in p.runs:
                    # 각 Run 텍스트 내에 중복 번호 문자가 섞여 들어가는 현상 정밀 필터링
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
                # 일반 보기 선지 및 일반 지문 단락은 서식 통째로 이전
                for run in p.runs:
                    new_run = new_p.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.font.size = run.font.size
                    new_run.font.name = run.font.name
                    if run._r.xpath('.//w:drawing'):
                        new_run._r.append(run._r.xpath('.//w:drawing')[0])
                        
        # [표 처리: 지문이 들어있는 Box 표 구조물]
        elif element.tag.endswith('tbl'):
            # 일반용 문서에 있던 원본 표 구조 가져오기
            src_tbl = docx.table.Table(element, source_doc)
            
            # 템플릿 문서 내부에 원본과 완벽히 일치하는 행/열 크기의 새 표 생성
            dst_tbl = target_doc.add_table(rows=len(src_tbl.rows), cols=len(src_tbl.columns))
            dst_tbl.style = src_tbl.style
            
            # 표 내부의 '모든 단락(Paragraph) 구조'와 '세부 서식(Runs)'을 1:1로 정밀 복사
            # 이 방식을 사용하여 표 안의 지문 내용, 쉼표, 특수기호, 줄바꿈이 일반용과 100% 동일하게 출력됩니다.
            for r_idx, row in enumerate(src_tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    dst_cell = dst_tbl.cell(r_idx, c_idx)
                    dst_cell.text = "" # 기본 생성 텍스트 초기화
                    
                    for p_cell in cell.paragraphs:
                        dst_p_cell = dst_cell.add_paragraph()
                        dst_p_cell.paragraph_format.space_before = p_cell.paragraph_format.space_before
                        dst_p_cell.paragraph_format.space_after = p_cell.paragraph_format.space_after
                        
                        for run in p_cell.runs:
                            dst_run = dst_p_cell.add_run(run.text)
                            dst_run.bold = run.bold
                            dst_run.italic = run.italic
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
def convert_exam_to_general_perfect(source_doc):
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
st.markdown("디자인 서식(상단 분홍색 배너, 2단 레이아웃)과 **표 내부 지문 데이터**를 완벽하게 결합하여 변환합니다.")

# 서버 내 템플릿 파일 유무 실시간 검증 및 에러 안내
if not os.path.exists("template_exam.docx"):
    st.error("🚨 [설정 에러] 현재 폴더에 `template_exam.docx` 디자인 템플릿 파일이 보이지 않습니다. 일반용 문서에 시험지 레이아웃을 입히려면 반드시 배경 틀이 마련된 템플릿 파일을 GitHub 같은 폴더에 업로드해 주셔야 합니다.")

uploaded_file = st.file_uploader("변환할 정기평가 Word 파일(.docx)을 업로드하세요.", type=["docx"])

if uploaded_file is not None:
    doc = Document(uploaded_file)
    sample_text = "\n".join([p.text for p in doc.paragraphs[:15] if p.text.strip()])
    
    if "[Chapter" in sample_text:
        default_index = 0
        detected_text = "🔍 **문서 양식 감지 결과:** [일반용 양식]이 확인되었습니다. 시험지용 디자인 템플릿 프레임에 맞춰 표 내부 지문까지 정밀 변환합니다."
    else:
        default_index = 1
        detected_text = "🔍 **문서 양식 감지 결과:** [시험지용 양식]이 확인되었습니다. [일반용] 양식으로 변환합니다."
        
    st.info(detected_text)
    
    mode = st.radio(
        "변환 방향을 확인하거나 변경하세요:",
        ("일반용 ➡️ 시험지용", "시험지용 ➡️ 일반용"),
        index=default_index
    )
    
    if st.button("🚀 디자인 및 표 보존 변환 시작", use_container_width=True):
        with st.spinner("표 내부 정밀 데이터 구조와 문항 서식을 병합하고 있습니다..."):
            try:
                if "일반용 ➡️ 시험지용" in mode:
                    out_bytes = convert_general_to_exam_perfect(doc)
                    file_name = "변환_시험지용_정기평가.docx"
                else:
                    out_bytes = convert_exam_to_general_perfect(doc)
                    file_name = "변환_일반용_정기평가.docx"
                
                st.success("🎉 변환 및 레이아웃 병합 처리가 완료되었습니다!")
                st.download_button(
                    label="💾 변환된 Word 파일 다운로드",
                    data=out_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ 변환 처리 중 오류가 발생했습니다: {str(e)}")
