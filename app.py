# ... (앞부분 생략)

        # [2] 지문 상자 처리 (Table 구조 발견 시 테두리 보존 및 너비 최적화)
        elif element.tag.endswith('tbl'):
            src_tbl = docx.table.Table(element, source_doc)
            
            # [수정된 부분] 2단 레이아웃 내에서 잘리지 않도록 autofit=True 적용
            box_table = target_doc.add_table(rows=len(src_tbl.rows), cols=len(src_tbl.columns))
            box_table.autofit = True 
            box_table.style = 'Table Grid' # 테두리가 확실하게 보이도록 스타일 지정
            
            # 테두리 설정 (검은색 단선)
            for r_idx, row in enumerate(src_tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    dst_cell = box_table.cell(r_idx, c_idx)
                    set_cell_borders(dst_cell)
                    
                    dst_cell.text = "" 
                    
                    # 셀 내부 텍스트 및 서식 정밀 복사
                    for p_cell in cell.paragraphs:
                        if p_cell.text.strip():
                            dst_p_cell = dst_cell.add_paragraph()
                            for run in p_cell.runs:
                                new_run = dst_p_cell.add_run(run.text)
                                new_run.bold = run.bold
                                new_run.italic = run.italic
                                new_run.font.size = run.font.size
                                new_run.font.name = run.font.name

# ... (나머지 코드 동일)
