import tkinter as tk
from tkinter import filedialog, messagebox
import PyPDF2
import pandas as pd
import re
import os


def select_pdf_file():
    """파일 탐색기를 열어 PDF 파일 선택"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    file_path = filedialog.askopenfilename(
        title='PDF 파일을 선택하세요',
        filetypes=[('PDF files', '*.pdf'), ('All files', '*.*')],
        initialdir=os.getcwd()
    )
    
    root.destroy()
    return file_path


def extract_text_from_pdf(pdf_path):
    """PDF 파일에서 텍스트 추출"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            num_pages = len(pdf_reader.pages)
            
            print(f"   - 파일명: {os.path.basename(pdf_path)}")
            print(f"   - 총 페이지: {num_pages}페이지")
            
            # 모든 페이지의 텍스트 추출
            all_text = []
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                all_text.append(text)
            
            full_text = '\n'.join(all_text)
            print(f"   - 추출된 텍스트 길이: {len(full_text)} 글자")
            
            return full_text
            
    except Exception as e:
        print(f" PDF 읽기 오류: {e}")
        return None


def extract_ref_des_with_location(text):
    """
    텍스트에서 Ref Des 추출 및 Top/BTM 구분
    
    필터링 조건:
    1. 탭 또는 공백으로 구분된 데이터 행
    2. 첫 번째 단어가 아래 3가지 패턴 중 하나:
       - D + 숫자 (D04, D1, D2, D10 등)
       - U + 숫자 (U04, U1, U2, U10 등)
       - 대문자로 시작하는 영문/숫자/언더스코어 조합 (FL13_14_85, FL28_20R 등)
    3. 두 번째, 세 번째 단어가 숫자 (X, Y 좌표값)
    4. "Bottom Assembly" 이전은 Top, 이후는 BTM으로 구분
    """
    
    print("\n Ref Des 추출 중...")
    print("\n 추출 패턴:")
    print("   1. D + 숫자 (예: D04, D1, D2)")
    print("   2. U + 숫자 (예: U04, U1, U2)")
    print("   3. 영문/숫자/언더스코어 조합 (예: FL13_14_85, FL28_20R)")
    print()
    
    lines = text.split('\n')
    ref_des_list = []
    is_bottom = False  # Bottom Assembly 이후인지 체크
    
    for line_num, line in enumerate(lines, 1):
        # "Bottom Assembly" 키워드 체크
        if 'Bottom' in line and 'Assembly' in line:
            is_bottom = True
            print(f"\n   'Bottom Assembly' 발견 (Line {line_num})")
            print(f"   이후 Ref Des는 BTM으로 분류됩니다.\n")
            continue
        
        # 빈 줄 건너뛰기
        if not line.strip():
            continue
        
        ref_des = None
        pattern_type = None
        extraction_method = None
        
        # ===== 방법 1: 구분자 기반 split 시도 =====
        # 1개 이상의 공백 또는 탭을 구분자로 사용
        parts = re.split(r'\s+|\t+', line.strip())
        parts = [p.strip() for p in parts if p.strip()]
        
        # 최소 3개 이상의 필드가 있어야 함 (Ref Des + X좌표 + Y좌표)
        if len(parts) >= 3:
            first_word = parts[0]
            
            # 조건 1: 첫 단어가 Ref Des 패턴인지 확인 (3가지 패턴)
            is_d_pattern = bool(re.match(r'^D\d+$', first_word))
            is_u_pattern = bool(re.match(r'^U\d+$', first_word))
            is_complex_pattern = bool(re.match(r'^[A-Z][A-Z0-9_]+$', first_word))
            
            if is_d_pattern or is_u_pattern or is_complex_pattern:
                # 조건 2: 두 번째 단어가 숫자로 시작 (좌표값)
                if re.match(r'^\d+\.?\d*', parts[1]):
                    # 조건 3: 세 번째 필드도 숫자인지 확인 (Y 좌표)
                    if len(parts) >= 3 and re.match(r'^\d+\.?\d*', parts[2]):
                        ref_des = first_word
                        extraction_method = "구분자split"
                        
                        if is_d_pattern:
                            pattern_type = "D+숫자"
                        elif is_u_pattern:
                            pattern_type = "U+숫자"
                        else:
                            pattern_type = "복합패턴"
        
        # ===== 방법 2: 정규표현식으로 직접 추출 (방법 1 실패 시) =====
        if ref_des is None:
            # 패턴: (Ref Des) + (공백) + (숫자.숫자) + (공백) + (숫자.숫자)
            # Cu+Sn/SnAg는 선택사항 (있어도 되고 없어도 됨)
            # 3가지 Ref Des 패턴을 모두 포함
            regex_pattern = r'^\s*(D\d+|U\d+|[A-Z][A-Z0-9_]+)\s+(\d+\.?\d*)\s+(\d+\.?\d*)'
            match = re.search(regex_pattern, line)
            
            if match:
                ref_des = match.group(1)
                extraction_method = "정규표현식"
                
                # 패턴 타입 결정
                if re.match(r'^D\d+$', ref_des):
                    pattern_type = "D+숫자"
                elif re.match(r'^U\d+$', ref_des):
                    pattern_type = "U+숫자"
                else:
                    pattern_type = "복합패턴"
        
        # ===== 추출 성공 시 리스트에 추가 =====
        if ref_des:
            # Top 또는 BTM 구분
            location = 'BTM' if is_bottom else 'Top'
            
            ref_des_list.append({
                'Ref Des': ref_des,
                'Top/BTM': location,
                'Line': line_num,
                'Pattern': pattern_type
            })
            
            print(f"   Line {line_num}: {ref_des} → {location} [{pattern_type}] ({extraction_method})")
    
    # 추출 결과 요약
    print(f"\n 추출 완료: {len(ref_des_list)}개의 Ref Des 발견")
    
    top_count = sum(1 for item in ref_des_list if item['Top/BTM'] == 'Top')
    btm_count = sum(1 for item in ref_des_list if item['Top/BTM'] == 'BTM')
    
    print(f"   - Top: {top_count}개")
    print(f"   - BTM: {btm_count}개")
    
    # 패턴별 통계
    d_count = sum(1 for item in ref_des_list if item['Pattern'] == 'D+숫자')
    u_count = sum(1 for item in ref_des_list if item['Pattern'] == 'U+숫자')
    complex_count = sum(1 for item in ref_des_list if item['Pattern'] == '복합패턴')
    
    print(f"\n 패턴별 통계:")
    print(f"   - D+숫자: {d_count}개")
    print(f"   - U+숫자: {u_count}개")
    print(f"   - 복합패턴 (FL13_14_85 등): {complex_count}개")
    
    return ref_des_list


def save_to_excel(ref_des_list, pdf_path):
    """Ref Des를 Excel 파일로 저장"""
    
    if not ref_des_list:
        print("\n 추출된 Ref Des가 없습니다.")
        return None
    
    try:
        # DataFrame 생성 (Line과 Pattern 컬럼 제외)
        df = pd.DataFrame([
            {'Ref Des': item['Ref Des'], 'Top/BTM': item['Top/BTM']}
            for item in ref_des_list
        ])
        
        # 출력 파일 경로 생성 (PDF와 같은 폴더에 저장)
        pdf_dir = os.path.dirname(pdf_path)
        pdf_basename = os.path.splitext(os.path.basename(pdf_path))[0]
        output_file = os.path.join(pdf_dir, f"{pdf_basename}_RefDes.xlsx")
        
        # Excel로 저장
        df.to_excel(output_file, index=False, engine='openpyxl')
        
        print(f"\n Excel 파일 저장 완료!")
        print(f"   - 저장 위치: {output_file}")
        print(f"   - 저장된 항목: {len(ref_des_list)}개")
        
        # 미리보기 출력 (최대 10개)
        print(f"\n 데이터 미리보기 (최대 10개):")
        print(df.head(10).to_string(index=False))
        if len(df) > 10:
            print(f"   ... 외 {len(df) - 10}개")
        
        return output_file
        
    except Exception as e:
        print(f"\n Excel 저장 오류: {e}")
        return None


def show_message_box(title, message, msg_type="info"):
    """메시지 박스 표시"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    if msg_type == "info":
        messagebox.showinfo(title, message)
    elif msg_type == "warning":
        messagebox.showwarning(title, message)
    elif msg_type == "error":
        messagebox.showerror(title, message)
    
    root.destroy()


def main():
    """메인 실행 함수"""
    
    print("=" * 60)
    print("  PDF Ref Des 추출 프로그램")
    print("=" * 60)
    
    # 1단계: PDF 파일 선택
    print("\n PDF 파일을 선택해주세요...")
    pdf_path = select_pdf_file()
    
    if not pdf_path:
        message = "파일이 선택되지 않았습니다."
        print("\n" + message)
        show_message_box("알림", message, "warning")
        return
    
    # 2단계: PDF 텍스트 추출
    print("\n" + "=" * 60)
    text = extract_text_from_pdf(pdf_path)
    
    if not text:
        message = "PDF에서 텍스트를 추출할 수 없습니다."
        print("\n" + message)
        show_message_box("오류", message, "error")
        return
    
    # 3단계: Ref Des 추출
    print("\n" + "=" * 60)
    ref_des_list = extract_ref_des_with_location(text)
    
    if not ref_des_list:
        message = "조건에 맞는 Ref Des를 찾을 수 없습니다."
        print("\n" + message)
        show_message_box("알림", message, "warning")
        return
    
    # 4단계: Excel로 저장
    print("\n" + "=" * 60)
    output_file = save_to_excel(ref_des_list, pdf_path)
    
    if output_file:
        print("\n" + "=" * 60)
        print("모든 작업이 완료되었습니다!")
        print("=" * 60)
        
        # 최종 결과 메시지 박스
        top_count = sum(1 for item in ref_des_list if item['Top/BTM'] == 'Top')
        btm_count = sum(1 for item in ref_des_list if item['Top/BTM'] == 'BTM')
        
        message = f"모든 작업이 완료되었습니다!\n\n"
        message += f"추출된 Ref Des: {len(ref_des_list)}개\n"
        message += f"  - Top: {top_count}개\n"
        message += f"  - BTM: {btm_count}개\n\n"
        message += f"저장 위치:\n{output_file}"
        
        show_message_box("완료", message, "info")
        
        return output_file


# 실행
if __name__ == "__main__":
    result = main()
else:
    # Jupyter 노트북에서 실행할 경우
    result = main()