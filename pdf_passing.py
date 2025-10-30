import pdfplumber
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import tempfile

def extract_pdf_to_text(pdf_path):
    """
    PDF 파일을 읽어서 모든 페이지의 텍스트를 추출하여 텍스트 파일로 저장합니다.
    
    Args:
        pdf_path (str): PDF 파일 경로
        
    Returns:
        str: 생성된 텍스트 파일의 경로
    """
    
    if not os.path.exists(pdf_path):
        print(f"오류: '{pdf_path}' 파일을 찾을 수 없습니다.")
        return None
    
    try:
        # PDF 파일명에서 확장자를 제거하고 텍스트 파일명 생성
        pdf_filename = os.path.basename(pdf_path)
        text_filename = os.path.splitext(pdf_filename)[0] + "_extracted.txt"
        text_path = os.path.join(os.path.dirname(pdf_path), text_filename)
        
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            
            print(f"PDF 파일: {pdf_path}")
            print(f"전체 페이지 수: {total_pages}")
            print(f"텍스트 추출 중...")
            print("=" * 50)
            
            # 모든 텍스트를 저장할 리스트
            all_text = []
            all_text.append(f"PDF 파일: {pdf_filename}\n")
            all_text.append(f"전체 페이지 수: {total_pages}\n")
            all_text.append(f"추출 일시: {os.path.basename(__file__)}\n")
            all_text.append("=" * 50 + "\n\n")
            
            for page_num in range(total_pages):
                page = pdf.pages[page_num]
                text = page.extract_text()
                
                print(f"페이지 {page_num + 1}/{total_pages} 처리 중...")
                
                all_text.append(f"[ 페이지 {page_num + 1} ]\n")
                all_text.append("-" * 30 + "\n")
                
                if text:
                    # 텍스트를 줄별로 나누어서 저장
                    lines = text.split('\n')
                    for line_num, line in enumerate(lines, 1):
                        if line.strip():  # 빈 줄이 아닌 경우만 저장
                            all_text.append(f"줄 {line_num:3d}: {line}\n")
                else:
                    all_text.append("이 페이지에서 텍스트를 추출할 수 없습니다.\n")
                
                all_text.append("\n")  # 페이지 구분을 위한 빈 줄
            
            # 텍스트 파일로 저장
            with open(text_path, 'w', encoding='utf-8') as f:
                f.writelines(all_text)
            
            print(f"\n텍스트 추출 완료!")
            print(f"저장된 파일: {text_path}")
            
            return text_path
                    
    except Exception as e:
        print(f"PDF 읽기 중 오류가 발생했습니다: {str(e)}")
        return None

def open_with_notepad(file_path):
    """
    텍스트 파일을 메모장으로 엽니다.
    
    Args:
        file_path (str): 열 텍스트 파일의 경로
    """
    try:
        # Windows 메모장으로 파일 열기
        subprocess.run(['notepad', file_path], check=True)
    except subprocess.CalledProcessError:
        print(f"메모장으로 파일을 열 수 없습니다: {file_path}")
    except FileNotFoundError:
        print("메모장을 찾을 수 없습니다. 기본 텍스트 에디터로 열어보세요.")
        # 기본 프로그램으로 열기 시도
        try:
            os.startfile(file_path)
        except:
            print(f"파일을 열 수 없습니다: {file_path}")

def select_pdf_file():
    """
    GUI로 PDF 파일을 선택하는 함수
    
    Returns:
        str: 선택된 PDF 파일의 경로, 취소시 None
    """
    # tkinter 윈도우 생성 (숨김)
    root = tk.Tk()
    root.withdraw()  # 메인 윈도우 숨기기
    
    # 파일 선택 대화상자 열기
    pdf_path = filedialog.askopenfilename(
        title="PDF 파일을 선택하세요",
        filetypes=[
            ("PDF 파일", "*.pdf"),
            ("모든 파일", "*.*")
        ],
        initialdir=os.getcwd()  # 현재 디렉토리에서 시작
    )
    
    root.destroy()  # tkinter 윈도우 제거
    
    return pdf_path if pdf_path else None

def main():
    """
    메인 함수: GUI로 PDF 파일을 선택받아 텍스트로 추출하고 메모장으로 엽니다.
    """
    
    if len(sys.argv) > 1:
        # 명령행 인수로 파일 경로가 제공된 경우
        pdf_path = sys.argv[1]
        print(f"명령행에서 제공된 파일: {pdf_path}")
    else:
        # GUI로 파일 선택
        print("PDF 파일 선택 창을 열고 있습니다...")
        pdf_path = select_pdf_file()
        
        if not pdf_path:
            print("파일 선택이 취소되었습니다.")
            return
            
        print(f"선택된 파일: {pdf_path}")
    
    # PDF에서 텍스트 추출
    text_file_path = extract_pdf_to_text(pdf_path)
    
    if text_file_path:
        print("\n메모장으로 파일을 열고 있습니다...")
        open_with_notepad(text_file_path)
    else:
        print("텍스트 추출에 실패했습니다.")

if __name__ == "__main__":
    main()
