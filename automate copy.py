import csv
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime  # 날짜와 시간을 위해 추가
import win32com.client  # Excel COM 객체를 위해 추가
import threading  # 백그라운드 작업을 위해 추가

# --- 1. 기본 설정 ---

def center_window(window):
    """창을 화면 중앙에 위치시키는 함수"""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

def select_file_dialog(title, filetypes, initial_dir=None):
    """파일 선택 다이얼로그를 화면 중앙에 표시하는 함수"""
    root = tk.Tk()
    root.withdraw()  # 메인 창 숨기기
    
    # 다이얼로그를 중앙에 위치시키기 위한 설정
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - 200
    y = (root.winfo_screenheight() // 2) - 150
    root.geometry(f'+{x}+{y}')
    
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=filetypes,
        initialdir=initial_dir
    )
    
    root.destroy()
    return file_path

def select_directory_dialog(title, initial_dir=None):
    """폴더 선택 다이얼로그를 화면 중앙에 표시하는 함수"""
    root = tk.Tk()
    root.withdraw()  # 메인 창 숨기기
    
    # 다이얼로그를 중앙에 위치시키기 위한 설정
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - 200
    y = (root.winfo_screenheight() // 2) - 150
    root.geometry(f'+{x}+{y}')
    
    directory = filedialog.askdirectory(
        title=title,
        initialdir=initial_dir
    )
    
    root.destroy()
    return directory

# 전역 Excel 애플리케이션 인스턴스 (재사용을 위해)
_global_excel_app = None

def get_excel_app():
    """Excel 애플리케이션 인스턴스를 가져오거나 생성하는 함수 (재사용)"""
    global _global_excel_app
    
    if _global_excel_app is None:
        try:
            _global_excel_app = win32com.client.DispatchEx("Excel.Application")
            _global_excel_app.Visible = False
            _global_excel_app.DisplayAlerts = False
            _global_excel_app.ScreenUpdating = False
            _global_excel_app.EnableEvents = False
            # Calculation 속성 설정 제거 (오류 방지)
            print("  - 새로운 Excel 애플리케이션 인스턴스 생성")
        except Exception as e:
            print(f"  ! Excel 애플리케이션 생성 오류: {e}")
            return None
    
    return _global_excel_app

def cleanup_excel_app():
    """전역 Excel 애플리케이션 정리"""
    global _global_excel_app
    
    if _global_excel_app:
        try:
            # Calculation 속성 복원 제거 (오류 방지)
            _global_excel_app.ScreenUpdating = True
            _global_excel_app.EnableEvents = True
            _global_excel_app.Quit()
            del _global_excel_app
            _global_excel_app = None
            print("  - Excel 애플리케이션 정리 완료")
        except Exception as e:
            print(f"  - Excel 정리 오류: {e}")

def export_first_three_sheets_to_pdf(excel_file_path, pdf_file_path):
    """
    엑셀 파일의 첫 번째, 두 번째, 세 번째 시트를 하나의 PDF로 내보내는 함수 (최적화됨)
    """
    import time
    
    workbook = None
    
    try:
        # 재사용 가능한 Excel 애플리케이션 가져오기
        excel_app = get_excel_app()
        if not excel_app:
            return False
        
        # 엑셀 파일 열기 (절대 경로로 변환)
        abs_excel_path = os.path.abspath(excel_file_path)
        workbook = excel_app.Workbooks.Open(
            abs_excel_path,
            UpdateLinks=0,  # 링크 업데이트 안 함
            ReadOnly=True,  # 읽기 전용으로 열기
            IgnoreReadOnlyRecommended=True
        )
        
        # 시트 개수 확인 (빠른 체크)
        if workbook.Worksheets.Count < 3:
            print(f"  경고: 시트가 {workbook.Worksheets.Count}개만 있습니다.")
            return False
        
        # PDF 경로 준비
        abs_pdf_path = os.path.abspath(pdf_file_path)
        
        # 첫 3개 시트를 배열로 선택 (더 빠른 방법)
        try:
            sheet_names = [workbook.Worksheets(i).Name for i in range(1, 4)]  # 1,2,3번째 시트
            excel_app.Sheets(sheet_names).Select()
        except Exception as e:
            # 배열 선택 실패 시 개별 선택
            workbook.Worksheets(1).Select()
            for i in range(2, 4):
                try:
                    workbook.Worksheets(i).Select(False)
                except:
                    pass
        
        # PDF로 내보내기 (호환성을 위해 간단한 설정)
        excel_app.ActiveSheet.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=abs_pdf_path
        )
        
        return True
        
    except Exception as e:
        print(f"  ! PDF 내보내기 오류: {e}")
        return False
        
    finally:
        # 워크북만 닫기 (Excel 앱은 재사용을 위해 유지)
        try:
            if workbook:
                workbook.Close(SaveChanges=False)
        except Exception as e:
            print(f"  - 워크북 닫기 오류: {e}")

def validate_excel_data_with_com(excel_file_path):
    """
    저장된 엑셀 파일을 COM으로 열어서 텍스트 내용으로 검증하는 함수 (최적화됨)
    Linearity 시트(3번째 시트)의 D29~H29 셀 텍스트 내용을 확인
    "Pass" = Pass, "Fail" = Fail
    
    Returns:
        tuple: (is_valid, status_suffix)
        - is_valid: True if all Pass, False if any Fail
        - status_suffix: "_P" for Pass, "_F" for Fail
    """
    workbook = None
    
    try:
        # 재사용 가능한 Excel 애플리케이션 가져오기
        excel_app = get_excel_app()
        if not excel_app:
            return False, "_F"
        
        # 엑셀 파일 열기 (최적화된 설정)
        abs_excel_path = os.path.abspath(excel_file_path)
        workbook = excel_app.Workbooks.Open(
            abs_excel_path,
            UpdateLinks=0,  # 링크 업데이트 안 함
            ReadOnly=True,  # 읽기 전용으로 열기
            IgnoreReadOnlyRecommended=True
        )
        
        # 계산 강제 실행 (수식 계산 완료)
        excel_app.Calculate()
        
        # Linearity 시트 가져오기 (인덱스로 빠르게 접근)
        try:
            ws_linearity = workbook.Worksheets(3)  # 3번째 시트 (인덱스로 빠른 접근)
        except:
            ws_linearity = workbook.Worksheets("Linearity")  # 이름으로 폴백
        
        # 개별 셀 확인 (안정성을 위해)
        check_cells_29 = ['D29', 'E29', 'F29', 'G29', 'H29']
        fail_found = False
        pass_count = 0
        
        print("  - Linearity 시트 29행 텍스트 검사:")
        for cell_29 in check_cells_29:
            try:
                cell_value = ws_linearity.Range(cell_29).Value
                cell_str = str(cell_value).strip() if cell_value is not None else ""
                
                print(f"    Linearity {cell_29}: '{cell_str}'")
                
                if "Pass" in cell_str:
                    pass_count += 1
                    print(f"    → 'Pass' 확인")
                elif "Fail" in cell_str:
                    fail_found = True
                    print(f"    → 'Fail' 확인")
                    break  # Fail 발견 시 즉시 중단
                else:
                    print(f"    → Pass/Fail 텍스트 없음: '{cell_str}'")
                    
            except Exception as e:
                print(f"    Linearity {cell_29}: 읽기 오류 ({e})")
                continue
        
        # 결과 판정
        if fail_found:
            print("  ✗ 검증 결과: 'Fail' 텍스트 발견 - PDF 생성하지 않음")
            return False, "_F"
        elif pass_count > 0:
            print(f"  ✓ 검증 결과: {pass_count}개 'Pass' 텍스트 확인 - PDF 생성 진행")
            return True, "_P"
        else:
            print("  ✗ 검증 결과: 'Pass' 텍스트가 없음 - PDF 생성하지 않음")
            return False, "_F"
            
    except Exception as e:
        print(f"  ! 데이터 검증 중 오류 발생: {e}")
        return False, "_F"
        
    finally:
        # 워크북만 닫기 (Excel 앱은 재사용을 위해 유지)
        try:
            if workbook:
                workbook.Close(SaveChanges=False)
        except Exception as e:
            print(f"  - 워크북 닫기 오류: {e}")

# CSV 데이터를 읽어 엑셀 셀에 매핑하는 헬퍼 함수
def get_csv_data(all_data, row_num, col_idx, as_number=False):
    """
    all_data (list of lists)에서 1기반 행/열 인덱스로 데이터를 안전하게 가져옵니다.
    as_number=True일 경우, 숫자로 변환을 시도하고 실패 시 0을 반환합니다.
    """
    try:
        # CSV 행 인덱스는 0부터 시작하므로 (row_num - 1)
        value = all_data[row_num - 1][col_idx]
        
        if as_number:
            try:
                # 쉼표 등 문자열 제거 후 float 변환
                return float(value.replace(',', ''))
            except (ValueError, TypeError):
                return 0.0  # 변환 실패 시 0
        
        return value # 문자열로 반환
        
    except IndexError:
        # 데이터 범위를 벗어난 경우
        print(f"  경고: CSV {row_num}행, {col_idx}열에 접근할 수 없습니다.")
        return 0.0 if as_number else ""

def count_valid_columns(all_csv_data):
    """CSV에서 C열부터 데이터가 있는 열의 개수를 세는 함수"""
    if not all_csv_data or len(all_csv_data) < 2:
        return 0
    
    num_cols = len(all_csv_data[0])
    valid_count = 0
    
    for col_idx in range(2, num_cols):  # C열(인덱스 2)부터 시작
        check_value = get_csv_data(all_csv_data, 2, col_idx).strip()
        if check_value:
            valid_count += 1
        else:
            break  # 빈 열을 만나면 중단
    
    return valid_count

class ProgressWindow:
    """진행 상황을 표시하는 로딩바 창"""
    
    def __init__(self, total_files):
        self.total_files = total_files
        self.current_file = 0
        
        # 메인 창 생성
        self.root = tk.Tk()
        self.root.title("엑셀 자동화 진행 상황")
        self.root.geometry("550x280")  # 크기 증가 (가로 550, 세로 280)
        self.root.resizable(False, False)
        
        # 창을 최상위에 표시
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.root.focus_force()
        
        # 창을 화면 중앙에 배치
        self.center_window()
        
        # UI 구성
        self.setup_ui()
        
        # 창 닫기 방지
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def center_window(self):
        """창을 화면 중앙에 위치시키는 함수"""
        self.root.update_idletasks()
        
        # 창 크기 설정
        width = 550
        height = 280
        
        # 화면 중앙 계산
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        
        # 창 위치 및 크기 설정
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def setup_ui(self):
        """UI 구성 요소 설정"""
        # 상단 여백
        top_spacer = tk.Frame(self.root, height=25)
        top_spacer.pack()
        
        # 제목 라벨
        title_label = tk.Label(
            self.root, 
            text="엑셀 파일 자동 생성 중...", 
            font=("맑은 고딕", 16, "bold")
        )
        title_label.pack(pady=15)
        
        # 진행 상황 텍스트
        self.status_label = tk.Label(
            self.root, 
            text=f"0 / {self.total_files} 파일 처리 완료", 
            font=("맑은 고딕", 12)
        )
        self.status_label.pack(pady=12)
        
        # 진행률 바
        self.progress_bar = ttk.Progressbar(
            self.root, 
            length=450,  # 길이 증가
            mode='determinate',
            style='TProgressbar'
        )
        self.progress_bar.pack(pady=20)
        
        # 현재 작업 라벨
        self.current_task_label = tk.Label(
            self.root, 
            text="작업 준비 중...", 
            font=("맑은 고딕", 10),
            fg="blue",
            wraplength=500  # 텍스트 줄바꿈 설정
        )
        self.current_task_label.pack(pady=15)
        
        # 하단 여백
        bottom_spacer = tk.Frame(self.root, height=20)
        bottom_spacer.pack()
        
        # 진행률 설정
        self.progress_bar['maximum'] = self.total_files
        self.progress_bar['value'] = 0
        
    def update_progress(self, current_file, current_task=""):
        """진행 상황 업데이트"""
        self.current_file = current_file
        
        # 진행률 바 업데이트
        self.progress_bar['value'] = current_file
        
        # 상태 텍스트 업데이트
        self.status_label.config(text=f"{current_file} / {self.total_files} 파일 처리 완료")
        
        # 현재 작업 텍스트 업데이트
        if current_task:
            self.current_task_label.config(text=current_task)
        
        # 진행률 퍼센트 계산
        if self.total_files > 0:
            percent = (current_file / self.total_files) * 100
            self.root.title(f"엑셀 자동화 진행 상황 - {percent:.1f}%")
        
        # UI 업데이트
        self.root.update()
        
    def close(self):
        """창 닫기"""
        self.root.destroy()
        
    def on_closing(self):
        """창 닫기 시도 시 호출 (무시)"""
        pass  # 사용자가 임의로 창을 닫지 못하게 함

# --- 2. 메인 스크립트 실행 ---
if __name__ == "__main__":
    print("--- 엑셀 자동화 스크립트 시작 ---")

    try:
        # --- 2-1. 엑셀 템플릿 파일 선택 ---
        print("Linearity_ED2 파일을 선택하세요...")
        excel_template_path = select_file_dialog(
            title="Linearity_ED2 파일을 선택해주세요",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        
        if not excel_template_path:
            print("! 작업 취소: 엑셀 템플릿 파일이 선택되지 않았습니다.")
            exit()
            
        print(f"선택된 엑셀 템플릿: {excel_template_path}")

        # --- 2-2. CSV 데이터 파일 선택 ---
        print("CSV 데이터 파일을 선택하세요...")
        csv_data_path = select_file_dialog(
            title="CSV 데이터 파일을 선택하세요",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not csv_data_path:
            print("! 작업 취소: CSV 데이터 파일이 선택되지 않았습니다.")
            exit()
            
        print(f"선택된 CSV 파일: {csv_data_path}")

        # --- 2-3. 저장 폴더 선택 ---
        print("결과 파일을 저장할 폴더를 선택하세요...")
        save_directory = select_directory_dialog(title="결과 파일을 저장할 폴더를 선택하세요")

        if not save_directory:
            print("! 작업 취소: 저장 폴더가 선택되지 않았습니다.")
            exit()
            
        print(f"저장 위치: {save_directory}")

        # --- 2-4. CSV 파일 전체 읽기 ---
        all_csv_data = []
        try:
            # utf-8-sig는 엑셀 CSV의 BOM(Byte Order Mark)을 처리합니다.
            with open(csv_data_path, mode='r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                all_csv_data = list(reader)
        except FileNotFoundError:
            print(f"! 오류: CSV 파일 '{csv_data_path}'을(를) 찾을 수 없습니다.")
            exit()
        except Exception as e:
            print(f"! 오류: CSV 파일을 읽는 중 오류가 발생했습니다: {e}")
            exit()
            
        if not all_csv_data:
            print("! 오류: CSV 파일이 비어있습니다.")
            exit()

        # --- 2-5. 처리할 파일 개수 계산 및 진행바 초기화 ---
        total_files = count_valid_columns(all_csv_data)
        
        if total_files == 0:
            print("! 오류: 처리할 데이터가 없습니다. CSV C열의 2행에 데이터가 있는지 확인하세요.")
            exit()
        
        print(f"총 {total_files}개의 파일을 생성합니다.")
        
        # 진행바 창 생성
        progress_window = ProgressWindow(total_files)
        progress_window.update_progress(0, "작업 시작...")
        
        # CSV의 총 열 개수 (첫 번째 행 기준)
        num_cols = len(all_csv_data[0])
        file_counter = 1

        # --- 2-6. C열(인덱스 2)부터 반복 시작 ---
        for col_idx in range(2, num_cols):
            
            # 기준 행(예: 2행 'Analyte')에 값이 있는지 확인하여 반복 중단 결정
            check_value = get_csv_data(all_csv_data, 2, col_idx).strip()
            
            if not check_value:
                print(f"\n- {col_idx}열(CSV의 {chr(65+col_idx)}열)에 데이터가 없어 작업을 중지합니다.")
                break

            print(f"\n▶ {file_counter}번째 파일 생성 중... (CSV {chr(65+col_idx)}열 데이터 기준)")
            
            # 진행바 업데이트
            progress_window.update_progress(
                file_counter - 1, 
                f"{file_counter}번째 파일 생성 중... (CSV {chr(65+col_idx)}열)"
            )

            # --- 2-7. 엑셀 템플릿 불러오기 ---
            try:
                wb = openpyxl.load_workbook(excel_template_path, keep_vba=True)
            except FileNotFoundError:
                print(f"! 오류: 엑셀 템플릿 '{excel_template_path}'을(를) 찾을 수 없습니다.")
                break
            except Exception as e:
                print(f"! 오류: 엑셀 템플릿을 불러오는 중 오류가 발생했습니다: {e}")
                break 

            # --- 2-8. 데이터 입력 ---
            
            # 파일명에 사용할 값을 미리 가져오기
            e19_value_str = get_csv_data(all_csv_data, 2, col_idx) # Analyte

            # "Instructions" 시트
            try:
                ws_inst = wb["Instructions"]
                ws_inst['E18'] = get_csv_data(all_csv_data, 6, col_idx) # Analyst
                ws_inst['E19'] = e19_value_str # Analyte (위에서 가져온 값)
                ws_inst['E20'] = get_csv_data(all_csv_data, 3, col_idx) # Units
                ws_inst['E21'] = get_csv_data(all_csv_data, 5, col_idx) # Instrument
            except KeyError:
                print("! 오류: 'Instructions' 시트를 찾을 수 없습니다.")
                continue 

            # "Linearity" 시트
            try:
                ws_lin = wb["Linearity"]
                ws_lin['E4'] = get_csv_data(all_csv_data, 5, col_idx) # InstClass
                ws_lin['E5'] = get_csv_data(all_csv_data, 7, col_idx) # Date
            except KeyError:
                print("! 오류: 'Linearity' 시트를 찾을 수 없습니다.")
                continue

            # "Data Entry" 시트
            try:
                ws_data = wb["Data Entry"]
                
                # *** 수정됨 (1) ***: F10 -> F11
                ws_data['F11'] = get_csv_data(all_csv_data, 15, col_idx, as_number=True) # ATEPct
                ws_data['I13'] = get_csv_data(all_csv_data, 7, col_idx) # Date (문자열)
                
                # 데이터 블럭 (숫자로 입력)
                cells_32 = ['E32', 'F32', 'G32', 'H32', 'I32']
                rows_32  = [37, 38, 39, 40, 41]
                for cell, row in zip(cells_32, rows_32):
                    ws_data[cell] = get_csv_data(all_csv_data, row, col_idx, as_number=True)
                
                cells_33 = ['E33', 'F33', 'G33', 'H33', 'I33']
                rows_33  = [42, 43, 44, 45, 46]
                for cell, row in zip(cells_33, rows_33):
                    ws_data[cell] = get_csv_data(all_csv_data, row, col_idx, as_number=True)

                # 평균값 '수식' 입력
                print("  - 'Data Entry' 시트에 평균 수식 입력 중...")
                ws_data['E19'] = '=AVERAGE(E32, E33)'
                ws_data['F19'] = '=AVERAGE(F32, F33)'
                ws_data['G19'] = '=AVERAGE(G32, G33)'
                ws_data['H19'] = '=AVERAGE(H32, H33)'
                ws_data['I19'] = '=AVERAGE(I32, I33)'
                
            except KeyError:
                print("! 오류: 'Data Entry' 시트를 찾을 수 없습니다.")
                continue

            # --- 2-9. 임시 파일 저장 (수식 계산을 위해) ---
            
            # *** 수정됨 (2) ***: 파일명 생성 로직 변경
            today_str = datetime.now().strftime("%Y%m%d") # 예: 20231027
            
            # 파일명에 부적절한 문자 제거 (예: /, \, :, *)
            safe_e19_value = e19_value_str.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').strip()
            
            # 임시 파일명 (검증 전)
            temp_filename = f"{today_str}_Linearity_{safe_e19_value}_TEMP.xlsm"
            temp_save_path = os.path.join(save_directory, temp_filename)
            
            try:
                # 임시 파일로 저장 (수식 계산을 위해)
                wb.save(temp_save_path)
                print(f"  ✔ 임시 엑셀 파일 저장 완료: {temp_save_path}")
                
                # 진행바 업데이트 - 검증 단계
                progress_window.update_progress(
                    file_counter - 1, 
                    f"{file_counter}번째 파일 검증 중..."
                )
                
                # --- 2-10. 데이터 검증 수행 (저장된 파일에서) ---
                is_valid, status_suffix = validate_excel_data_with_com(temp_save_path)
                
                # --- 2-11. 최종 파일명으로 변경 ---
                final_filename = f"{today_str}_Linearity_{safe_e19_value}{status_suffix}.xlsm"
                final_save_path = os.path.join(save_directory, final_filename)
                
                # 임시 파일을 최종 파일명으로 변경
                import shutil
                shutil.move(temp_save_path, final_save_path)
                print(f"  ✔ 최종 엑셀 파일명 변경 완료: {final_save_path}")
                
                # --- 2-12. PDF 파일 생성 (Pass인 경우에만) ---
                if is_valid:
                    # 진행바 업데이트 - PDF 생성 단계
                    progress_window.update_progress(
                        file_counter - 1, 
                        f"{file_counter}번째 파일 PDF 생성 중..."
                    )
                    
                    pdf_filename = f"{today_str}_Linearity_{safe_e19_value}{status_suffix}.pdf"
                    pdf_save_path = os.path.join(save_directory, pdf_filename)
                    
                    print("  - PDF 파일 생성 중...")
                    
                    if export_first_three_sheets_to_pdf(final_save_path, pdf_save_path):
                        print(f"  ✔ PDF 파일 저장 완료: {pdf_save_path}")
                    else:
                        print("  ! PDF 파일 생성에 실패했습니다.")
                else:
                    print("  - 검증 실패로 인해 PDF 파일을 생성하지 않습니다.")
                
                # 진행바 업데이트 - 파일 완료
                progress_window.update_progress(
                    file_counter, 
                    f"{file_counter}번째 파일 완료!"
                )
                
                file_counter += 1
                
            except PermissionError:
                print(f"! 저장 오류: {temp_filename} 파일이 열려있거나 권한이 없습니다. 이 파일을 건너뜁니다.")
            except Exception as e:
                print(f"! 저장 중 알 수 없는 오류 발생: {e}")
                # 임시 파일 정리
                try:
                    if os.path.exists(temp_save_path):
                        os.remove(temp_save_path)
                except:
                    pass

        # --- 2-13. 최종 완료 ---
        progress_window.update_progress(total_files, "모든 작업 완료!")
        
        print("\n--- 작업 완료 ---")
        if file_counter == 1:
            print("생성된 파일이 없습니다. CSV C열의 2행에 데이터가 있는지 확인하세요.")
        else:
            print(f"총 {file_counter - 1}개의 엑셀 파일이 '{save_directory}' 폴더에 생성되었습니다.")
        
        # 완료 메시지 표시
        messagebox.showinfo(
            "작업 완료", 
            f"총 {file_counter - 1}개의 엑셀 파일이 생성되었습니다.\n저장 위치: {save_directory}"
        )
        
        # 진행바 창 닫기
        progress_window.close()
        
        # Excel 애플리케이션 정리
        cleanup_excel_app()

    except Exception as e:
        print(f"\n! 치명적인 오류 발생: {e}")
        
        # 오류 발생 시 진행바 창 닫기
        try:
            progress_window.close()
        except:
            pass
        
        # Excel 애플리케이션 정리
        cleanup_excel_app()
        
        messagebox.showerror("오류 발생", f"치명적인 오류가 발생했습니다:\n{e}")
    
    input("엔터 키를 누르면 프로그램이 종료됩니다...")