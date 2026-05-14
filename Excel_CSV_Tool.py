import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import shutil
import threading  # 멀티스레딩을 위한 라이브러리 추가

class ExcelCSVTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/CSV Advanced Tool (K1PS)")
        self.root.geometry("550x620")
        
        # 스타일 설정
        style = ttk.Style()
        style.configure("TButton", padding=5)
        
        # --- 1. 엑셀 -> CSV 변환 섹션 ---
        frame1 = ttk.LabelFrame(root, text="1. 엑셀 -> CSV 변환 (복수 선택 가능)", padding=10)
        frame1.pack(fill="x", padx=10, pady=5)
        
        # 인코딩 선택 레이아웃
        enc_layout = ttk.Frame(frame1)
        enc_layout.pack(fill="x", pady=5)
        ttk.Label(enc_layout, text="변환 인코딩 설정:").pack(side="left")
        
        # 3가지 인코딩 옵션 추가
        self.enc_option = ttk.Combobox(enc_layout, 
                                       values=["utf-8-sig (엑셀 호환/권장)", "utf-8 (표준)", "cp949 (오래된 한글)"], 
                                       state="readonly", width=25)
        self.enc_option.current(0)
        self.enc_option.pack(side="left", padx=5)
        
        ttk.Button(frame1, text="엑셀 파일들을 선택하여 변환", command=self.convert_excel).pack(fill="x", pady=5)
        
        # --- 2. CSV 파일 병합 섹션 ---
        frame2 = ttk.LabelFrame(root, text="2. CSV 파일 병합 (복수 선택 가능)", padding=10)
        frame2.pack(fill="both", expand=True, padx=10, pady=5)
        
        ttk.Label(frame2, text="병합 방식을 선택하세요:").pack(anchor="w", pady=2)
        
        self.merge_mode = tk.StringVar(value="fast")
        ttk.Radiobutton(frame2, text="빠른 병합 (단순 이어붙이기 / 헤더 포함)", 
                        variable=self.merge_mode, value="fast").pack(anchor="w")
        ttk.Radiobutton(frame2, text="고급 병합 (인코딩 교정 및 헤더 제어)", 
                        variable=self.merge_mode, value="advanced").pack(anchor="w")
        
        self.remove_header = tk.BooleanVar(value=True)
        self.header_check = ttk.Checkbutton(frame2, text="첫 줄(헤더) 중복 제거 (첫 파일만 헤더 유지)", 
                                           variable=self.remove_header)
        self.header_check.pack(anchor="w", padx=20, pady=5)
        
        ttk.Button(frame2, text="CSV 파일들을 선택하여 병합", command=self.merge_csv).pack(fill="x", pady=10)
        
        # 로그 출력창
        self.log_text = tk.Text(root, height=12, state="disabled", bg="#f8f9fa", font=("Consolas", 9))
        self.log_text.pack(fill="both", padx=10, pady=5)

    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def convert_excel(self):
        files = filedialog.askopenfilenames(title="변환할 엑셀 파일 선택", filetypes=[("Excel files", "*.xlsx *.xls")])
        if not files: return
        
        # 버튼을 비활성화하여 중복 실행 방지
        # self.convert_btn.config(state="disabled") 
        
        # 별도의 스레드에서 실제 변환 작업 실행
        thread = threading.Thread(target=self._run_conversion, args=(files,))
        thread.daemon = True # 프로그램 종료 시 스레드도 함께 종료
        thread.start()

    def _run_conversion(self, files):
        total_files = len(files)
        self.log(f"--- 변환 시작 (총 {total_files}개 파일) ---")

        enc_map = {
            "utf-8-sig (엑셀 호환/권장)": "utf-8-sig",
            "utf-8 (표준)": "utf-8",
            "cp949 (오래된 한글)": "cp949"
        }
        selected_enc = enc_map.get(self.enc_option.get(), "utf-8-sig")

        for i, f in enumerate(files, 1):
            try:
                self.log(f"\n[{i} / {total_files}] 처리 중: {os.path.basename(f)}")
                
                # Pandas 작업 실행 (이 부분이 돌아가는 동안에도 화면은 안 멈춤)
                # df = pd.read_excel(f) # pandas가 지능형 숫자 변환. 00123 -> 123
                df = pd.read_excel(f, dtype=str) # pandas가 지능형 숫자 변환 안함. 000123 -> "000123"
                
                out_name = os.path.splitext(f)[0] + f"_{selected_enc}.csv"
                df.to_csv(out_name, index=False, encoding=selected_enc)
                
                self.log(f" > 완료: {os.path.basename(out_name)}")
            except Exception as e:
                self.log(f" > [오류]: {e}")
        
        self.log("\n--- 모든 작업 완료 ---")
        messagebox.showinfo("알림", "작업이 끝났습니다.")
        # self.convert_btn.config(state="normal")

    def merge_csv(self):
        files = filedialog.askopenfilenames(title="병합할 CSV 파일 선택", filetypes=[("CSV files", "*.csv")])
        if not files: return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", title="저장할 파일 이름 입력")
        if not save_path: return
        
        try:
            if self.merge_mode.get() == "fast":
                with open(save_path, 'wb') as outfile:
                    for f in files:
                        with open(f, 'rb') as infile:
                            shutil.copyfileobj(infile, outfile)
            else:
                # 고급 병합 시 한글 깨짐 방지를 위해 결과물은 항상 utf-8-sig로 저장
                with open(save_path, 'w', encoding='utf-8-sig', newline='') as f_out:
                    for i, f in enumerate(files):
                        try:
                            # 다양한 인코딩 대응을 위해 시도
                            try:
                                df = pd.read_csv(f, encoding='utf-8-sig')
                            except:
                                df = pd.read_csv(f, encoding='cp949')
                            
                            include_header = not (self.remove_header.get() and i > 0)
                            df.to_csv(f_out, index=False, header=include_header, mode='a')
                        except Exception as e:
                            self.log(f"[경고] {os.path.basename(f)} 처리 실패: {e}")
            
            self.log(f"[성공] 병합 완료: {os.path.basename(save_path)}")
            messagebox.showinfo("알림", "병합 작업이 완료되었습니다.")
        except Exception as e:
            self.log(f"[오류] 병합 중 중대 오류: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCSVTool(root)
    root.mainloop()
