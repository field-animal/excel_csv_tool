import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import shutil

class ExcelCSVTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel & CSV 통합 관리 도구")
        self.root.geometry("500x520")
        
        # 1. 엑셀 -> CSV 변환 섹션
        frame1 = ttk.LabelFrame(root, text="1. 엑셀 파일을 CSV로 변환", padding=10)
        frame1.pack(fill="x", padx=10, pady=5)
        ttk.Button(frame1, text="엑셀 파일 선택 (복수 가능)", command=self.convert_excel).pack(fill="x")

        # 2. CSV 병합 섹션
        frame2 = ttk.LabelFrame(root, text="2. CSV 파일 병합", padding=10)
        frame2.pack(fill="both", expand=True, padx=10, pady=5)

        # 병합 옵션 선택
        self.merge_mode = tk.StringVar(value="fast")
        ttk.Radiobutton(frame2, text="빠르게 병합 (copy 명령어 방식, 헤더 제외 불가)", 
                        variable=self.merge_mode, value="fast").pack(anchor="w")
        ttk.Radiobutton(frame2, text="고급 병합 (인코딩 교정 및 헤더 제어)", 
                        variable=self.merge_mode, value="advanced").pack(anchor="w")

        # 고급 병합 전용 세부 옵션
        self.remove_header = tk.BooleanVar(value=True)
        self.header_check = ttk.Checkbutton(frame2, text="첫 줄(헤더) 제거 (첫 파일만 헤더 유지)", 
                                           variable=self.remove_header)
        self.header_check.pack(anchor="w", padx=25, pady=5)

        ttk.Button(frame2, text="CSV 파일 선택 및 병합 실행", command=self.merge_csv).pack(fill="x", pady=10)

        # 로그 출력창
        self.log_text = tk.Text(root, height=8, state="disabled", bg="#f8f9fa")
        self.log_text.pack(fill="both", padx=10, pady=5)

    def log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def convert_excel(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xls")])
        for f in files:
            try:
                df = pd.read_excel(f)
                out = os.path.splitext(f)[0] + ".csv"
                df.to_csv(out, index=False, encoding='utf-8-sig')
                self.log(f"변환 성공: {os.path.basename(out)}")
            except Exception as e: self.log(f"오류: {os.path.basename(f)} - {e}")
        messagebox.showinfo("완료", "엑셀 변환 완료")

    def merge_csv(self):
        files = filedialog.askopenfilenames(filetypes=[("CSV Files", "*.csv")])
        if not files: return
        save_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if not save_path: return

        try:
            if self.merge_mode.get() == "fast":
                # 바이너리 복사 방식 (copy /b와 동일)
                with open(save_path, 'wb') as outfile:
                    for f in files:
                        with open(f, 'rb') as infile:
                            shutil.copyfileobj(infile, outfile)
                self.log(f"빠른 병합 완료: {os.path.basename(save_path)}")
            else:
                # 고급 병합 방식 (Pandas 활용)
                with open(save_path, 'w', encoding='utf-8-sig', newline='') as f_out:
                    for i, f in enumerate(files):
                        df = pd.read_csv(f)
                        # 첫 파일이 아니고 헤더 제거 옵션이 켜져 있으면 헤더 없이 기록
                        include_header = not (self.remove_header.get() and i > 0)
                        df.to_csv(f_out, index=False, header=include_header, mode='a')
                self.log(f"고급 병합 완료: {os.path.basename(save_path)}")
            messagebox.showinfo("완료", "병합이 완료되었습니다.")
        except Exception as e: self.log(f"병합 중 오류: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    ExcelCSVTool(root)
    root.mainloop()
