"""
p-net Order Reply Excel File Comparison Tool
GUI application for comparing p-net download file with factory reply file
and generating manual upload file.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
from excel_processor import ExcelProcessor


class OrderReplyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("p-net Order Reply Tool")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        self.processor = ExcelProcessor()
        self.pnet_file = None
        self.factory_file = None
        self.output_file = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup UI components."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="p-net 납기회신 파일 생성 도구", 
                               font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # File selection section
        ttk.Label(main_frame, text="1. p-net 다운로드 파일:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.pnet_label = ttk.Label(main_frame, text="선택되지 않음", foreground="gray")
        self.pnet_label.grid(row=1, column=1, sticky=tk.W, padx=5)
        ttk.Button(main_frame, text="선택", command=self.select_pnet_file).grid(row=1, column=2, padx=5)
        
        ttk.Label(main_frame, text="2. 공장납기회신 파일:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.factory_label = ttk.Label(main_frame, text="선택되지 않음", foreground="gray")
        self.factory_label.grid(row=2, column=1, sticky=tk.W, padx=5)
        ttk.Button(main_frame, text="선택", command=self.select_factory_file).grid(row=2, column=2, padx=5)
        
        # Output file selection
        ttk.Label(main_frame, text="3. 저장 경로:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.output_label = ttk.Label(main_frame, text="선택되지 않음", foreground="gray")
        self.output_label.grid(row=3, column=1, sticky=tk.W, padx=5)
        ttk.Button(main_frame, text="선택", command=self.select_output_file).grid(row=3, column=2, padx=5)
        
        # Process button
        ttk.Button(main_frame, text="파일 처리 및 생성", command=self.process_files).grid(
            row=4, column=0, columnspan=3, pady=20, sticky=(tk.W, tk.E))
        
        # Progress/Status section
        ttk.Label(main_frame, text="처리 상태:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.status_text = tk.Text(main_frame, height=10, width=70)
        self.status_text.grid(row=6, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.grid(row=6, column=3, sticky=(tk.N, tk.S))
        self.status_text.config(yscrollcommand=scrollbar.set)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
    
    def select_pnet_file(self):
        """Select p-net download file."""
        file_path = filedialog.askopenfilename(
            title="p-net 다운로드 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.pnet_file = file_path
            filename = Path(file_path).name
            self.pnet_label.config(text=filename, foreground="black")
            self.log_status(f"✓ p-net 파일 선택: {filename}")
    
    def select_factory_file(self):
        """Select factory reply file."""
        file_path = filedialog.askopenfilename(
            title="공장납기회신 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.factory_file = file_path
            filename = Path(file_path).name
            self.factory_label.config(text=filename, foreground="black")
            self.log_status(f"✓ 공장납기회신 파일 선택: {filename}")
    
    def select_output_file(self):
        """Select output file location."""
        file_path = filedialog.asksaveasfilename(
            title="저장할 파일 선택",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.output_file = file_path
            filename = Path(file_path).name
            self.output_label.config(text=filename, foreground="black")
            self.log_status(f"✓ 저장 경로 선택: {filename}")
    
    def process_files(self):
        """Process files in a separate thread."""
        if not self.pnet_file or not self.factory_file or not self.output_file:
            messagebox.showerror("오류", "모든 파일을 선택해주세요.")
            return
        
        # Disable process button during processing
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.config(state=tk.DISABLED)
        
        # Run processing in separate thread
        thread = threading.Thread(target=self._process_files_thread)
        thread.start()
    
    def _process_files_thread(self):
        """Process files in background thread."""
        try:
            self.log_status("처리 시작...")
            self.log_status("")
            
            # Read files
            self.log_status("1/3: p-net 다운로드 파일 읽기...")
            if not self.processor.read_pnet_download(self.pnet_file):
                raise Exception("p-net 파일 읽기 실패")
            self.log_status("✓ p-net 파일 읽기 완료")
            
            self.log_status("2/3: 공장납기회신 파일 읽기...")
            if not self.processor.read_factory_reply(self.factory_file):
                raise Exception("공장납기회신 파일 읽기 실패")
            self.log_status("✓ 공장납기회신 파일 읽기 완료")
            
            self.log_status("3/3: 파일 비교 및 결과 생성...")
            if not self.processor.compare_and_generate():
                raise Exception("파일 비교 실패")
            self.log_status("✓ 파일 비교 완료")
            
            self.log_status("4/4: 결과 저장...")
            if not self.processor.save_result(self.output_file):
                raise Exception("결과 저장 실패")
            self.log_status("✓ 결과 저장 완료")
            
            self.log_status("")
            self.log_status("성공! 파일이 생성되었습니다.")
            self.log_status(f"저장 위치: {self.output_file}")
            
            messagebox.showinfo("성공", "파일 처리가 완료되었습니다!")
        
        except Exception as e:
            self.log_status("")
            self.log_status(f"오류 발생: {str(e)}")
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{str(e)}")
        
        finally:
            # Re-enable process button
            for widget in self.root.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.config(state=tk.NORMAL)
    
    def log_status(self, message: str):
        """Log status message to status text widget."""
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update()


def main():
    try:
        root = tk.Tk()
        app = OrderReplyApp(root)
        print("GUI 창이 열립니다. 창이 보이지 않으면 작업 표시줄을 확인하세요.")
        root.mainloop()
    except Exception as e:
        print(f"GUI 실행 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
