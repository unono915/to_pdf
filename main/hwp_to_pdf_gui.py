import os
import sys
import glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from pathlib import Path

class HwpToPdfConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("한글 파일 PDF 변환기")
        # 아이콘 설정 (ICO 파일 경로를 실제 경로로 변경하세요)
        try:
            # PNG 아이콘 사용 시
            ico = tk.PhotoImage(file="app_icon.png")
            self.root.iconphoto(True, ico)
        except Exception:
            pass

        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # 중단 요청 플래그
        self.stop_requested = False
        self.is_converting = False

        # 입력/출력 폴더 변수
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar(value="pdf_output")

        self.setup_ui()

    def setup_ui(self):
        """UI 구성"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        title_label = ttk.Label(main_frame, text="한글(.hwp) 파일 PDF 일괄 변환", 
                               font=("맑은 고딕", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        ttk.Label(main_frame, text="입력 폴더:", font=("맑은 고딕", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_folder, width=50).grid(row=1, column=1, padx=(10, 5), pady=5, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="폴더 선택", command=self.select_input_folder).grid(row=1, column=2, padx=(5, 0), pady=5)
        
        ttk.Label(main_frame, text="출력 폴더:", font=("맑은 고딕", 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_folder, width=50).grid(row=2, column=1, padx=(10, 5), pady=5, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="폴더 선택", command=self.select_output_folder).grid(row=2, column=2, padx=(5, 0), pady=5)
        
        # 변환 및 정지 버튼
        self.convert_button = ttk.Button(main_frame, text="변환 시작", command=self.start_conversion)
        self.convert_button.grid(row=3, column=1, pady=20)
        
        self.stop_button = ttk.Button(main_frame, text="정지", command=self.stop_conversion)
        self.stop_button.grid(row=3, column=2, padx=(10, 0), pady=20)
        
        self.progress_var = tk.StringVar(value="변환 대기 중...")
        ttk.Label(main_frame, textvariable=self.progress_var, font=("맑은 고딕", 10)).grid(row=4, column=0, columnspan=3, pady=5)
        
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(main_frame, text="변환 로그:", font=("맑은 고딕", 10)).grid(row=6, column=0, sticky=tk.W, pady=(20, 5))
        self.log_text = scrolledtext.ScrolledText(main_frame, height=15, width=70)
        self.log_text.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(7, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        
        self.log("한글 파일 PDF 변환기가 준비되었습니다.")
        self.log("1. 입력 폴더: 변환할 .hwp 파일들이 있는 폴더를 선택하세요.")
        self.log("2. 출력 폴더: PDF 파일이 저장될 폴더를 선택하세요.")
        self.log("3. '변환 시작' 버튼을 클릭하여 변환을 시작하세요.\n")
        
        if sys.platform != "win32":
            self.log("⚠️ 경고: 이 프로그램은 Windows에서만 작동합니다.\n")

    def select_input_folder(self):
        folder = filedialog.askdirectory(title="한글 파일이 있는 폴더를 선택하세요")
        if folder:
            self.input_folder.set(folder)
            self.log(f"입력 폴더 선택: {folder}")
            hwp_files = glob.glob(os.path.join(folder, "*.hwp"))
            self.log(f"발견된 한글 파일: {len(hwp_files)}개\n")
    
    def select_output_folder(self):
        folder = filedialog.askdirectory(title="PDF 파일을 저장할 폴더를 선택하세요")
        if folder:
            self.output_folder.set(folder)
            self.log(f"출력 폴더 선택: {folder}\n")
    
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def start_conversion(self):
        if self.is_converting:
            return
        if not self.input_folder.get():
            messagebox.showerror("오류", "입력 폴더를 선택해주세요.")
            return
        if not os.path.exists(self.input_folder.get()):
            messagebox.showerror("오류", "선택한 입력 폴더가 존재하지 않습니다.")
            return
        hwp_files = glob.glob(os.path.join(self.input_folder.get(), "*.hwp"))
        if not hwp_files:
            messagebox.showwarning("경고", "선택한 폴더에 .hwp 파일이 없습니다.")
            return

        self.is_converting = True
        self.stop_requested = False
        self.convert_button.config(state="disabled")
        self.progress_bar.start()

        thread = threading.Thread(target=self.convert_files)
        thread.daemon = True
        thread.start()

    def stop_conversion(self):
        if self.is_converting:
            self.stop_requested = True
            self.log("⏹️ 변환 중지 요청됨")

    def convert_files(self):
        try:
            self.log("=" * 50)
            self.log("변환을 시작합니다...")
            try:
                import win32com.client
                self.log("win32com 라이브러리를 사용합니다.")
                use_win32com = True
            except ImportError:
                try:
                    import comtypes.client
                    self.log("comtypes 라이브러리를 사용합니다.")
                    use_win32com = False
                except ImportError:
                    self.log("❌ 오류: 필요한 COM 라이브러리가 없습니다.")
                    self.log("win32com 또는 comtypes 라이브러리가 필요합니다.")
                    self.conversion_finished()
                    return

            output_folder = self.output_folder.get()
            os.makedirs(output_folder, exist_ok=True)
            hwp_files = glob.glob(os.path.join(self.input_folder.get(), "*.hwp"))
            total_files = len(hwp_files)
            self.log(f"총 {total_files}개의 한글 파일을 변환합니다.\n")

            if use_win32com:
                import win32com.client
                hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            else:
                import comtypes.client
                hwp = comtypes.client.CreateObject("HWPFrame.HwpObject")
            try:
                hwp.SetMessageBoxMode(0x00000020)
                hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            except:
                pass

            success_count = 0
            failed_files = []

            for i, hwp_file in enumerate(hwp_files, 1):
                if self.stop_requested:
                    self.log("⏹️ 변환을 중단합니다.")
                    break

                filename = os.path.basename(hwp_file)
                filename_without_ext = os.path.splitext(filename)[0]
                output_path = os.path.join(output_folder, f"{filename_without_ext}.pdf")

                self.progress_var.set(f"변환 중... ({i}/{total_files}) {filename}")
                self.log(f"[{i}/{total_files}] 변환 중: {filename}")

                try:
                    hwp_file_abs = os.path.abspath(hwp_file)
                    output_path_abs = os.path.abspath(output_path)
                    open_opts = "versionwarning:false;securitywarning:false;updatechecking:false;passworddlg:false;repair:false"
                    hwp.Open(hwp_file_abs, "HWP", open_opts)
                    hwp.SaveAs(output_path_abs, "PDF", "")
                    hwp.Clear(1)

                    success_count += 1
                    self.log(f"✅ {filename} 변환 완료")
                except Exception as e:
                    failed_files.append(filename)
                    self.log(f"❌ {filename} 변환 실패: {e}")
                    try:
                        hwp.Clear(1)
                    except:
                        pass

            try:
                hwp.Quit()
            except:
                pass

            self.log("\n" + "=" * 50)
            self.log(f"변환 완료: {success_count}개 성공")
            if failed_files:
                self.log(f"변환 실패: {len(failed_files)}개")
                for f in failed_files:
                    self.log(f"  - {f}")

            self.progress_var.set(f"변환 완료! 성공: {success_count}개, 실패: {len(failed_files)}개")
            if success_count > 0:
                messagebox.showinfo("변환 완료",
                    f"변환이 완료되었습니다!\n\n"
                    f"성공: {success_count}개\n"
                    f"실패: {len(failed_files)}개\n\n"
                    f"PDF 파일은 다음 폴더에 저장됨:\n{output_folder}")
        except Exception as e:
            self.log(f"❌ 전체 변환 중 오류 발생: {e}")
            self.progress_var.set("변환 중 오류 발생")
            messagebox.showerror("오류", f"변환 중 오류가 발생했습니다:\n{e}")
        finally:
            self.conversion_finished()

    def conversion_finished(self):
        self.is_converting = False
        self.convert_button.config(state="normal")
        self.progress_bar.stop()


def main():
    if sys.platform != "win32":
        messagebox.showerror("오류", "이 프로그램은 Windows에서만 실행 가능합니다.")
        return
    root = tk.Tk()
    app = HwpToPdfConverter(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    main()
