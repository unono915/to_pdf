import os
import sys
import glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from pathlib import Path

class HwpWordToPdfConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("한글/워드 파일 PDF 변환기")
        # 아이콘 설정 (ICO 파일 경로를 실제 경로로 변경하세요)
        try:
            self.root.iconbitmap("app_icon.ico")
        except Exception:
            pass

        self.root.geometry("650x550")
        self.root.resizable(True, True)
        
        # 중단 요청 플래그
        self.stop_requested = False
        self.is_converting = False

        # 입력/출력 폴더 변수
        self.input_folder = tk.StringVar(value="폴더를 설정하세요.")
        self.output_folder = tk.StringVar(value="폴더를 설정하세요.")
        
        # 폴더 선택 상태 추적
        self.input_folder_selected = False
        self.output_folder_selected = False

        self.setup_ui()

    def setup_ui(self):
        """UI 구성"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        title_label = ttk.Label(main_frame, text="한글/워드 파일 PDF 일괄 변환", 
                               font=("맑은 고딕", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 폴더 선택 부분
        ttk.Label(main_frame, text="입력 폴더:", font=("맑은 고딕", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_folder, width=50).grid(row=1, column=1, padx=(10, 5), pady=5, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="폴더 선택", command=self.select_input_folder).grid(row=1, column=2, padx=(5, 0), pady=5)
        
        ttk.Label(main_frame, text="출력 폴더:", font=("맑은 고딕", 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_folder, width=50).grid(row=2, column=1, padx=(10, 5), pady=5, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="폴더 선택", command=self.select_output_folder).grid(row=2, column=2, padx=(5, 0), pady=5)
        
        # 변환 버튼들
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=20)
                
        self.all_convert_button = ttk.Button(button_frame, text="일괄 변환", command=self.start_all_conversion, state="disabled")
        self.all_convert_button.grid(row=0, column=0, padx=5)

        self.hwp_convert_button = ttk.Button(button_frame, text="한글 변환", command=self.start_hwp_conversion, state="disabled")
        self.hwp_convert_button.grid(row=0, column=1, padx=5)
        
        self.word_convert_button = ttk.Button(button_frame, text="워드 변환", command=self.start_word_conversion, state="disabled")
        self.word_convert_button.grid(row=0, column=2, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="정지", command=self.stop_conversion)
        self.stop_button.grid(row=0, column=3, padx=5)
        
        # 진행 상태 표시
        self.progress_var = tk.StringVar(value="변환 대기 중...")
        ttk.Label(main_frame, textvariable=self.progress_var, font=("맑은 고딕", 10)).grid(row=4, column=0, columnspan=3, pady=5)
        
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 로그 표시
        ttk.Label(main_frame, text="변환 로그:", font=("맑은 고딕", 10)).grid(row=6, column=0, sticky=tk.W, pady=(20, 5))
        self.log_text = scrolledtext.ScrolledText(main_frame, height=15, width=70)
        self.log_text.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 그리드 설정
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(7, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        
        # 초기 메시지
        self.log("기능추가 및 기타 문의:yoonhojiji@gmail.com")
        self.log("한글/워드 파일 PDF 변환기가 준비되었습니다.")
        self.log("1. 입력 폴더: 변환할 .hwp/.hwpx/.doc/.docx 파일들이 있는 폴더를 선택하세요.")
        self.log("2. 출력 폴더: PDF 파일이 저장될 폴더를 선택하세요.")
        self.log("3. 입력/출력 폴더를 모두 선택하면 변환 버튼이 활성화됩니다.")
        self.log("4. '한글 변환', '워드 변환', 또는 '일괄 변환' 버튼을 클릭하여 변환을 시작하세요.\n")
        
        if sys.platform != "win32":
            self.log("⚠️ 경고: 이 프로그램은 Windows에서만 작동합니다.\n")

    def select_input_folder(self):
        folder = filedialog.askdirectory(title="변환할 파일이 있는 폴더를 선택하세요")
        if folder:
            self.input_folder.set(folder)
            self.input_folder_selected = True
            self.log(f"입력 폴더 선택: {folder}")
            
            # 한글 파일 개수 확인 (.hwp, .hwpx)
            hwp_files = glob.glob(os.path.join(folder, "*.hwp"))
            hwpx_files = glob.glob(os.path.join(folder, "*.hwpx"))
            total_hwp_files = len(hwp_files + hwpx_files)
            
            # 워드 파일 개수 확인 (.doc, .docx)
            doc_files = glob.glob(os.path.join(folder, "*.doc"))
            docx_files = glob.glob(os.path.join(folder, "*.docx"))
            total_word_files = len(doc_files + docx_files)
            
            self.log(f"발견된 한글 파일: {total_hwp_files}개")
            self.log(f"발견된 워드 파일: {total_word_files}개\n")
            
            # 버튼 상태 업데이트
            self.update_button_state()
    
    def select_output_folder(self):
        folder = filedialog.askdirectory(title="PDF 파일을 저장할 폴더를 선택하세요")
        if folder:
            self.output_folder.set(folder)
            self.output_folder_selected = True
            self.log(f"출력 폴더 선택: {folder}\n")
            
            # 버튼 상태 업데이트
            self.update_button_state()
    
    def update_button_state(self):
        """입력/출력 폴더 선택 상태에 따라 변환 버튼 활성화/비활성화"""
        if self.input_folder_selected and self.output_folder_selected and not self.is_converting:
            self.hwp_convert_button.config(state="normal")
            self.word_convert_button.config(state="normal")
            self.all_convert_button.config(state="normal")
        else:
            self.hwp_convert_button.config(state="disabled")
            self.word_convert_button.config(state="disabled")
            self.all_convert_button.config(state="disabled")
    
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def start_hwp_conversion(self):
        if self.is_converting:
            return
        if not self.input_folder.get():
            messagebox.showerror("오류", "입력 폴더를 선택해주세요.")
            return
        if not os.path.exists(self.input_folder.get()):
            messagebox.showerror("오류", "선택한 입력 폴더가 존재하지 않습니다.")
            return
            
        hwp_files = glob.glob(os.path.join(self.input_folder.get(), "*.hwp"))
        hwpx_files = glob.glob(os.path.join(self.input_folder.get(), "*.hwpx"))
        total_hwp_files = hwp_files + hwpx_files
        
        if not total_hwp_files:
            messagebox.showwarning("경고", "선택한 폴더에 .hwp/.hwpx 파일이 없습니다.")
            return

        self.is_converting = True
        self.stop_requested = False
        self.hwp_convert_button.config(state="disabled")
        self.word_convert_button.config(state="disabled")
        self.all_convert_button.config(state="disabled")
        self.progress_bar.start()

        thread = threading.Thread(target=lambda: self.convert_files("hwp"))
        thread.daemon = True
        thread.start()

    def start_word_conversion(self):
        if self.is_converting:
            return
        if not self.input_folder.get():
            messagebox.showerror("오류", "입력 폴더를 선택해주세요.")
            return
        if not os.path.exists(self.input_folder.get()):
            messagebox.showerror("오류", "선택한 입력 폴더가 존재하지 않습니다.")
            return
            
        doc_files = glob.glob(os.path.join(self.input_folder.get(), "*.doc"))
        docx_files = glob.glob(os.path.join(self.input_folder.get(), "*.docx"))
        word_files = doc_files + docx_files
        
        if not word_files:
            messagebox.showwarning("경고", "선택한 폴더에 .doc/.docx 파일이 없습니다.")
            return

        self.is_converting = True
        self.stop_requested = False
        self.hwp_convert_button.config(state="disabled")
        self.word_convert_button.config(state="disabled")
        self.all_convert_button.config(state="disabled")
        self.progress_bar.start()

        thread = threading.Thread(target=lambda: self.convert_files("word"))
        thread.daemon = True
        thread.start()

    def start_all_conversion(self):
        if self.is_converting:
            return
        if not self.input_folder.get():
            messagebox.showerror("오류", "입력 폴더를 선택해주세요.")
            return
        if not os.path.exists(self.input_folder.get()):
            messagebox.showerror("오류", "선택한 입력 폴더가 존재하지 않습니다.")
            return
            
        # 한글 파일과 워드 파일 검색
        hwp_files = glob.glob(os.path.join(self.input_folder.get(), "*.hwp"))
        hwpx_files = glob.glob(os.path.join(self.input_folder.get(), "*.hwpx"))
        total_hwp_files = hwp_files + hwpx_files
        
        doc_files = glob.glob(os.path.join(self.input_folder.get(), "*.doc"))
        docx_files = glob.glob(os.path.join(self.input_folder.get(), "*.docx"))
        word_files = doc_files + docx_files
        
        if not total_hwp_files and not word_files:
            messagebox.showwarning("경고", "선택한 폴더에 변환할 파일이 없습니다.")
            return

        self.is_converting = True
        self.stop_requested = False
        self.hwp_convert_button.config(state="disabled")
        self.word_convert_button.config(state="disabled")
        self.all_convert_button.config(state="disabled")
        self.progress_bar.start()

        thread = threading.Thread(target=lambda: self.convert_files("all"))
        thread.daemon = True
        thread.start()

    def stop_conversion(self):
        if self.is_converting:
            self.stop_requested = True
            self.log("⏹️ 변환 중지 요청됨")

    def convert_files(self, file_type):
        try:
            self.log("=" * 50)
            
            output_folder = self.output_folder.get()
            os.makedirs(output_folder, exist_ok=True)
            
            if file_type == "hwp":
                hwp_files = glob.glob(os.path.join(self.input_folder.get(), "*.hwp"))
                hwpx_files = glob.glob(os.path.join(self.input_folder.get(), "*.hwpx"))
                files_to_convert = hwp_files + hwpx_files
                self.log(f"한글 파일 변환을 시작합니다...")
                self.convert_hwp_files(files_to_convert, output_folder)
            elif file_type == "word":
                doc_files = glob.glob(os.path.join(self.input_folder.get(), "*.doc"))
                docx_files = glob.glob(os.path.join(self.input_folder.get(), "*.docx"))
                files_to_convert = doc_files + docx_files
                self.log(f"워드 파일 변환을 시작합니다...")
                self.convert_word_files(files_to_convert, output_folder)
            elif file_type == "all":
                self.log(f"일괄 변환을 시작합니다...")
                
                total_hwp_success = 0
                total_hwp_failed = []
                total_word_success = 0
                total_word_failed = []
                
                # 한글 파일 먼저 변환
                hwp_files = glob.glob(os.path.join(self.input_folder.get(), "*.hwp"))
                hwpx_files = glob.glob(os.path.join(self.input_folder.get(), "*.hwpx"))
                total_hwp_files = hwp_files + hwpx_files
                
                if total_hwp_files and not self.stop_requested:
                    self.log(f"한글 파일 변환 시작...")
                    total_hwp_success, total_hwp_failed = self.convert_hwp_files(total_hwp_files, output_folder, show_result=False)
                
                # 워드 파일 변환
                doc_files = glob.glob(os.path.join(self.input_folder.get(), "*.doc"))
                docx_files = glob.glob(os.path.join(self.input_folder.get(), "*.docx"))
                word_files = doc_files + docx_files
                
                if word_files and not self.stop_requested:
                    self.log(f"워드 파일 변환 시작...")
                    total_word_success, total_word_failed = self.convert_word_files(word_files, output_folder, show_result=False)
                
                # 일괄 변환 결과 표시
                self.show_batch_conversion_result(total_hwp_success, total_hwp_failed, 
                                                total_word_success, total_word_failed, output_folder)
                    
        except Exception as e:
            self.log(f"❌ 전체 변환 중 오류 발생: {e}")
            self.progress_var.set("변환 중 오류 발생")
            messagebox.showerror("오류", f"변환 중 오류가 발생했습니다:\n{e}")
        finally:
            self.conversion_finished()

    def convert_hwp_files(self, hwp_files, output_folder, show_result=True):
        try:
            import win32com.client
            import win32gui
            import win32con
            import time
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
                return 0, []

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

        def handle_permission_dialog():
            """권한 관련 대화상자 처리"""
            try:
                # 권한 대화상자 찾기
                dialog_titles = ["경고", "알림", "Warning", "Alert", "한글"]
                for title in dialog_titles:
                    hwnd = win32gui.FindWindow(None, title)
                    if hwnd:
                        # N키 누르기 (모두 허용)
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.1)
                        win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, ord('N'), 0)
                        win32gui.SendMessage(hwnd, win32con.WM_KEYUP, ord('N'), 0)
                        return True
            except Exception:
                pass
            return False

        def watch_permission_dialog(stop_event):
            """권한 팝업이 나타날 때까지 대기 후 N 키를 눌러 닫음"""
            while not stop_event.is_set():
                if handle_permission_dialog():
                    break
                time.sleep(0.3)

        for i, hwp_file in enumerate(hwp_files, 1):
            if self.stop_requested:
                self.log("⏹️ 변환을 중단합니다.")
                break

            filename = os.path.basename(hwp_file)
            filename_without_ext = os.path.splitext(filename)[0]
            output_path = os.path.join(output_folder, f"{filename_without_ext}.pdf")
            file_ext = os.path.splitext(filename)[1].lower()

            self.progress_var.set(f"한글 파일 변환 중... ({i}/{total_files}) {filename}")
            self.log(f"[{i}/{total_files}] 변환 중: {filename}")

            try:
                hwp_file_abs = os.path.abspath(hwp_file)
                output_path_abs = os.path.abspath(output_path)
                open_opts = "versionwarning:false;securitywarning:false;updatechecking:false;passworddlg:false;repair:false"
                
                stop_event = threading.Event()
                watcher = threading.Thread(target=watch_permission_dialog, args=(stop_event,))
                watcher.daemon = True
                watcher.start()

                # HWPX 파일은 다른 형식으로 열기
                if file_ext == '.hwpx':
                    hwp.Open(hwp_file_abs, "HWPX", open_opts)
                else:
                    hwp.Open(hwp_file_abs, "HWP", open_opts)

                stop_event.set()
                watcher.join(timeout=0.1)

                # 추가로 권한 대화상자가 남아있는지 확인
                for _ in range(20):  # 약 6초 동안 재확인
                    if handle_permission_dialog():
                        time.sleep(0.3)
                    else:
                        break
                    
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

        if show_result:
            self.show_conversion_result("한글", success_count, failed_files, output_folder)
        
        return success_count, failed_files

    def convert_word_files(self, word_files, output_folder, show_result=True):
        try:
            import win32com.client
            self.log("Microsoft Word COM 객체를 사용합니다.")
        except ImportError:
            self.log("❌ 오류: win32com 라이브러리가 없습니다.")
            self.log("pip install pywin32로 설치해주세요.")
            return 0, []

        total_files = len(word_files)
        self.log(f"총 {total_files}개의 워드 파일을 변환합니다.\n")

        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
        except Exception as e:
            self.log(f"❌ Microsoft Word 실행 실패: {e}")
            self.log("Microsoft Word가 설치되어 있는지 확인해주세요.")
            return 0, []

        success_count = 0
        failed_files = []

        for i, word_file in enumerate(word_files, 1):
            if self.stop_requested:
                self.log("⏹️ 변환을 중단합니다.")
                break

            filename = os.path.basename(word_file)
            filename_without_ext = os.path.splitext(filename)[0]
            output_path = os.path.join(output_folder, f"{filename_without_ext}.pdf")

            self.progress_var.set(f"워드 파일 변환 중... ({i}/{total_files}) {filename}")
            self.log(f"[{i}/{total_files}] 변환 중: {filename}")

            try:
                word_file_abs = os.path.abspath(word_file)
                output_path_abs = os.path.abspath(output_path)
                
                doc = word_app.Documents.Open(word_file_abs)
                doc.SaveAs2(output_path_abs, FileFormat=17)  # 17 = PDF format
                doc.Close()

                success_count += 1
                self.log(f"✅ {filename} 변환 완료")
            except Exception as e:
                failed_files.append(filename)
                self.log(f"❌ {filename} 변환 실패: {e}")
                try:
                    if 'doc' in locals():
                        doc.Close()
                except:
                    pass

        try:
            word_app.Quit()
        except:
            pass

        if show_result:
            self.show_conversion_result("워드", success_count, failed_files, output_folder)
        
        return success_count, failed_files

    def show_batch_conversion_result(self, hwp_success, hwp_failed, word_success, word_failed, output_folder):
        """일괄 변환 결과 표시"""
        self.log("\n" + "=" * 50)
        self.log("일괄 변환 완료!")
        
        total_success = hwp_success + word_success
        total_failed = len(hwp_failed) + len(word_failed)
        
        if hwp_success > 0 or hwp_failed:
            self.log(f"한글 파일: {hwp_success}개 성공, {len(hwp_failed)}개 실패")
            if hwp_failed:
                for f in hwp_failed:
                    self.log(f"  - {f}")
        
        if word_success > 0 or word_failed:
            self.log(f"워드 파일: {word_success}개 성공, {len(word_failed)}개 실패")
            if word_failed:
                for f in word_failed:
                    self.log(f"  - {f}")
        
        self.log(f"전체 결과: {total_success}개 성공, {total_failed}개 실패")
        
        self.progress_var.set(f"일괄 변환 완료! 성공: {total_success}개, 실패: {total_failed}개")
        
        if total_success > 0:
            messagebox.showinfo("일괄 변환 완료",
                f"일괄 변환이 완료되었습니다!\n\n"
                f"한글 파일: {hwp_success}개 성공, {len(hwp_failed)}개 실패\n"
                f"워드 파일: {word_success}개 성공, {len(word_failed)}개 실패\n"
                f"전체: {total_success}개 성공, {total_failed}개 실패\n\n"
                f"PDF 파일은 다음 폴더에 저장됨:\n{output_folder}")

    def show_conversion_result(self, file_type, success_count, failed_files, output_folder):
        self.log("\n" + "=" * 50)
        self.log(f"{file_type} 파일 변환 완료: {success_count}개 성공")
        if failed_files:
            self.log(f"변환 실패: {len(failed_files)}개")
            for f in failed_files:
                self.log(f"  - {f}")

        self.progress_var.set(f"{file_type} 변환 완료! 성공: {success_count}개, 실패: {len(failed_files)}개")
        if success_count > 0:
            messagebox.showinfo("변환 완료",
                f"{file_type} 파일 변환이 완료되었습니다!\n\n"
                f"성공: {success_count}개\n"
                f"실패: {len(failed_files)}개\n\n"
                f"PDF 파일은 다음 폴더에 저장됨:\n{output_folder}")

    def conversion_finished(self):
        self.is_converting = False
        self.update_button_state()  # 변환 완료 후 버튼 상태 업데이트
        self.progress_bar.stop()


def main():
    if sys.platform != "win32":
        messagebox.showerror("오류", "이 프로그램은 Windows에서만 실행 가능합니다.")
        return
    root = tk.Tk()
    app = HwpWordToPdfConverter(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    main()