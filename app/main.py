import os
import sys
import yaml
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from docx import Document
from app.config import load_config, save_config, AppConfig
from app.llm.openai_compatible import OpenAICompatibleProvider
from app.llm.gemini_cli import GeminiCLIProvider
from app.core.learning_manager import LearningManager
from app.pipeline.light_fix_flow import LightFixPipeline
from app.pipeline.heavy_fix_flow import HeavyFixPipeline

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class AssignmentFixerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AssignmentFixer — NEU Academic Assignment Editor")
        self.geometry("1100x720")
        self.minsize(950, 600)

        self.config = load_config()
        self._init_provider()

        self.current_output_path = None
        self._load_available_courses()
        self._build_ui()

    def _init_provider(self):
        if self.config.default_provider == "gemini_cli":
            cli_cmd = self.config.gemini_cli_path or "gemini.cmd"
            self.provider = GeminiCLIProvider(cli_cmd=cli_cmd, api_key=self.config.gemini_api_key)
        else:
            self.provider = OpenAICompatibleProvider(
                api_key=self.config.kira_api_key,
                base_url=self.config.kira_base_url,
                model=self.config.kira_model
            )
        self.learner = LearningManager(data_dir=self.config.gdrive_data_dir, llm_provider=self.provider)
        self.light_pipeline = LightFixPipeline(llm_provider=self.provider, learning_manager=self.learner)
        self.heavy_pipeline = HeavyFixPipeline(llm_provider=self.provider, learning_manager=self.learner)

    def _load_available_courses(self):
        courses_dir = os.path.join(self.config.gdrive_data_dir, "courses")
        self.courses = {}
        if os.path.exists(courses_dir):
            for c_code in os.listdir(courses_dir):
                p_yaml = os.path.join(courses_dir, c_code, "preset.yaml")
                if os.path.exists(p_yaml):
                    try:
                        with open(p_yaml, "r", encoding="utf-8") as f:
                            data = yaml.safe_load(f)
                            self.courses[c_code] = data
                    except Exception:
                        pass
        if not self.courses:
            self.courses["BS"] = {"course_name": "Business Strategy", "course_code": "BS"}

    def _build_ui(self):
        # Grid layout (1 row, 2 cols: sidebar + main content)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 1. Sidebar Navigation
        self.sidebar_frame = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)

        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame,
            text="AssignmentFixer",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.sublogo_label = ctk.CTkLabel(
            self.sidebar_frame,
            text="NEU BTEC Academic Tool",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.sublogo_label.grid(row=1, column=0, padx=20, pady=(0, 25))

        self.btn_tab_fix = ctk.CTkButton(
            self.sidebar_frame, text="Sửa Bài Assignment",
            command=lambda: self._select_tab("fix")
        )
        self.btn_tab_fix.grid(row=2, column=0, padx=20, pady=10)

        self.btn_tab_learn = ctk.CTkButton(
            self.sidebar_frame, text="Dạy Tool & Style Guide",
            command=lambda: self._select_tab("learn")
        )
        self.btn_tab_learn.grid(row=3, column=0, padx=20, pady=10)

        self.btn_tab_settings = ctk.CTkButton(
            self.sidebar_frame, text="Cài Đặt Hệ Thống",
            command=lambda: self._select_tab("settings")
        )
        self.btn_tab_settings.grid(row=4, column=0, padx=20, pady=10)

        # 2. Main Content Views (Frames)
        self.view_fix = self._create_fix_view()
        self.view_learn = self._create_learn_view()
        self.view_settings = self._create_settings_view()

        self._select_tab("fix")

    def _select_tab(self, tab_name: str):
        self.view_fix.grid_forget()
        self.view_learn.grid_forget()
        self.view_settings.grid_forget()

        if tab_name == "fix":
            self.view_fix.grid(row=0, column=1, sticky="nsew", padx=25, pady=25)
        elif tab_name == "learn":
            self.view_learn.grid(row=0, column=1, sticky="nsew", padx=25, pady=25)
        elif tab_name == "settings":
            self.view_settings.grid(row=0, column=1, sticky="nsew", padx=25, pady=25)

    # ---------------- VIEW 1: SỬA BÀI ----------------
    def _create_fix_view(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self)
        frame.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(frame, text="Sửa & Nâng Cấp Bài Assignment", font=ctk.CTkFont(size=22, weight="bold"))
        title.grid(row=0, column=0, sticky="w", pady=(0, 15))

        # File Selectors Card
        card_file = ctk.CTkFrame(frame)
        card_file.grid(row=1, column=0, sticky="ew", pady=(0, 15), padx=5)
        card_file.grid_columnconfigure(1, weight=1)

        # Paper input
        ctk.CTkLabel(card_file, text="Bài sinh viên (.docx):", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=15, pady=10, sticky="w")
        self.entry_paper = ctk.CTkEntry(card_file, placeholder_text="Chọn tệp bài làm sinh viên...")
        self.entry_paper.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(card_file, text="Chọn tệp...", width=90, command=self._browse_paper).grid(row=0, column=2, padx=15, pady=10)

        # Brief input
        ctk.CTkLabel(card_file, text="Đề bài / Rubric (.docx):", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, padx=15, pady=10, sticky="w")
        self.entry_brief = ctk.CTkEntry(card_file, placeholder_text="Chọn tệp đề bài & thang điểm...")
        self.entry_brief.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(card_file, text="Chọn tệp...", width=90, command=self._browse_brief).grid(row=1, column=2, padx=15, pady=10)

        # Course and Mode selection
        ctk.CTkLabel(card_file, text="Môn học (Course):", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, padx=15, pady=10, sticky="w")
        course_names = [f"{k} — {v.get('course_name', '')}" for k, v in self.courses.items()]
        self.combo_course = ctk.CTkComboBox(card_file, values=course_names)
        self.combo_course.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(card_file, text="Chế độ sửa:", font=ctk.CTkFont(weight="bold")).grid(row=3, column=0, padx=15, pady=10, sticky="w")
        self.mode_var = tk.StringVar(value="auto")
        mode_frame = ctk.CTkFrame(card_file, fg_color="transparent")
        mode_frame.grid(row=3, column=1, columnspan=2, padx=10, pady=10, sticky="w")

        ctk.CTkRadioButton(mode_frame, text="Tự động đề xuất", variable=self.mode_var, value="auto").pack(side="left", padx=(0, 15))
        ctk.CTkRadioButton(mode_frame, text="Sửa nhẹ (Track Changes + Comments)", variable=self.mode_var, value="light").pack(side="left", padx=(0, 15))
        ctk.CTkRadioButton(mode_frame, text="Sửa nặng (Viết lại + Bổ sung Rubric)", variable=self.mode_var, value="heavy").pack(side="left")

        # Action Buttons
        self.btn_start = ctk.CTkButton(
            frame, text="BẮT ĐẦU SỬA BÀI",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=42, fg_color="#1f6aa5", hover_color="#144870",
            command=self._start_fix_pipeline_thread
        )
        self.btn_start.grid(row=2, column=0, sticky="ew", pady=(5, 15), padx=5)

        # Progress and Log Console
        self.log_textbox = ctk.CTkTextbox(frame, height=220, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_textbox.grid(row=3, column=0, sticky="nsew", padx=5, pady=(0, 15))
        self.log_textbox.insert("end", "[Hệ thống] AssignmentFixer đã sẵn sàng.\n[Hệ thống] Hãy chọn tệp bài làm sinh viên và đề bài để bắt đầu.\n")

        # Open Output Button
        self.btn_open_word = ctk.CTkButton(
            frame, text="Mở File Kết Quả Trong Microsoft Word",
            state="disabled", height=38, fg_color="#2e7d32", hover_color="#1b5e20",
            command=self._open_in_word
        )
        self.btn_open_word.grid(row=4, column=0, sticky="ew", padx=5)

        return frame

    def _browse_paper(self):
        f = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx"), ("PDF Documents", "*.pdf")])
        if f:
            self.entry_paper.delete(0, "end")
            self.entry_paper.insert(0, f)

    def _browse_brief(self):
        f = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx"), ("PDF Documents", "*.pdf")])
        if f:
            self.entry_brief.delete(0, "end")
            self.entry_brief.insert(0, f)

    def _log(self, msg: str):
        self.log_textbox.insert("end", msg + "\n")
        self.log_textbox.see("end")

    def _start_fix_pipeline_thread(self):
        paper_path = self.entry_paper.get().strip()
        brief_path = self.entry_brief.get().strip()
        if not os.path.exists(paper_path):
            messagebox.showerror("Thiếu tệp", "Vui lòng chọn tệp bài làm của sinh viên!")
            return
        if not os.path.exists(brief_path):
            messagebox.showerror("Thiếu tệp", "Vui lòng chọn tệp đề bài / rubric!")
            return

        self.btn_start.configure(state="disabled")
        self.btn_open_word.configure(state="disabled")
        threading.Thread(target=self._run_pipeline_task, args=(paper_path, brief_path), daemon=True).start()

    def _run_pipeline_task(self, paper_path: str, brief_path: str):
        try:
            self._log("\n" + "="*50)
            self._log(f"[1/4] Khởi động luồng xử lý: {os.path.basename(paper_path)}")

            # Lấy thông tin preset môn
            sel_course = self.combo_course.get().split(" — ")[0].strip()
            preset = self.courses.get(sel_course, list(self.courses.values())[0])

            mode = self.mode_var.get()
            base_name = os.path.splitext(os.path.basename(paper_path))[0]
            out_dir = os.path.join(self.config.gdrive_data_dir, "outputs")
            os.makedirs(out_dir, exist_ok=True)

            eval_rep = None
            if mode == "auto":
                self._log("[2/4] Đang đối chiếu rubric và thẩm định sơ bộ...")
                doc_orig = Document(paper_path)
                doc_brief = Document(brief_path)
                p_text = "\n".join(p.text for p in doc_orig.paragraphs if p.text.strip())
                b_text = "\n".join(p.text for p in doc_brief.paragraphs if p.text.strip())
                eval_rep = self.provider.evaluate_paper(p_text, b_text, preset)
                mode = eval_rep.triage_mode
                self._log(f" -> Đánh giá điểm: {eval_rep.overall_score_estimate}")
                self._log(f" -> Tự động chọn chế độ: {'SỬA NHẸ' if mode == 'light' else 'SỬA NẶNG'}")

            if mode == "light":
                out_path = os.path.join(out_dir, f"{base_name}_fixed_trackchanges.docx")
                self._log("[3/4] Đang chạy luồng Sửa nhẹ (Kiểm tra hình thức + Sửa câu từ + Track Changes + Comments)...")
                res = self.light_pipeline.run(
                    paper_path,
                    brief_path,
                    preset,
                    out_path,
                    max_chunks_to_fix=10,
                    eval_report=eval_rep,
                    log_cb=self._log
                )
                self._log(f"[4/4] Hoàn tất! Phát hiện {res['rule_check']['issues_count']} lỗi hình thức, tạo {res['total_edits']} điểm sửa.")
            else:
                out_path = os.path.join(out_dir, f"{base_name}_rewritten_augmented.docx")
                self._log("[3/4] Đang chạy luồng Sửa nặng (Viết lại + Bổ sung tiêu chí thiếu + Highlight vàng)...")
                res = self.heavy_pipeline.run(paper_path, brief_path, preset, out_path)
                self._log(f"[4/4] Hoàn tất! Viết mới {res['augmented_sections_count']} phần highlight và {res['verify_citations_count']} nguồn cần kiểm chứng.")

            self.current_output_path = out_path
            self._log(f"[Xong] File kết quả: {out_path}")
            self.btn_open_word.configure(state="normal")
            messagebox.showinfo("Thành công", f"Bài tập đã được sửa xong!\nĐường dẫn: {out_path}")

        except Exception as e:
            self._log(f"[LỖI] Có lỗi xảy ra trong quá trình xử lý: {e}")
            messagebox.showerror("Lỗi xử lý", str(e))
        finally:
            self.btn_start.configure(state="normal")

    def _open_in_word(self):
        if self.current_output_path and os.path.exists(self.current_output_path):
            try:
                os.startfile(self.current_output_path)
            except Exception as e:
                messagebox.showerror("Lỗi mở Word", f"Không thể mở file: {e}")

    # ---------------- VIEW 2: DẠY TOOL ----------------
    def _create_learn_view(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self)
        frame.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(frame, text="Dạy Tool Học Phong Cách (Style Learning)", font=ctk.CTkFont(size=22, weight="bold"))
        title.grid(row=0, column=0, sticky="w", pady=(0, 15))

        card_pair = ctk.CTkFrame(frame)
        card_pair.grid(row=1, column=0, sticky="ew", pady=(0, 15), padx=5)
        card_pair.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card_pair, text="Bài gốc (chưa sửa):", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=15, pady=10, sticky="w")
        self.entry_learn_orig = ctk.CTkEntry(card_pair, placeholder_text="Chọn tệp bài gốc...")
        self.entry_learn_orig.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(card_pair, text="Chọn tệp...", width=90, command=self._browse_learn_orig).grid(row=0, column=2, padx=15, pady=10)

        ctk.CTkLabel(card_pair, text="Bản bạn đã sửa:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, padx=15, pady=10, sticky="w")
        self.entry_learn_fixed = ctk.CTkEntry(card_pair, placeholder_text="Chọn tệp bản bạn đã sửa hoàn thiện...")
        self.entry_learn_fixed.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(card_pair, text="Chọn tệp...", width=90, command=self._browse_learn_fixed).grid(row=1, column=2, padx=15, pady=10)

        self.btn_start_learn = ctk.CTkButton(
            frame, text="BÓC TÁCH & HỌC QUY TẮC SỬA",
            font=ctk.CTkFont(size=14, weight="bold"), height=38,
            command=self._start_learning_thread
        )
        self.btn_start_learn.grid(row=2, column=0, sticky="ew", pady=(5, 15), padx=5)

        # Style Guide Editor
        ctk.CTkLabel(frame, text="Living Style Guide (Quy chuẩn học thuật đang áp dụng):", font=ctk.CTkFont(size=15, weight="bold")).grid(row=3, column=0, sticky="w", pady=(5, 5))
        self.txt_style_guide = ctk.CTkTextbox(frame, height=260, font=ctk.CTkFont(family="Consolas", size=12))
        self.txt_style_guide.grid(row=4, column=0, sticky="nsew", padx=5, pady=(0, 10))
        self.txt_style_guide.insert("end", self.learner.get_style_guide())

        btn_save_style = ctk.CTkButton(frame, text="Lưu Cập Nhật Style Guide", width=200, command=self._save_style_guide)
        btn_save_style.grid(row=5, column=0, sticky="e", padx=5)

        return frame

    def _browse_learn_orig(self):
        f = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if f:
            self.entry_learn_orig.delete(0, "end")
            self.entry_learn_orig.insert(0, f)

    def _browse_learn_fixed(self):
        f = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if f:
            self.entry_learn_fixed.delete(0, "end")
            self.entry_learn_fixed.insert(0, f)

    def _start_learning_thread(self):
        orig_p = self.entry_learn_orig.get().strip()
        fixed_p = self.entry_learn_fixed.get().strip()
        if not os.path.exists(orig_p) or not os.path.exists(fixed_p):
            messagebox.showerror("Thiếu tệp", "Vui lòng chọn cả bài gốc và bài đã sửa!")
            return
        
        self.btn_start_learn.configure(state="disabled")
        threading.Thread(target=self._run_learning_task, args=(orig_p, fixed_p), daemon=True).start()

    def _run_learning_task(self, orig_p: str, fixed_p: str):
        try:
            pair_id = f"pair_user_{len(os.listdir(self.learner.corrections_dir)) + 1}"
            records = self.learner.learn_from_pair(orig_p, fixed_p, course_code="GENERAL", pair_id=pair_id, max_analyze=10)
            messagebox.showinfo("Hoàn tất", f"Đã học xong {len(records)} mẫu sửa lỗi!\nDữ liệu đã được nạp vào kho {pair_id}.jsonl.")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
        finally:
            self.btn_start_learn.configure(state="normal")

    def _save_style_guide(self):
        content = self.txt_style_guide.get("1.0", "end-1c")
        try:
            with open(self.learner.style_guide_path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Thành công", "Đã lưu Style Guide lên bộ nhớ thành công!")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    # ---------------- VIEW 3: CÀI ĐẶT ----------------
    def _create_settings_view(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self)
        frame.grid_columnconfigure(1, weight=1)

        title = ctk.CTkLabel(frame, text="Cài Đặt Hệ Thống", font=ctk.CTkFont(size=22, weight="bold"))
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))

        # Provider Selection
        ctk.CTkLabel(frame, text="Nhà cung cấp AI (Provider):", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, padx=15, pady=10, sticky="w")
        current_p_name = "Google AI Pro (Gemini CLI)" if self.config.default_provider == "gemini_cli" else "KiraAI (API Key)"
        self.combo_provider = ctk.CTkOptionMenu(
            frame,
            values=["Google AI Pro (Gemini CLI)", "KiraAI (API Key)"],
            width=300
        )
        self.combo_provider.set(current_p_name)
        self.combo_provider.grid(row=1, column=1, padx=15, pady=10, sticky="w")

        # Gemini CLI Path
        ctk.CTkLabel(frame, text="Lệnh Gemini CLI (Windows):", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, padx=15, pady=8, sticky="w")
        self.entry_gemini_cli = ctk.CTkEntry(frame, width=300)
        self.entry_gemini_cli.grid(row=2, column=1, padx=15, pady=8, sticky="w")
        self.entry_gemini_cli.insert(0, self.config.gemini_cli_path or "gemini.cmd")

        # Gemini API Key (Google AI Studio)
        ctk.CTkLabel(frame, text="Gemini API Key (Google AI):", font=ctk.CTkFont(weight="bold")).grid(row=3, column=0, padx=15, pady=8, sticky="w")
        self.entry_gemini_api_key = ctk.CTkEntry(frame, width=450, placeholder_text="Dán key từ aistudio.google.com/apikey (nếu dùng)")
        self.entry_gemini_api_key.grid(row=3, column=1, padx=15, pady=8, sticky="w")
        self.entry_gemini_api_key.insert(0, self.config.gemini_api_key or "")

        # Kira API Key
        ctk.CTkLabel(frame, text="KiraAI API Key:", font=ctk.CTkFont(weight="bold")).grid(row=4, column=0, padx=15, pady=8, sticky="w")
        self.entry_api_key = ctk.CTkEntry(frame, width=450)
        self.entry_api_key.grid(row=4, column=1, padx=15, pady=8, sticky="w")
        self.entry_api_key.insert(0, self.config.kira_api_key)

        # Base URL
        ctk.CTkLabel(frame, text="API Base URL:", font=ctk.CTkFont(weight="bold")).grid(row=5, column=0, padx=15, pady=8, sticky="w")
        self.entry_base_url = ctk.CTkEntry(frame, width=450)
        self.entry_base_url.grid(row=5, column=1, padx=15, pady=8, sticky="w")
        self.entry_base_url.insert(0, self.config.kira_base_url)

        # Model
        ctk.CTkLabel(frame, text="Model AI:", font=ctk.CTkFont(weight="bold")).grid(row=6, column=0, padx=15, pady=8, sticky="w")
        self.entry_model = ctk.CTkEntry(frame, width=250)
        self.entry_model.grid(row=6, column=1, padx=15, pady=8, sticky="w")
        self.entry_model.insert(0, self.config.kira_model)

        # GDrive Data dir
        ctk.CTkLabel(frame, text="Thư mục Data / GDrive:", font=ctk.CTkFont(weight="bold")).grid(row=7, column=0, padx=15, pady=8, sticky="w")
        self.entry_gdrive = ctk.CTkEntry(frame, width=450)
        self.entry_gdrive.grid(row=7, column=1, padx=15, pady=8, sticky="w")
        self.entry_gdrive.insert(0, self.config.gdrive_data_dir)

        # Save Button
        btn_save_config = ctk.CTkButton(
            frame, text="Lưu Cấu Hình", height=38, width=150,
            command=self._save_settings
        )
        btn_save_config.grid(row=8, column=1, sticky="w", padx=15, pady=18)

        return frame

    def _save_settings(self):
        p_val = self.combo_provider.get()
        if "Google AI Pro" in p_val:
            self.config.default_provider = "gemini_cli"
        else:
            self.config.default_provider = "kira_api"

        self.config.gemini_cli_path = self.entry_gemini_cli.get().strip() or "gemini.cmd"
        self.config.gemini_api_key = self.entry_gemini_api_key.get().strip()
        self.config.kira_api_key = self.entry_api_key.get().strip()
        self.config.kira_base_url = self.entry_base_url.get().strip()
        self.config.kira_model = self.entry_model.get().strip()
        self.config.gdrive_data_dir = self.entry_gdrive.get().strip()
        save_config(self.config)
        self._init_provider()
        messagebox.showinfo("Thành công", f"Đã lưu cài đặt và chuyển sang: {p_val}!")

def main():
    app = AssignmentFixerApp()
    app.mainloop()

if __name__ == "__main__":
    main()
