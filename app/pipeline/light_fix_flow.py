import os
import tempfile
from typing import Dict, Any, List, Optional
from docx import Document
from app.core.word_com import WordController
from app.core.rule_checker import AcademicRuleChecker
from app.core.learning_manager import LearningManager
from app.llm.base import LLMProviderInterface, EvaluationReport, ChunkEdit

class LightFixPipeline:
    """
    Luồng xử lý Sửa nhẹ (Light Fix Workflow):
    1. Kiểm tra quy chuẩn hình thức bằng RuleChecker (font, margin, citations).
    2. Đánh giá tổng thể bài viết bằng LLM (EvaluationReport).
    3. Cắt bài thành các cụm đoạn văn theo section, gọi LLM sửa từng đoạn dựa trên Style Guide.
    4. Tạo file Word sửa đổi tạm thời.
    5. Gọi WordController.compare_documents để sinh Track Changes nguyên bản.
    6. Chèn các Comments giải thích lý do sửa vào file kết quả.
    """
    def __init__(
        self,
        llm_provider: LLMProviderInterface,
        learning_manager: LearningManager
    ):
        self.llm = llm_provider
        self.learner = learning_manager

    def run(
        self,
        input_doc_path: str,
        brief_doc_path: str,
        preset: Dict[str, Any],
        output_doc_path: str,
        max_chunks_to_fix: int = 10,
        eval_report: Optional[Any] = None,
        log_cb: Optional[Any] = None
    ) -> Dict[str, Any]:
        def _log(msg: str):
            if log_cb:
                try:
                    log_cb(msg)
                except Exception:
                    pass

        result = {
            "input_file": input_doc_path,
            "output_file": output_doc_path,
            "rule_check": {},
            "evaluation": None,
            "total_edits": 0,
            "comments_added": 0
        }

        # 1. Kiểm tra hình thức (Rule-based)
        _log("  -> [3.1] Đang quét kiểm tra định dạng & trích dẫn tham khảo...")
        checker = AcademicRuleChecker(preset)
        result["rule_check"] = checker.check_document(input_doc_path)

        # 2. Đọc nội dung bài và brief
        doc_orig = Document(input_doc_path)
        doc_brief = Document(brief_doc_path)

        paper_text = "\n".join(p.text for p in doc_orig.paragraphs if p.text.strip())
        brief_text = "\n".join(p.text for p in doc_brief.paragraphs if p.text.strip())

        # 3. Đánh giá tổng quan đối chiếu Rubric (nếu chưa có)
        if eval_report is None:
            _log("  -> [3.2] Đang thẩm định đối chiếu Rubric BTEC...")
            eval_report = self.llm.evaluate_paper(paper_text, brief_text, preset)
        result["evaluation"] = eval_report.model_dump() if hasattr(eval_report, "model_dump") else eval_report

        # 4. Sửa từng đoạn văn học thuật
        style_guide = self.learner.get_style_guide()
        
        # Gom các đoạn văn cần sửa (bỏ qua các đoạn quá ngắn hoặc tiêu đề)
        paragraphs_to_fix = []
        for idx, p in enumerate(doc_orig.paragraphs):
            t = p.text.strip()
            if len(t) > 60 and not p.style.name.lower().startswith("heading"):
                paragraphs_to_fix.append((idx, t))

        selected_paras = paragraphs_to_fix[:max_chunks_to_fix]
        chunk_texts = [p[1] for p in selected_paras]

        _log(f"  -> [3.3] Đang dùng AI sửa và nâng chuẩn học thuật cho {len(chunk_texts)} đoạn văn...")
        edits = self.llm.fix_chunks(
            chunks=chunk_texts,
            section_context="Body Paragraphs",
            style_guide=style_guide,
            preset=preset
        )
        result["total_edits"] = len(edits)

        # 5. Tạo file Word đã sửa tạm thời
        temp_dir = tempfile.mkdtemp()
        temp_revised_path = os.path.join(temp_dir, "revised_temp.docx")

        # Sao chép và thay thế các đoạn đã sửa
        doc_revised = Document(input_doc_path)
        comments_payload = []

        for edit, (orig_idx, _) in zip(edits, selected_paras):
            if edit.revised_text and edit.revised_text != edit.original_text:
                doc_revised.paragraphs[orig_idx].text = edit.revised_text
                comments_payload.append({
                    "target_text": edit.revised_text[:100],
                    "comment_text": f"[{edit.category.upper()}] {edit.rationale}"
                })

        doc_revised.save(temp_revised_path)

        # 6. Dùng Word COM so sánh tạo Track Changes và chèn Comments
        os.makedirs(os.path.dirname(output_doc_path), exist_ok=True)
        _log("  -> [3.4] Đang gọi Microsoft Word tạo Track Changes bản địa...")
        with WordController(visible=False) as word:
            # So sánh tạo track changes
            word.compare_documents(
                original_path=input_doc_path,
                revised_path=temp_revised_path,
                output_path=output_doc_path,
                author_name="AssignmentFixer"
            )
            # Chèn nhận xét
            if comments_payload:
                _log(f"  -> [3.5] Đang chèn {len(comments_payload)} nhận xét (Comments) vào Word...")
                word.add_comments(output_doc_path, comments_payload)
                result["comments_added"] = len(comments_payload)

        # Dọn dẹp tệp tạm
        try:
            os.remove(temp_revised_path)
            os.rmdir(temp_dir)
        except Exception:
            pass

        return result
