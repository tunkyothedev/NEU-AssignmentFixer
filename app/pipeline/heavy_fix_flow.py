import os
import re
from typing import Dict, Any, List, Optional
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from app.core.rule_checker import AcademicRuleChecker
from app.core.learning_manager import LearningManager
from app.llm.base import LLMProviderInterface, EvaluationReport

class HeavyFixPipeline:
    """
    Luồng xử lý Sửa nặng (Heavy Rewrite & Augmentation Pipeline - M4):
    1. Kiểm tra hình thức (RuleChecker) để lấy thông số font, margin, citation.
    2. Gọi LLM đánh giá toàn diện đối chiếu với Brief & Rubric (EvaluationReport).
    3. Trích xuất các Learning Outcomes / Criteria bị thiếu hoặc yếu.
    4. Gọi LLM thực hiện Rewrite & Augmentation:
       - Tái cấu trúc bài viết khoa học theo chuẩn BTEC (1., 1.1, 1.2, 2., 2.1...).
       - Giữ nguyên dữ liệu của sinh viên, nâng cấp văn phong học thuật.
       - Viết bổ sung đầy đủ các tiêu chí còn thiếu.
       - Gắn thẻ <AUGMENTED>...</AUGMENTED> cho các đoạn viết mới.
       - Gắn tiền tố [VERIFY] cho các nguồn tham khảo sinh mới.
    5. Tạo tài liệu Word .docx mới chuẩn hóa theo preset.yaml:
       - Thiết lập margins 1 inch (2.54 cm).
       - Thiết lập font family (Times New Roman), cỡ chữ (12pt body, 14pt h1, 13pt h2...).
       - Dãn dòng 1.5 line, spacing after 6pt.
       - Tự động bóc tách các thẻ <AUGMENTED> và tô màu highlight vàng nhạt (YELLOW).
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
        output_doc_path: str
    ) -> Dict[str, Any]:
        result = {
            "input_file": input_doc_path,
            "output_file": output_doc_path,
            "rule_check": {},
            "evaluation": None,
            "augmented_sections_count": 0,
            "verify_citations_count": 0
        }

        # 1. Kiểm tra hình thức ban đầu
        checker = AcademicRuleChecker(preset)
        result["rule_check"] = checker.check_document(input_doc_path)

        # 2. Đọc toàn văn bài gốc và brief
        doc_orig = Document(input_doc_path)
        doc_brief = Document(brief_doc_path)

        orig_text = "\n".join(p.text for p in doc_orig.paragraphs if p.text.strip())
        brief_text = "\n".join(p.text for p in doc_brief.paragraphs if p.text.strip())

        # 3. Đánh giá toàn diện với LLM
        eval_report = self.llm.evaluate_paper(orig_text, brief_text, preset)
        result["evaluation"] = eval_report.model_dump()

        missing_points = eval_report.missing_or_weak_elements
        if not missing_points:
            missing_points = ["Enhance strategic recommendations, analytical rigor, and provide complete alignment with BTEC Merit/Distinction criteria."]

        # 4. Gọi LLM viết lại và bổ sung nội dung
        rewritten_markdown = self.llm.rewrite_and_augment(
            paper_text=orig_text,
            brief_text=brief_text,
            preset=preset,
            missing_points=missing_points
        )

        # 5. Tạo file Word .docx mới và format chuẩn theo preset
        self._build_formatted_docx(rewritten_markdown, preset, output_doc_path)

        # 6. Đếm số lượng phần augmented và verify citations
        result["augmented_sections_count"] = len(re.findall(r'<AUGMENTED>', rewritten_markdown, re.IGNORECASE))
        result["verify_citations_count"] = len(re.findall(r'\[VERIFY\]', rewritten_markdown))

        return result

    def _build_formatted_docx(self, text_content: str, preset: Dict[str, Any], output_path: str):
        doc = Document()
        fmt = preset.get("formatting", {})

        font_name = fmt.get("font_family", "Times New Roman")
        body_size = Pt(fmt.get("font_size_body", 12))
        h1_size = Pt(fmt.get("font_size_h1", 14))
        h2_size = Pt(fmt.get("font_size_h2", 13))
        h3_size = Pt(fmt.get("font_size_h3", 12))
        line_spacing = fmt.get("line_spacing", 1.5)
        space_after = Pt(fmt.get("space_after", 6))

        # Thiết lập lề 1 inch
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        # Phân tích từng dòng từ kết quả LLM
        lines = text_content.split("\n")
        in_augmented_block = False

        for raw_line in lines:
            line = raw_line.strip()
            if not line:
                continue

            # Kiểm tra thẻ mở/đóng <AUGMENTED> ở cấp dòng
            if "<AUGMENTED>" in line and "</AUGMENTED>" in line:
                line = line.replace("<AUGMENTED>", "").replace("</AUGMENTED>", "").strip()
                is_line_augmented = True
            elif "<AUGMENTED>" in line:
                in_augmented_block = True
                line = line.replace("<AUGMENTED>", "").strip()
                is_line_augmented = True
            elif "</AUGMENTED>" in line:
                in_augmented_block = False
                line = line.replace("</AUGMENTED>", "").strip()
                is_line_augmented = True
            else:
                is_line_augmented = in_augmented_block

            if not line:
                continue

            # Nhận diện Heading
            is_h1 = bool(re.match(r'^(?:#\s+|\d+\.\s+[A-Z])', line))
            is_h2 = bool(re.match(r'^(?:##\s+|\d+\.\d+\s+[A-Z])', line))
            is_h3 = bool(re.match(r'^(?:###\s+|\d+\.\d+\.\d+\s+[A-Z])', line))

            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = line_spacing
            p.paragraph_format.space_after = space_after

            # Loại bỏ ký tự markdown header nếu có
            clean_line = re.sub(r'^#+\s*', '', line)

            # Xử lý runs và highlight
            run = p.add_run(clean_line)
            run.font.name = font_name

            if is_h1:
                run.font.size = h1_size
                run.font.bold = True
                p.paragraph_format.space_before = Pt(12)
            elif is_h2:
                run.font.size = h2_size
                run.font.bold = True
                p.paragraph_format.space_before = Pt(8)
            elif is_h3:
                run.font.size = h3_size
                run.font.bold = True
            else:
                run.font.size = body_size

            # Highlight màu vàng nếu là phần AI viết mới
            if is_line_augmented:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)
