import re
import difflib
import uuid
from typing import List, Dict, Any, Tuple
from docx import Document

class DiffExtractor:
    """
    So sánh tài liệu gốc và tài liệu đã sửa để bóc tách các đoạn văn bản thay đổi.
    Hỗ trợ cả trường hợp sửa trực tiếp (đoạn tương ứng) và phân loại thay đổi:
    - Sửa câu từ / ngữ pháp (Modified)
    - Thêm mới đoạn / luận điểm (Added)
    - Xóa bỏ đoạn (Deleted)
    """

    @staticmethod
    def _clean_text(text: str) -> str:
        return re.sub(r'\s+', ' ', text).strip()

    def extract_paragraphs(self, doc_path: str) -> List[Dict[str, Any]]:
        doc = Document(doc_path)
        items = []
        current_heading = "Introduction / Overview"
        
        for idx, p in enumerate(doc.paragraphs):
            text = self._clean_text(p.text)
            if not text:
                continue

            style_name = p.style.name.lower()
            if "heading" in style_name or "title" in style_name or text.startswith(('I.', 'II.', 'III.', 'IV.', '1.', '2.', '3.', 'Part ')):
                current_heading = text[:100]

            items.append({
                "para_index": idx,
                "heading": current_heading,
                "text": text
            })
        return items

    def compare_documents(
        self,
        orig_doc_path: str,
        fixed_doc_path: str,
        course_code: str,
        pair_id: str
    ) -> List[Dict[str, Any]]:
        orig_items = self.extract_paragraphs(orig_doc_path)
        fixed_items = self.extract_paragraphs(fixed_doc_path)

        orig_texts = [item["text"] for item in orig_items]
        fixed_texts = [item["text"] for item in fixed_items]

        matcher = difflib.SequenceMatcher(None, orig_texts, fixed_texts)
        differences = []

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                continue
            
            elif tag == 'replace':
                # Đoạn văn bản bị sửa đổi
                for o_idx, f_idx in zip(range(i1, i2), range(j1, j2)):
                    differences.append({
                        "correction_id": f"corr_{uuid.uuid4().hex[:12]}",
                        "pair_id": pair_id,
                        "course_code": course_code,
                        "change_type": "modification",
                        "section_context": fixed_items[f_idx]["heading"],
                        "original_text": orig_texts[o_idx],
                        "revised_text": fixed_texts[f_idx],
                        "similarity_ratio": round(difflib.SequenceMatcher(None, orig_texts[o_idx], fixed_texts[f_idx]).ratio(), 3)
                    })
                # Nếu số lượng đoạn replace không bằng nhau
                if (i2 - i1) > (j2 - j1):
                    for o_idx in range(i1 + (j2 - j1), i2):
                        differences.append({
                            "correction_id": f"corr_{uuid.uuid4().hex[:12]}",
                            "pair_id": pair_id,
                            "course_code": course_code,
                            "change_type": "deletion",
                            "section_context": orig_items[o_idx]["heading"],
                            "original_text": orig_texts[o_idx],
                            "revised_text": "",
                            "similarity_ratio": 0.0
                        })
                elif (j2 - j1) > (i2 - i1):
                    for f_idx in range(j1 + (i2 - i1), j2):
                        differences.append({
                            "correction_id": f"corr_{uuid.uuid4().hex[:12]}",
                            "pair_id": pair_id,
                            "course_code": course_code,
                            "change_type": "addition",
                            "section_context": fixed_items[f_idx]["heading"],
                            "original_text": "",
                            "revised_text": fixed_texts[f_idx],
                            "similarity_ratio": 0.0
                        })

            elif tag == 'delete':
                for o_idx in range(i1, i2):
                    differences.append({
                        "correction_id": f"corr_{uuid.uuid4().hex[:12]}",
                        "pair_id": pair_id,
                        "course_code": course_code,
                        "change_type": "deletion",
                        "section_context": orig_items[o_idx]["heading"],
                        "original_text": orig_texts[o_idx],
                        "revised_text": "",
                        "similarity_ratio": 0.0
                    })

            elif tag == 'insert':
                for f_idx in range(j1, j2):
                    differences.append({
                        "correction_id": f"corr_{uuid.uuid4().hex[:12]}",
                        "pair_id": pair_id,
                        "course_code": course_code,
                        "change_type": "addition",
                        "section_context": fixed_items[f_idx]["heading"],
                        "original_text": "",
                        "revised_text": fixed_texts[f_idx],
                        "similarity_ratio": 0.0
                    })

        return differences
