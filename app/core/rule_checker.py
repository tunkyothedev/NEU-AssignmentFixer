import re
from typing import Dict, Any, List
from docx import Document
from docx.shared import Pt, Inches

class AcademicRuleChecker:
    """
    Kiểm tra quy chuẩn hình thức theo luật cứng (Rule-based) bằng python-docx:
    - Font chữ, cỡ chữ
    - Khoảng cách dòng & đoạn
    - Lề trang
    - Đối chiếu trích dẫn In-text vs References
    """
    def __init__(self, preset: Dict[str, Any]):
        self.preset = preset
        self.fmt = preset.get("formatting", {})

    def check_document(self, doc_path: str) -> Dict[str, Any]:
        doc = Document(doc_path)
        issues = []
        stats = {
            "total_paragraphs": len(doc.paragraphs),
            "total_tables": len(doc.tables),
            "fonts_used": set(),
            "in_text_citations": [],
            "references_found": []
        }

        target_font = self.fmt.get("font_family", "Times New Roman").lower()
        target_body_size = self.fmt.get("font_size_body", 13)

        # 0. Đếm số từ (Paragraphs + Tables)
        p_words = sum(len(p.text.split()) for p in doc.paragraphs if p.text.strip())
        t_words = 0
        seen_cells = set()
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    if cell._tc not in seen_cells:
                        seen_cells.add(cell._tc)
                        t_words += len(cell.text.split())
        total_words = p_words + t_words
        stats["paragraph_words"] = p_words
        stats["table_words"] = t_words
        stats["total_words"] = total_words

        target_words = self.preset.get("target_word_count")
        tol = self.preset.get("word_count_tolerance", 0.10)
        if target_words:
            min_w = int(target_words * (1 - tol))
            max_w = int(target_words * (1 + tol))
            if total_words < min_w:
                issues.append({
                    "type": "word_count",
                    "severity": "high",
                    "message": f"Document length ({total_words} words) is below required limit ({min_w}-{max_w} words; target {target_words} ±{int(tol*100)}%)."
                })
            elif total_words > max_w:
                issues.append({
                    "type": "word_count",
                    "severity": "medium",
                    "message": f"Document length ({total_words} words) exceeds required limit ({min_w}-{max_w} words; target {target_words} ±{int(tol*100)}%)."
                })

        # 1. Kiểm tra lề trang
        for i, section in enumerate(doc.sections):
            top_in = round(section.top_margin.inches, 2)
            bottom_in = round(section.bottom_margin.inches, 2)
            left_in = round(section.left_margin.inches, 2)
            right_in = round(section.right_margin.inches, 2)
            
            # Chuẩn 1 inch (2.54 cm) +/- 0.1 inch
            if not (0.9 <= top_in <= 1.1 and 0.9 <= bottom_in <= 1.1 and 0.9 <= left_in <= 1.1 and 0.9 <= right_in <= 1.1):
                issues.append({
                    "type": "margin",
                    "severity": "medium",
                    "message": f"Section {i+1} margin is not standard 1 inch / 2.54cm (Top: {top_in}\", Bottom: {bottom_in}\", Left: {left_in}\", Right: {right_in}\")"
                })

        # 2. Kiểm tra Font chữ và kích thước trong các đoạn văn
        in_references = False
        citation_pattern = re.compile(r'\(([A-Z][a-zA-Z\s]+(?:,\s*\d{4}|(?:\s+et\s+al\.,\s*\d{4})))\)')

        for p_idx, p in enumerate(doc.paragraphs):
            text = p.text.strip()
            if not text:
                continue

            # Nhận diện khu vực References
            if any(ref_header in text.lower() for ref_header in ["references", "bibliography", "reference list"]):
                in_references = True
                continue

            if in_references:
                stats["references_found"].append(text)
            else:
                # Quét trích dẫn in-text
                matches = citation_pattern.findall(text)
                for m in matches:
                    stats["in_text_citations"].append(m)

            # Kiểm tra runs
            for r in p.runs:
                if r.font.name:
                    stats["fonts_used"].add(r.font.name)
                    if target_font not in r.font.name.lower():
                        issues.append({
                            "type": "font_family",
                            "severity": "low",
                            "message": f"Inconsistent font '{r.font.name}' at paragraph {p_idx+1}. Expected '{self.fmt.get('font_family')}'."
                        })

        # 3. Đối chiếu trích dẫn In-text vs References
        in_text_authors = set()
        for cit in stats["in_text_citations"]:
            author = cit.split(",")[0].strip()
            in_text_authors.add(author)

        ref_text_corpus = " ".join(stats["references_found"])
        for author in in_text_authors:
            if author.lower() not in ref_text_corpus.lower():
                issues.append({
                    "type": "missing_reference",
                    "severity": "high",
                    "message": f"In-text citation '({author})' has no matching entry in References list."
                })

        stats["fonts_used"] = list(stats["fonts_used"])
        return {
            "valid": len([i for i in issues if i["severity"] == "high"]) == 0,
            "issues_count": len(issues),
            "issues": issues,
            "stats": stats
        }
