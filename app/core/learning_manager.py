import os
import json
from typing import List, Dict, Any, Optional
from app.core.diff_extractor import DiffExtractor
from app.llm.base import LLMProviderInterface

class LearningManager:
    """
    Quản lý vòng lặp học hỏi phong cách sửa bài (Continual Learning Loop):
    1. Bóc tách các cặp [Đoạn gốc -> Đoạn sửa] từ cặp tài liệu.
    2. Dùng LLM để phân loại lý do sửa (category & rationale).
    3. Lưu trữ vào kho data/corrections/<pair_id>.jsonl.
    4. Tự động đề xuất cập nhật quy tắc vào data/style_guide.md.
    """
    def __init__(self, data_dir: str, llm_provider: LLMProviderInterface):
        self.data_dir = data_dir
        self.corrections_dir = os.path.join(data_dir, "corrections")
        self.style_guide_path = os.path.join(data_dir, "style_guide.md")
        self.llm = llm_provider
        self.extractor = DiffExtractor()

        os.makedirs(self.corrections_dir, exist_ok=True)
        self._ensure_style_guide()

    def _ensure_style_guide(self):
        if not os.path.exists(self.style_guide_path):
            initial_content = """# Living Academic Style Guide — AssignmentFixer

## 1. Academic Vocabulary & Tone
- Avoid colloquial words: replace "a lot of" with "substantial/considerable", "very big" with "unprecedented/extensive".
- Maintain passive voice or objective statements: Avoid "I think", "In my opinion".
- Use precise management terminology appropriate to the unit.

## 2. Structure & Cohesion
- Every analytical section must follow Point - Evidence - Explanation - Link (PEEL).
- Ensure explicit citations for empirical claims and industry data.

## 3. Rubric Compliance
- Align subheadings with specific assessment criteria (e.g. LO1, LO2, P1, M1, D1).
- When expanding incomplete sections, cite valid academic frameworks (Kotter, Schein, Porter, Barney VRIO).
"""
            with open(self.style_guide_path, "w", encoding="utf-8") as f:
                f.write(initial_content)

    def learn_from_pair(
        self,
        orig_doc_path: str,
        fixed_doc_path: str,
        course_code: str,
        pair_id: str,
        max_analyze: int = 15
    ) -> List[Dict[str, Any]]:
        diffs = self.extractor.compare_documents(orig_doc_path, fixed_doc_path, course_code, pair_id)
        
        # Lọc các đoạn sửa đổi có giá trị học hỏi cao (modifications)
        mods = [d for d in diffs if d["change_type"] == "modification" and d["similarity_ratio"] < 0.95]
        selected = mods[:max_analyze]

        enriched_records = []
        for item in selected:
            # Nhờ LLM phân loại nguyên nhân sửa
            prompt = f"""
Analyze this academic editing correction made by an expert assessor for an NEU BTEC assignment:
Section: {item['section_context']}
Original text: "{item['original_text']}"
Revised text: "{item['revised_text']}"

Classify the correction into one category: ["academic_tone", "grammar", "critical_thinking", "citation", "clarity"]
And provide a 1-sentence rationale explaining WHY this change was necessary.

Return ONLY a JSON object:
{{
  "category": "...",
  "rationale": "..."
}}
"""
            try:
                raw = self.llm._chat_completion([{"role": "user", "content": prompt}], temperature=0.1)
                cleaned = raw.strip().replace("```json", "").replace("```", "").strip()
                meta = json.loads(cleaned)
                item["category"] = meta.get("category", "academic_tone")
                item["rationale"] = meta.get("rationale", "Enhanced academic expression and rigour.")
            except Exception:
                item["category"] = "academic_tone"
                item["rationale"] = "Refined tone and academic vocabulary."

            enriched_records.append(item)

        # Lưu vào file JSONL
        out_file = os.path.join(self.corrections_dir, f"{pair_id}.jsonl")
        with open(out_file, "w", encoding="utf-8") as f:
            for rec in enriched_records:
                f.write(json.dumps(rec, ensure_ascii=False) + "\n")

        return enriched_records

    def get_style_guide(self) -> str:
        if os.path.exists(self.style_guide_path):
            with open(self.style_guide_path, "r", encoding="utf-8") as f:
                return f.read()
        return ""
