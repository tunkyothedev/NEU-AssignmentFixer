import json
import re
import requests
from typing import List, Dict, Any, Optional
from app.llm.base import LLMProviderInterface, EvaluationReport, ChunkEdit

class OpenAICompatibleProvider(LLMProviderInterface):
    """
    Tích hợp các cổng API chuẩn OpenAI (như KiraAI, OpenRouter, DeepSeek, OpenAI).
    Hỗ trợ endpoint https://kiraai.vn/api/v1 với API Key trả phí theo token/gói.
    """
    def __init__(
        self,
        api_key: str,
        base_url: str = "https://kiraai.vn/api/v1",
        model: str = "gemini-2.5-pro",
        timeout_sec: int = 180
    ):
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.timeout_sec = timeout_sec

    def _chat_completion(self, messages: List[Dict[str, str]], temperature: float = 0.2) -> str:
        if not self.api_key:
            raise ValueError("Chưa thiết lập API Key. Vui lòng cấu hình API Key trong local_config.json hoặc giao diện Cài đặt.")

        url = f"{self.base_url}/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }

        # Danh sách model dự phòng thông minh khi gặp 503 (High demand) hoặc 404 (Not found)
        models_to_try = [self.model]
        if "generativelanguage.googleapis.com" in self.base_url:
            for alt in ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-3.1-flash-lite", "gemini-3.8-flash"]:
                if alt not in models_to_try:
                    models_to_try.append(alt)

        last_error = ""
        import time
        for m in models_to_try:
            payload = {
                "model": m,
                "messages": messages,
                "temperature": temperature
            }

            # Thử lại nếu gặp 503 (quá tải tạm thời) hoặc 429 (nghẽn tần suất)
            for attempt in range(2):
                try:
                    response = requests.post(url, headers=headers, json=payload, timeout=self.timeout_sec)
                    if response.status_code == 200:
                        data = response.json()
                        self.model = m  # Ghi nhớ model đang hoạt động mượt mà
                        try:
                            return data["choices"][0]["message"]["content"]
                        except (KeyError, IndexError):
                            raise RuntimeError(f"Unexpected response format from API: {data}")

                    # Nếu gặp lỗi 503 quá tải, chờ 2.5s rồi thử lại
                    if response.status_code in (503, 429) and attempt == 0:
                        time.sleep(2.5)
                        continue

                    last_error = f"API Error ({response.status_code}): {response.text}"
                    if response.status_code in (503, 404):
                        # Chuyển sang model tiếp theo trong danh sách dự phòng
                        break
                    else:
                        raise RuntimeError(last_error)
                except requests.exceptions.RequestException as e:
                    last_error = str(e)
                    break

        raise RuntimeError(last_error or "Không thể kết nối đến máy chủ AI.")

    def evaluate_paper(
        self,
        paper_text: str,
        brief_text: str,
        preset: Dict[str, Any]
    ) -> EvaluationReport:
        criteria_str = json.dumps(preset.get("assessment_criteria", []), indent=2, ensure_ascii=False)
        prompt = f"""
You are an expert academic assessor at the International School of Management and Economics (ISME) - National Economics University (NEU).
Evaluate the following student submission against the official BTEC Assignment Brief & Assessment Criteria.

Course: {preset.get('course_name')} ({preset.get('course_code')})
Assessment Criteria (Pass, Merit, Distinction):
{criteria_str}

=== ASSIGNMENT BRIEF ===
{brief_text[:3000]}

=== STUDENT PAPER ===
{paper_text[:12000]}

Respond ONLY with a valid JSON object matching this schema:
{{
  "course_code": "{preset.get('course_code')}",
  "overall_score_estimate": "Pass / Merit / Distinction / Refer",
  "triage_mode": "light" or "heavy",
  "summary": "2-3 sentences overview in English",
  "key_strengths": ["list of strengths"],
  "missing_or_weak_elements": ["list of missing or weak LO/criteria"],
  "criteria_feedback": [
    {{
      "id": "e.g. P1",
      "status": "Achieved / Partially Achieved / Missing",
      "comments": "specific academic feedback in English"
    }}
  ]
}}
"""
        messages = [
            {"role": "system", "content": "You are a professional academic reviewer. Respond only with valid JSON."},
            {"role": "user", "content": prompt}
        ]
        raw = self._chat_completion(messages, temperature=0.1)
        cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw.strip(), flags=re.MULTILINE)
        data = json.loads(cleaned)
        return EvaluationReport(**data)

    def fix_chunks(
        self,
        chunks: List[str],
        section_context: str,
        style_guide: str,
        preset: Dict[str, Any]
    ) -> List[ChunkEdit]:
        chunks_json = json.dumps([{"id": i, "text": c} for i, c in enumerate(chunks)], ensure_ascii=False)
        prompt = f"""
You are an elite academic proofreader and business editor for NEU ISME BTEC assignments.
Refine and enhance the following paragraphs for the section: "{section_context}".

Formatting & Citation rules:
- Citation Style: {preset.get('formatting', {}).get('citation_style', 'HARVARD')}
- Maintain formal, rigorous academic tone (avoid colloquial phrasing, emotional prose, or PR language).

Style Guide rules:
{style_guide}

Original paragraphs:
{chunks_json}

Return ONLY a JSON array of edited paragraphs matching this schema:
[
  {{
    "chunk_id": 0,
    "original_text": "original paragraph text",
    "revised_text": "enhanced, academically rigorous paragraph",
    "category": "grammar" or "academic_tone" or "logic" or "citation",
    "rationale": "Clear, concise academic explanation of the correction"
  }}
]
"""
        messages = [
            {"role": "system", "content": "You are a professional academic editor. Respond only with valid JSON array."},
            {"role": "user", "content": prompt}
        ]
        raw = self._chat_completion(messages, temperature=0.2)
        cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw.strip(), flags=re.MULTILINE)
        data = json.loads(cleaned)
        return [ChunkEdit(**item) for item in data]

    def rewrite_and_augment(
        self,
        paper_text: str,
        brief_text: str,
        preset: Dict[str, Any],
        missing_points: List[str]
    ) -> str:
        prompt = f"""
You are an expert academic writer and strategic management scholar at NEU.
The following student assignment is incomplete and requires a comprehensive rewrite and academic expansion.

Course: {preset.get('course_name')} ({preset.get('course_code')})
Target Program: BTEC Higher National Diploma in Business (RQF)

Missing or weak areas that MUST be fully developed:
{json.dumps(missing_points, indent=2)}

=== BRIEF SUMMARY ===
{brief_text[:3000]}

=== ORIGINAL PAPER ===
{paper_text}

TASK:
1. Re-structure the paper logically with standard numbered headings (1., 1.1, 1.2, 2., 2.1...).
2. Preserve valid student arguments and case data, but completely elevate the prose to rigorous academic business English.
3. Write thorough, fully-elaborated analyses for all missing learning outcomes and criteria.
4. For any newly augmented sections or paragraphs, enclose them in tags <AUGMENTED> ... </AUGMENTED> so they can be highlighted in Word.
5. For any newly introduced citations, tag them with [VERIFY] (e.g., '[VERIFY] (Kotter & Heskett, 1992)').

Generate the complete, publication-ready academic paper.
"""
        messages = [
            {"role": "system", "content": "You are an elite academic writer. Output formal, high-standard academic English text."},
            {"role": "user", "content": prompt}
        ]
        return self._chat_completion(messages, temperature=0.3)
