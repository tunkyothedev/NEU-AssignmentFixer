import subprocess
import json
import os
import re
import tempfile
from typing import List, Dict, Any
from app.llm.base import LLMProviderInterface, EvaluationReport, ChunkEdit

class GeminiCLIProvider(LLMProviderInterface):
    """
    Tích hợp Gemini CLI ở chế độ headless (-p ... -o json).
    Sử dụng tài khoản Google AI Pro (1.500 requests/ngày).
    """
    def __init__(self, cli_cmd: str = "gemini.cmd", model: str = "gemini-3.1-flash-lite", api_key: str = "", timeout_sec: int = 180):
        if os.name == "nt" and cli_cmd == "gemini":
            cli_cmd = "gemini.cmd"
        self.cli_cmd = cli_cmd
        self.model = model
        self.api_key = api_key
        self.timeout_sec = timeout_sec

    def _call_cli(self, prompt: str) -> str:
        import shutil
        cli_available = (
            shutil.which(self.cli_cmd) or
            (os.name == "nt" and (
                shutil.which(f"{self.cli_cmd}.cmd") or
                shutil.which(f"{self.cli_cmd}.exe") or
                shutil.which("gemini.cmd") or
                shutil.which("gemini")
            ))
        )
        if not cli_available:
            raise RuntimeError(
                f"Không tìm thấy công cụ Gemini CLI ('{self.cli_cmd}') trên máy tính này.\n\n"
                "Máy tính của bạn chưa cài đặt Node.js / Gemini CLI.\n"
                "Giải pháp dễ nhất:\n"
                "1. Vào tab 'Cài Đặt Hệ Thống' -> Dán Gemini API Key (lấy miễn phí tại: https://aistudio.google.com/apikey)\n"
                "2. Hoặc chuyển Nhà cung cấp AI sang 'KiraAI (API Key)'\n"
                "3. Hoặc nếu muốn dùng CLI: Cài Node.js rồi mở CMD chạy: npm install -g @google/gemini-cli"
            )

        # Chạy trong thư mục cô lập để tránh CLI đọc mã nguồn
        try:
            from app.config import BASE_DIR
            sandbox_dir = str(BASE_DIR / "data" / "temp")
        except Exception:
            sandbox_dir = os.path.join(tempfile.gettempdir(), "AssignmentFixer_temp")
        os.makedirs(sandbox_dir, exist_ok=True)

        cmd = [
            self.cli_cmd,
            "-m", self.model,
            "-o", "json"
        ]

        env = os.environ.copy()
        if self.api_key:
            env["GEMINI_API_KEY"] = self.api_key

        result = subprocess.run(
            cmd,
            input=prompt,
            cwd=sandbox_dir,
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=self.timeout_sec,
            shell=True
        )

        if result.returncode != 0:
            error_msg = result.stderr or result.stdout
            raise RuntimeError(f"Gemini CLI Error (code {result.returncode}): {error_msg}")

        # Parse JSON output từ CLI (bỏ qua warning text ở đầu stdout)
        try:
            raw_out = result.stdout.strip()
            first_brace = raw_out.find('{')
            last_brace = raw_out.rfind('}')
            if first_brace != -1 and last_brace != -1:
                json_str = raw_out[first_brace:last_brace+1]
                data = json.loads(json_str)
                if "response" in data:
                    return data["response"]
                elif "error" in data:
                    raise RuntimeError(f"Gemini API returned error: {data['error']}")
            return result.stdout
        except json.JSONDecodeError:
            return result.stdout

    def _chat_completion(self, messages: List[Dict[str, str]], temperature: float = 0.2) -> str:
        prompt = "\n\n".join(f"{m.get('role', 'user').upper()}:\n{m.get('content', '')}" for m in messages)
        return self._call_cli(prompt)

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
        raw = self._call_cli(prompt)
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
        raw = self._call_cli(prompt)
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
        return self._call_cli(prompt)
