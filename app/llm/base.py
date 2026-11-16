from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional
from pydantic import BaseModel

class EvaluationReport(BaseModel):
    course_code: str
    overall_score_estimate: str
    triage_mode: str  # "light" | "heavy"
    summary: str
    criteria_feedback: List[Dict[str, Any]]
    key_strengths: List[str]
    missing_or_weak_elements: List[str]

class ChunkEdit(BaseModel):
    chunk_id: int
    original_text: str
    revised_text: str
    category: str  # "grammar" | "academic_tone" | "logic" | "citation"
    rationale: str

class LLMProviderInterface(ABC):
    @abstractmethod
    def evaluate_paper(
        self,
        paper_text: str,
        brief_text: str,
        preset: Dict[str, Any]
    ) -> EvaluationReport:
        pass

    @abstractmethod
    def fix_chunks(
        self,
        chunks: List[str],
        section_context: str,
        style_guide: str,
        preset: Dict[str, Any]
    ) -> List[ChunkEdit]:
        pass

    @abstractmethod
    def rewrite_and_augment(
        self,
        paper_text: str,
        brief_text: str,
        preset: Dict[str, Any],
        missing_points: List[str]
    ) -> str:
        pass
