import os
import json
import socket
from pathlib import Path
from pydantic import BaseModel, Field

CONFIG_FILE = Path(r"E:\AssignmentFixer\local_config.json")

class AppConfig(BaseModel):
    machine_id: str = Field(default_factory=lambda: socket.gethostname())
    gdrive_data_dir: str = r"E:\AssignmentFixer\data"
    
    # Provider: 'kira_api' | 'gemini_cli' | 'gemini_api'
    default_provider: str = "kira_api"
    
    # KiraAI / OpenAI-Compatible settings
    kira_api_key: str = ""
    kira_base_url: str = "https://kiraai.vn/api/v1"
    kira_model: str = "gemini-2.5-pro"
    
    # Gemini CLI settings (Alternative)
    gemini_cli_path: str = "gemini"
    gemini_api_key: str = ""
    
    auto_open_word: bool = True

def load_config() -> AppConfig:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return AppConfig(**data)
        except Exception:
            pass
    cfg = AppConfig()
    save_config(cfg)
    return cfg

def save_config(config: AppConfig) -> None:
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config.model_dump(), f, indent=2, ensure_ascii=False)
