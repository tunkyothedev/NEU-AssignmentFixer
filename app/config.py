import os
import sys
import json
import socket
from pathlib import Path
from pydantic import BaseModel, Field

def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent

BASE_DIR = get_base_dir()
CONFIG_FILE = BASE_DIR / "local_config.json"

class AppConfig(BaseModel):
    machine_id: str = Field(default_factory=lambda: socket.gethostname())
    gdrive_data_dir: str = str(BASE_DIR / "data")
    
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
    check_update_on_startup: bool = True

def load_config() -> AppConfig:
    cfg = None
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            cfg = AppConfig(**data)
        except Exception:
            pass
    if cfg is None:
        cfg = AppConfig()

    # Tự động thẩm định và hiệu chỉnh gdrive_data_dir để tránh lỗi ổ đĩa khác máy (vd: E:\ sang máy không có E:\)
    default_data_dir = str(BASE_DIR / "data")
    if not cfg.gdrive_data_dir:
        cfg.gdrive_data_dir = default_data_dir
    else:
        try:
            p = Path(cfg.gdrive_data_dir)
            if not p.is_absolute():
                cfg.gdrive_data_dir = str((BASE_DIR / p).resolve())
            elif not p.exists():
                cfg.gdrive_data_dir = default_data_dir
        except Exception:
            cfg.gdrive_data_dir = default_data_dir

    save_config(cfg)
    return cfg

def save_config(config: AppConfig) -> None:
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config.model_dump(), f, indent=2, ensure_ascii=False)
