import os
import sys
import shutil
import subprocess
from pathlib import Path
import customtkinter

def build():
    ctk_dir = os.path.dirname(customtkinter.__file__)
    print(f"CustomTkinter located at: {ctk_dir}")

    base_dir = Path(__file__).resolve().parent

    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",
        "--windowed",
        "--name", "AssignmentFixer",
        "--add-data", f"{ctk_dir};customtkinter",
        "--add-data", f"{base_dir / 'data' / 'courses'};data/courses",
        "--add-data", f"{base_dir / 'data' / 'style_guide.md'};data",
        "--hidden-import", "docx",
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        "--hidden-import", "yaml",
        "--hidden-import", "pydantic",
        "--hidden-import", "requests",
        "--hidden-import", "app.version",
        "--hidden-import", "app.core.updater",
        "--hidden-import", "app.llm.gemini_cli",
        "--hidden-import", "app.llm.openai_compatible",
        "--collect-all", "customtkinter",
        os.path.join("app", "main.py")
    ]

    print("Running command:", " ".join(cmd))
    res = subprocess.run(cmd)
    if res.returncode != 0:
        print(f"Build failed with exit code: {res.returncode}")
        sys.exit(res.returncode)

    # Copy external data folder to dist so user can access and customize presets
    dist_dir = base_dir / "dist" / "AssignmentFixer"
    dist_data_courses = dist_dir / "data" / "courses"
    src_data_courses = base_dir / "data" / "courses"
    if src_data_courses.exists():
        print(f"[*] Copying {src_data_courses} -> {dist_data_courses}...")
        if dist_data_courses.exists():
            shutil.rmtree(dist_data_courses)
        shutil.copytree(src_data_courses, dist_data_courses)

    src_style_guide = base_dir / "data" / "style_guide.md"
    dist_style_guide = dist_dir / "data" / "style_guide.md"
    if src_style_guide.exists():
        dist_style_guide.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_style_guide, dist_style_guide)

    src_example_cfg = base_dir / "local_config.example.json"
    dist_example_cfg = dist_dir / "local_config.example.json"
    if src_example_cfg.exists():
        shutil.copy2(src_example_cfg, dist_example_cfg)

    print("Build completed successfully! Assets and presets bundled.")

if __name__ == "__main__":
    build()
