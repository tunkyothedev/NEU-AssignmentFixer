import os
import sys
import subprocess
import customtkinter

def build():
    ctk_dir = os.path.dirname(customtkinter.__file__)
    print(f"CustomTkinter located at: {ctk_dir}")

    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",
        "--windowed",
        "--name", "AssignmentFixer",
        "--add-data", f"{ctk_dir};customtkinter",
        "--hidden-import", "docx",
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        "--hidden-import", "yaml",
        "--hidden-import", "pydantic",
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
    print("Build completed successfully!")

if __name__ == "__main__":
    build()
