import os
import sys
import shutil
import hashlib
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from app.version import APP_VERSION

def compute_sha256(filepath: Path) -> str:
    h = hashlib.sha256()
    with open(filepath, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()

def package():
    base_dir = Path(__file__).resolve().parent
    dist_dir = base_dir / "dist" / "AssignmentFixer"
    releases_dir = base_dir / "dist" / "releases"
    releases_dir.mkdir(parents=True, exist_ok=True)

    if not dist_dir.exists():
        print(f"[Error] Thư mục '{dist_dir}' không tồn tại. Vui lòng chạy 'python build_exe.py' trước!")
        sys.exit(1)

    zip_filename = f"AssignmentFixer-v{APP_VERSION}-windows.zip"
    zip_path = releases_dir / zip_filename

    print(f"[*] Đang đóng gói bản phát hành: {zip_filename}...")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(dist_dir):
            for file in files:
                full_path = Path(root) / file
                rel_path = full_path.relative_to(dist_dir)
                zf.write(full_path, arcname=str(rel_path))

    file_size_mb = zip_path.stat().st_size / (1024 * 1024)
    sha256 = compute_sha256(zip_path)

    print("\n" + "=" * 60)
    print("  ĐÓNG GÓI BẢN PHÁT HÀNH THÀNH CÔNG!")
    print("=" * 60)
    print(f"Phiên bản : v{APP_VERSION}")
    print(f"Tệp nén   : {zip_path}")
    print(f"Dung lượng: {file_size_mb:.2f} MB")
    print(f"SHA256    : {sha256}")
    print("\n[Hướng dẫn tạo GitHub Release]:")
    print(f"1. Vào: https://github.com/tunkyothedev/NEU-AssignmentFixer/releases/new")
    print(f"2. Tag version: v{APP_VERSION}")
    print(f"3. Release title: AssignmentFixer v{APP_VERSION}")
    print(f"4. Đính kèm tệp: {zip_filename} vào mục Assets")
    print("=" * 60 + "\n")

if __name__ == "__main__":
    package()
