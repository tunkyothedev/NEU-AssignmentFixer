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

    version = os.getenv("RELEASE_VERSION") or (sys.argv[1] if len(sys.argv) > 1 else APP_VERSION)
    version = version.strip().lstrip("vV")

    # Ensure presets and essentials exist in dist before zipping
    src_courses = base_dir / "data" / "courses"
    dist_courses = dist_dir / "data" / "courses"
    if src_courses.exists():
        print(f"[*] Đồng bộ dữ liệu môn học: {src_courses} -> {dist_courses}")
        if dist_courses.exists():
            shutil.rmtree(dist_courses)
        shutil.copytree(src_courses, dist_courses)

    src_style_guide = base_dir / "data" / "style_guide.md"
    dist_style_guide = dist_dir / "data" / "style_guide.md"
    if src_style_guide.exists():
        dist_style_guide.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_style_guide, dist_style_guide)

    src_example_cfg = base_dir / "local_config.example.json"
    dist_example_cfg = dist_dir / "local_config.example.json"
    if src_example_cfg.exists():
        shutil.copy2(src_example_cfg, dist_example_cfg)

    zip_filename = f"AssignmentFixer-v{version}-windows.zip"
    zip_path = releases_dir / zip_filename

    print(f"[*] Đang đóng gói bản phát hành: {zip_filename}...")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(dist_dir):
            for file in files:
                # Không bao giờ đóng gói file cấu hình cục bộ của máy build vào file zip phát hành
                if file.lower() == "local_config.json":
                    continue
                full_path = Path(root) / file
                rel_path = full_path.relative_to(dist_dir)
                zf.write(full_path, arcname=str(rel_path))

    file_size_mb = zip_path.stat().st_size / (1024 * 1024)
    sha256 = compute_sha256(zip_path)

    sha256_filename = f"{zip_filename}.sha256"
    sha256_path = releases_dir / sha256_filename
    with open(sha256_path, "w", encoding="utf-8") as f:
        f.write(f"{sha256}  {zip_filename}\n")

    print("\n" + "=" * 60)
    print("  ĐÓNG GÓI BẢN PHÁT HÀNH THÀNH CÔNG!")
    print("=" * 60)
    print(f"Phiên bản : v{version}")
    print(f"Tệp nén   : {zip_path}")
    print(f"Dung lượng: {file_size_mb:.2f} MB")
    print(f"SHA256    : {sha256}")
    print(f"Checksum  : {sha256_path}")
    print("\n[Hướng dẫn tạo GitHub Release]:")
    print(f"1. Vào: https://github.com/tunkyothedev/NEU-AssignmentFixer/releases/new")
    print(f"2. Tag version: v{version}")
    print(f"3. Release title: AssignmentFixer v{version}")
    print(f"4. Đính kèm tệp: {zip_filename} và {sha256_filename} vào mục Assets")
    print("=" * 60 + "\n")

if __name__ == "__main__":
    package()
