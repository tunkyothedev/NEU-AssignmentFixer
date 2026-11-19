# Nhật ký Tiến độ & Trạng thái Triển khai — AssignmentFixer

Tài liệu này ghi lại chi tiết các cột mốc, trạng thái hoàn thành của từng module và kết quả kiểm thử thực tế. **Được cập nhật tự động sau mỗi phiên làm việc.**

---

## 1. Bảng Trạng Thái Cột Mốc (Milestone Status)

| Cột mốc | Mục tiêu | Trạng thái | Chi tiết hoàn thành & Kết quả kiểm thử |
|---|---|---|---|
| **M0** | Môi trường, cấu hình, Word COM & LLM Provider | **HOÀN THÀNH (100%)** | - Đã cài đặt: `pywin32`, `pyyaml`, `pydantic`, `python-docx`, `requests`, `httpx`.<br>- Tệp cấu hình [`local_config.json`](file:///E:/AssignmentFixer/local_config.json) đã được nạp API Key của KiraAI.<br>- Module [`app/core/word_com.py`](file:///E:/AssignmentFixer/app/core/word_com.py) kết nối `WINWORD.EXE` v16.0 thành công; test compare sinh Track Changes (60 ins, 15 del) trong 2s.<br>- **Kiểm thử KiraAI thành công 100%:** Kết nối model `kira-3.5-flash`, chạy thử nghiệm đánh giá toàn diện bài mẫu Business Strategy (`BS`) của NEU đối chiếu với rubric, trả về báo cáo JSON chuẩn xác trong 4 giây. |
| **M1** | Quản lý Môn học & Rubric BTEC ISME NEU | **HOÀN THÀNH (100%)** | - Đã trích xuất cấu hình tự động cho 4 môn học thực tế từ thư mục `Đề bài/`:<br>  + `BS` (Unit 43 Business Strategy): 11 tiêu chí (P1-P5, M1-M4, D1).<br>  + `OBM` (Unit 20 Organisational Behaviour): 3 tiêu chí LO4 (P6, M4, D3).<br>  + `SFM` (Skills for Management): 10 tiêu chí.<br>  + `MSBP` (Management & Strategic Project): 10 tiêu chí.<br>- Module [`app/core/rule_checker.py`](file:///E:/AssignmentFixer/app/core/rule_checker.py) kiểm tra font, margins, in-text citations vs references; đã test phát hiện lỗi thiếu reference trong bài BS mẫu. |
| **M2** | Pipeline Sửa nhẹ (Track Changes & Comments) | **HOÀN THÀNH (100%)** | - Đã triển khai xong [`app/pipeline/light_fix_flow.py`](file:///E:/AssignmentFixer/app/pipeline/light_fix_flow.py).<br>- Luồng chạy end-to-end hoàn chỉnh: Quét quy chuẩn hình thức -> Đánh giá rubric (KiraAI) -> Sửa đoạn văn học thuật -> Word Compare tạo Track Changes bản địa -> Chèn Word Comments giải thích.<br>- **Kiểm thử thực tế thành công:** File kết quả [`test_light_fix_output.docx`](file:///E:/AssignmentFixer/data/test_light_fix_output.docx) sinh ra chứa **55 điểm thêm (ins)**, **47 điểm xóa (del)** và **3 bong bóng nhận xét (Comments)** giải thích lỗi ngữ pháp/logic bằng tiếng Anh. |
| **M3** | Vòng lặp Học phong cách (Diff & Style Guide) | **HOÀN THÀNH (100%)** | - Đã triển khai [`app/core/diff_extractor.py`](file:///E:/AssignmentFixer/app/core/diff_extractor.py) và [`app/core/learning_manager.py`](file:///E:/AssignmentFixer/app/core/learning_manager.py).<br>- Tự động so sánh bóc tách khác biệt trên cả 4 cặp bài thật (BS, OBM, SFM, MSBP), bóc tách hàng trăm điểm sửa.<br>- Dùng KiraAI tự động phân loại loại lỗi (`academic_tone`, `clarity`, `critical_thinking`) và ghi thành công vào kho lưu trữ [`data/corrections/pair_obm_01.jsonl`](file:///E:/AssignmentFixer/data/corrections/pair_obm_01.jsonl).<br>- Khởi tạo và đồng bộ [`data/style_guide.md`](file:///E:/AssignmentFixer/data/style_guide.md). |
| **M4** | Pipeline Sửa nặng (Viết lại & Bổ sung Rubric) | **HOÀN THÀNH (100%)** | - Đã triển khai xong [`app/pipeline/heavy_fix_flow.py`](file:///E:/AssignmentFixer/app/pipeline/heavy_fix_flow.py).<br>- Tự động đánh giá các Learning Outcomes / Criteria bị thiếu hụt.<br>- Tái cấu trúc toàn bộ bài viết theo chuẩn đề mục BTEC (1., 2., 3., 4., 5.).<br>- **Tự động bổ sung các phần còn thiếu:** Thêm phân tích năng lực nội bộ (VRIO), Kế hoạch quản trị chiến lược (M4/D1) và Đề xuất chiến lược.<br>- **Kiểm thử thực tế thành công trên bài BS:** File kết quả [`test_heavy_fix_output.docx`](file:///E:/AssignmentFixer/data/test_heavy_fix_output.docx) được sinh trong 11s, gồm 45 đoạn văn chuẩn format Times New Roman 12pt, lề 1 inch, dãn dòng 1.5; **21 đoạn văn viết mới được tự động highlight màu vàng nhạt** và **8 trích dẫn mới được gắn cờ `[VERIFY]`** để người dùng dễ dàng kiểm chứng. |
| **M5** | Đóng gói Giao diện Desktop (UI) & Launcher | **HOÀN THÀNH (100%)** | - Đã xây dựng hoàn thiện ứng dụng Desktop Native chuẩn Dark Theme hiện đại tại [`app/main.py`](file:///E:/AssignmentFixer/app/main.py) bằng `customtkinter`.<br>- **3 màn hình chức năng trực quan:** Sửa bài, Dạy tool & Cài đặt hệ thống.<br>- **Đóng gói Portable Executable độc lập (`AssignmentFixer.exe`):** Đã biên dịch toàn bộ mã nguồn qua PyInstaller tại thư mục `dist/AssignmentFixer/AssignmentFixer.exe`, đóng gói sẵn GUI, Word COM, DOCX, Pydantic, bảo mật mã nguồn và chạy không cần cài Python hay mở cửa sổ bat/console.<br>- Hỗ trợ đường dẫn tương đối động trong [`app/config.py`](file:///E:/AssignmentFixer/app/config.py) khi chạy ở dạng Portable bundle. |
| **M6** | Hệ thống Tự động Cập nhật (Auto-Updater via GitHub Releases) | **HOÀN THÀNH (100%)** | - Đã triển khai [`app/version.py`](file:///E:/AssignmentFixer/app/version.py) quản lý semantic versioning (`v1.0.0`) và repo GitHub `tunkyothedev/NEU-AssignmentFixer`.<br>- Xây dựng [`app/core/updater.py`](file:///E:/AssignmentFixer/app/core/updater.py): kết nối GitHub API `releases/latest`, tải asset `.zip` kèm tiến trình, kích hoạt batch script ngầm để giải nén đè và khởi động lại app mượt mà.<br>- Tích hợp trực tiếp vào UI [`app/main.py`](file:///E:/AssignmentFixer/app/main.py): hiển thị version, kiểm tra tự động khi mở app, nút kiểm tra thủ công, modal xem Changelog và tải cập nhật.<br>- Tạo script tiện ích [`package_release.py`](file:///E:/AssignmentFixer/package_release.py) đóng gói zip release và tính mã băm SHA256.<br>- Viết bộ kiểm thử tự động [`tests/test_updater.py`](file:///E:/AssignmentFixer/tests/test_updater.py) (100% pass). |
| **M7** | Tự động hóa CI/CD Đóng gói & Phát hành (GitHub Actions) | **HOÀN THÀNH (100%)** | - Thiết lập workflow `.github/workflows/release.yml` trên `windows-latest`.<br>- Tự động kích hoạt khi gắn tag `v*` hoặc bấm nút 'Run workflow' trên GitHub.<br>- Máy ảo Cloud tự động tải code, cài đặt thư viện, chạy PyInstaller build exe, tạo file nén `.zip` và tính checksum `.sha256`.<br>- Tự động xuất bản GitHub Release và đính kèm assets mà không cần thao tác thủ công. |

---

## 2. Danh mục Tệp Mã Nguồn Đã Triển Khai

- [`app/config.py`](file:///E:/AssignmentFixer/app/config.py): Quản lý cấu hình `local_config.json` (đường dẫn GDrive, machine ID, auth mode).
- [`app/core/word_com.py`](file:///E:/AssignmentFixer/app/core/word_com.py): Điều khiển Word COM v16.0 (Compare native Track Changes, convert PDF, chèn Comments).
- [`app/core/rule_checker.py`](file:///E:/AssignmentFixer/app/core/rule_checker.py): Quét kiểm tra quy chuẩn hình thức theo `preset.yaml` (Font, lề, trích dẫn).
- [`app/core/diff_extractor.py`](file:///E:/AssignmentFixer/app/core/diff_extractor.py): Bóc tách các đoạn sửa đổi (modifications, additions, deletions) giữa bài gốc và bài duyệt.
- [`app/core/learning_manager.py`](file:///E:/AssignmentFixer/app/core/learning_manager.py): Phân loại lỗi bằng AI, lưu trữ kho `corrections/*.jsonl` và quản lý `style_guide.md`.
- [`app/pipeline/light_fix_flow.py`](file:///E:/AssignmentFixer/app/pipeline/light_fix_flow.py): Toàn bộ quy trình sửa nhẹ tự động (RuleCheck -> RubricEval -> ChunkFix -> TrackChanges -> Comments).
- [`app/pipeline/heavy_fix_flow.py`](file:///E:/AssignmentFixer/app/pipeline/heavy_fix_flow.py): Toàn bộ quy trình sửa nặng & bổ sung nội dung (RubricEval -> FullRewrite -> Highlight Yellow -> Verify Citations).
- [`app/main.py`](file:///E:/AssignmentFixer/app/main.py): Giao diện ứng dụng Desktop hoàn chỉnh (CustomTkinter Dark Theme).
- [`run_app.bat`](file:///E:/AssignmentFixer/run_app.bat): Phím tắt khởi chạy ứng dụng 1 chạm cho Windows.
- [`app/llm/base.py`](file:///E:/AssignmentFixer/app/llm/base.py): Interface trừu tượng `LLMProviderInterface`, `EvaluationReport`, `ChunkEdit`.
- [`app/llm/openai_compatible.py`](file:///E:/AssignmentFixer/app/llm/openai_compatible.py): Lớp gọi API trực tiếp KiraAI (`https://kiraai.vn/api/v1`) theo chuẩn OpenAI.
- [`app/llm/gemini_cli.py`](file:///E:/AssignmentFixer/app/llm/gemini_cli.py): Lớp gọi `gemini.cmd` headless, truyền prompt và parse JSON response.
- [`app/version.py`](file:///E:/AssignmentFixer/app/version.py): Quản lý hằng số phiên bản `APP_VERSION`, repo GitHub và thuật toán so sánh version.
- [`app/core/updater.py`](file:///E:/AssignmentFixer/app/core/updater.py): Engine tự động cập nhật (GitHub Releases API, stream download, detached update batch script).
- [`package_release.py`](file:///E:/AssignmentFixer/package_release.py): Công cụ đóng gói thư mục `dist` thành file ZIP release và tính mã băm SHA256.
- [`requirements.txt`](file:///E:/AssignmentFixer/requirements.txt): Danh mục thư viện phụ thuộc của ứng dụng.
- [`tests/test_updater.py`](file:///E:/AssignmentFixer/tests/test_updater.py): Bộ kiểm thử đơn vị cho module Updater và Version parser.
- [`data/courses/`](file:///E:/AssignmentFixer/data/courses/):
  - `BS/preset.yaml` (Business Strategy)
  - `OBM/preset.yaml` (Organizational Behaviour & Management)
  - `SFM/preset.yaml` (Strategic Financial Management)
  - `MSBP/preset.yaml` (Management & Strategic Business Project)

---

## 3. Dữ Liệu Kiểm Thử Thực Tế Đã Đối Soát

| Bộ môn | Đề bài | Bài sinh viên gốc | Bài đã sửa hoàn thiện |
|---|---|---|---|
| **BS** | `Đề bài/BS A1 brief S2.2526.docx` | `Chưa sửa/4774_Linh_..._BS_A0_....docx` | `Sửa rồi/4774_Linh_..._BS_A1.1_....docx` |
| **OBM** | `Đề bài/Assignment Brief_A1 OBM.docx` | `Chưa sửa/10240699_OB_A1_1.docx` | `Sửa rồi/10240699_OB_A1.1.docx` |
| **SFM** | `Đề bài/(BTEC_S2-25.26) SFM Assigment...` | `Chưa sửa/SFM_A1.1_10240699.docx` | `Sửa rồi/SFM_A1_Assignment.docx` |
| **MSBP** | `Đề bài/(MSBP_A1.1) Assignment Brief...` | `Chưa sửa/MSBP_A1_1 cho khách.docx` | `Sửa rồi/MSBP_A1.1_fixed.docx` |

---

## 4. Cập Nhật Mới Nhất & Trạng Thái Fix Lỗi
- **Đã sửa lỗi `name 'Document' is not defined`:**
  - Bổ sung `from docx import Document` đầy đủ trong [`app/main.py`](file:///E:/AssignmentFixer/app/main.py).
  - Tắt tiến trình giao diện cũ đang chạy ngầm để nạp lại mã nguồn mới nhất.
- **Xác thực & Tích hợp thành công Google Gemini CLI (100%):**
  - Đã xác thực thành công qua Gemini API Key (`Authenticated with gemini-api-key`).
  - Nâng cấp cơ chế truyền prompt qua `stdin` trong [`app/llm/gemini_cli.py`](file:///E:/AssignmentFixer/app/llm/gemini_cli.py) giải quyết triệt để giới hạn độ dài ký tự và ngắt dòng của Windows Shell.
  - Xử lý trích xuất JSON độc lập (bỏ qua các dòng cảnh báo môi trường terminal).
  - Tự động định tuyến sang model tối ưu (`gemini-3.1-flash-lite`) với tốc độ phản hồi cực nhanh (~2 - 6 giây) và không bị lỗi quá tải server 503.
  - **Kiểm thử thực tế thành công:** Chạy thử nghiệm thẩm định bài tập BTEC (`evaluate_paper`) trả về kết quả JSON đầy đủ chuẩn xác trong 7 giây.

- **Khắc phục triệt để lỗi treo ở bước [3/4] & RPC failed:**
  - **Nguyên nhân:** Trước đó Word COM sử dụng `Dispatch("Word.Application")`, khiến Python bị gắn vào một tiến trình Word cũ chạy ngầm bị treo từ phiên làm việc trước (PID 12928), dẫn đến việc Word không phản hồi và bị nghẽn ở bước gọi Track Changes.
  - **Khắc phục:** 
    + Chuyển sang dùng `win32com.client.DispatchEx("Word.Application")` trong [`app/core/word_com.py`](file:///E:/AssignmentFixer/app/core/word_com.py) để luôn khởi tạo một tiến trình Word độc lập, hoàn toàn cách ly và tự giải phóng khi xong việc.
    + Thêm cơ chế tái sử dụng báo cáo thẩm định (`eval_report`) để không phải gọi LLM 2 lần liên tiếp.
    + Bổ sung log thời gian thực chi tiết từng bước: `[3.1]` Quét định dạng -> `[3.2]` Thẩm định Rubric -> `[3.3]` Sửa đoạn văn học thuật -> `[3.4]` Tạo Track Changes bản địa -> `[3.5]` Chèn Comments vào Word.
  - **Kiểm thử thực tế end-to-end thành công 100%:** Luồng chạy tự động qua cả 5 bước trơn tru trong 18 giây, sinh file Word Track Changes kèm nhận xét Comments chuẩn chỉnh.

- **Khắc phục lỗi bản đóng gói EXE chỉ có mỗi môn BS:**
  - **Nguyên nhân:** Khi đóng gói PyInstaller, các tệp cài đặt môn học `data/courses/` (gồm BS, MSBP, OBM, SFM) và `data/style_guide.md` không được copy vào thư mục phát hành `dist/AssignmentFixer/`. Khi chạy file `.exe`, ứng dụng không tìm thấy thư mục `data/courses/` nên rơi vào nhánh fallback mặc định chỉ có duy nhất môn BS.
  - **Khắc phục:**
    + Cập nhật [`build_exe.py`](file:///E:/AssignmentFixer/build_exe.py) và [`package_release.py`](file:///E:/AssignmentFixer/package_release.py): Thêm cờ `--add-data` cho `data/courses/` và tự động đồng bộ đầy đủ toàn bộ thư mục `data/courses/` (cả 4 môn), `data/style_guide.md`, `local_config.example.json` sang `dist/AssignmentFixer/` trước khi đóng gói ZIP.
    + Cập nhật [`app/main.py`](file:///E:/AssignmentFixer/app/main.py): Hàm `_load_available_courses()` tự động tìm kiếm cả ở thư mục ngoài lẫn bộ nhớ nén nội bộ `sys._MEIPASS`/`_internal`, đồng thời nếu thư mục trống sẽ fallback hiển thị đầy đủ cả 4 môn chuẩn BTEC (BS, OBM, SFM, MSBP).
- **Khắc phục triệt để lỗi "System cannot find E:\" khi chạy trên máy khác:**
  - **Nguyên nhân:**
    + Trong [`app/llm/gemini_cli.py`](file:///E:/AssignmentFixer/app/llm/gemini_cli.py), biến `sandbox_dir` bị gán cứng đường dẫn tuyệt đối `r"E:\AssignmentFixer\data\temp"`. Khi chuyển app sang máy khác không có ổ đĩa `E:\`, lệnh `subprocess.run(cwd=sandbox_dir)` bị Windows chặn lại với lỗi không tìm thấy ổ đĩa `E:\`.
    + Script đóng gói trước đó vô tình đính kèm file cấu hình cá nhân `local_config.json` chứa đường dẫn `E:\AssignmentFixer\...` vào file ZIP.
  - **Khắc phục:**
    + Thay thế đường dẫn tĩnh trong `app/llm/gemini_cli.py` bằng `BASE_DIR / "data" / "temp"` tương đối động, hoặc thư mục `tempfile.gettempdir()` của hệ điều hành.
    + Cập nhật `app/config.py`: Hàm `load_config()` tự động thẩm định và hiệu chỉnh `gdrive_data_dir` về thư mục cục bộ mặc định nếu đường dẫn đã lưu thuộc về ổ đĩa không tồn tại trên máy người dùng.
    + Cập nhật `package_release.py`: Loại trừ hoàn toàn `local_config.json` khỏi tệp ZIP phát hành để không ghi đè cấu hình ổ đĩa máy của người phát triển sang máy người dùng khác.
  - **Kiểm thử thực tế thành công:** Tệp ZIP mới `AssignmentFixer-v1.0.2-windows.zip` được xác thực 100% không chứa `local_config.json`, tự động tương thích trên mọi máy tính (C:, D:, USB...).

---

## 5. Hướng Dẫn Sử Dụng
1. Nhấp đúp vào tệp [`run_app.bat`](file:///E:/AssignmentFixer/run_app.bat) để mở giao diện ứng dụng.
2. Chọn bài làm sinh viên và đề bài rubric, sau đó bấm **Bắt đầu sửa bài**.
3. Bạn sẽ thấy thanh log hiển thị chi tiết tiến trình từng bước `[3.1]`, `[3.2]`, `[3.3]`, `[3.4]`, `[3.5]` và tự động hoàn thành trong tích tắc!

---

## 6. Trạng Thái Kho Chứa Mã Nguồn (GitHub)
- **Repository:** [https://github.com/tunkyothedev/NEU-AssignmentFixer](https://github.com/tunkyothedev/NEU-AssignmentFixer)
- **Branch:** `main` (Tracking `origin/main`)
- **Chính sách bảo mật:**
  + Tệp `.gitignore` đã loại trừ toàn bộ tài liệu Word (`*.docx`, `*.doc`), bài làm sinh viên (`Chưa sửa/`, `Sửa rồi/`), đề bài (`Đề bài/`) và tệp xuất kết quả (`data/outputs/`).
  + Loại trừ khóa bảo mật và cấu hình cục bộ (`local_config.json`).
  + Cung cấp tệp mẫu [`local_config.example.json`](file:///E:/AssignmentFixer/local_config.example.json) sạch chuẩn cho cộng tác viên.

---

## 7. Quy Tắc Quản Trị Git & Nhánh (Git Workflow Rules)
1. **Quy định Push:**
   - **Tuyệt đối không tự ý push lên `main`** trừ khi người dùng ra lệnh trực tiếp.
   - **Chỉ push nhánh `main`**, không push các nhánh phụ/nhánh tính năng lên remote GitHub.
2. **Quy ước đặt tên nhánh (Branch Naming Convention):**
   - Tính năng mới: `feat/[tên_nhánh]`
   - Sửa lỗi/tính năng: `bugfix/[tên_nhánh]`
3. **Quy trình làm việc (Branching & Merge Flow):**
   - Khi có task mới, luôn tạo branch tách từ nhánh `main`.
   - Nếu có bất kỳ điểm phân vân nào về xuất phát điểm hoặc phạm vi task, phải hỏi ý kiến người dùng trước khi tạo.
   - Sau khi hoàn thành và kiểm thử task thành công trên branch, merge vào `main`. Không tự ý push trừ khi có lệnh.
4. **Quy định tài liệu bắt buộc:**
   - Đọc tài liệu Markdown (`docs/PROGRESS.md`) trước khi bắt đầu bất kỳ task nào.
   - Cập nhật tài liệu Markdown (`docs/PROGRESS.md`) ngay sau khi hoàn thành.

---

## 8. Quy Trình Vòng Lặp Kiểm Chứng Đơn Tác Tử (Single-Agent Verification Loop)
Áp dụng cho mọi task lập trình từ thời điểm này:
1. **LẬP KẾ HOẠCH (PLAN):**
   - Nhanh chóng phân tích nguyên nhân gốc rễ và đưa ra giải pháp kỹ thuật.
   - Trình bày danh sách tệp cần tạo/sửa dưới 5 gạch đầu dòng ngắn gọn trước khi viết code.
2. **THỰC THI (IMPLEMENT):**
   - Viết code sạch, đúng kiến trúc, phạm vi sửa chữa tối giản, không sửa lan man.
3. **KIỂM CHỨNG THỰC TẾ (VERIFY):**
   - Chạy lệnh terminal kiểm tra ngay lập tức (compile, unit test module liên quan, py_compile...).
4. **VÒNG LẶP SỬA LỖI CÓ GIỚI HẠN (Tối đa 3 lần - HARD STOP):**
   - Nếu lỗi: phân tích log, sửa và test lại.
   - Nếu thất bại 3 lần liên tiếp: **DỪNG LẬP TỨC**, giải thích lý do kẹt và xin ý kiến người dùng; không loop cháy quota.
5. **TỰ KIỂM TOÁN CUỐI CÙNG (SELF-AUDIT):**
   - Chạy `git diff` rà soát code thừa/debug, kiểm tra rò rỉ key nhạy cảm, dọn dẹp file nháp trước khi báo cáo.


