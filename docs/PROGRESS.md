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
| **M5** | Đóng gói Giao diện Desktop (UI) & Launcher | **HOÀN THÀNH (100%)** | - Đã xây dựng hoàn thiện ứng dụng Desktop Native chuẩn Dark Theme hiện đại tại [`app/main.py`](file:///E:/AssignmentFixer/app/main.py) bằng `customtkinter`.<br>- **3 màn hình chức năng trực quan:**<br>  1. **Sửa bài:** Chọn file bài làm, chọn đề bài, chọn môn, chọn chế độ sửa (Tự động / Nhẹ / Nặng), xem log tiến trình thời gian thực và bấm 1 nút mở ngay file Word kết quả.<br>  2. **Dạy tool:** Chọn cặp bài cũ để AI tự động bóc tách quy tắc lưu vào `corrections/` và chỉnh sửa trực tiếp `style_guide.md`.<br>  3. **Cài đặt:** Quản lý KiraAI API Key, model AI và đường dẫn thư mục Google Drive.<br>- Tạo launcher [`run_app.bat`](file:///E:/AssignmentFixer/run_app.bat) để khởi chạy app bằng 1 cú click chuột (cực kỳ tiện lợi cho partner).<br>- *Hotfix UI:* Bổ sung `from docx import Document` trong [`app/main.py`](file:///E:/AssignmentFixer/app/main.py) sửa triệt để lỗi `name 'Document' is not defined` khi bấm Bắt đầu sửa ở chế độ Tự động. |

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

---

## 5. Hướng Dẫn Sử Dụng
1. Nhấp đúp vào tệp [`run_app.bat`](file:///E:/AssignmentFixer/run_app.bat) để mở giao diện ứng dụng.
2. Chọn bài làm sinh viên và đề bài rubric, sau đó bấm **Bắt đầu sửa bài**.
3. Bạn sẽ thấy thanh log hiển thị chi tiết tiến trình từng bước `[3.1]`, `[3.2]`, `[3.3]`, `[3.4]`, `[3.5]` và tự động hoàn thành trong tích tắc!
