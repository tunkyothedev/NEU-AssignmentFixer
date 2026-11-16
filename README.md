# AssignmentFixer — Intelligent Academic Assignment Editor

**AssignmentFixer** là công cụ desktop chạy trên Windows hỗ trợ sửa và nâng cấp bài tập lớn (assignments) bằng tiếng Anh của sinh viên Đại học Kinh Tế Quốc Dân (NEU) theo chuẩn rubric môn học và quy cách học thuật quốc tế.

Hệ thống kết hợp kiểm tra quy chuẩn hình thức theo luật cứng (rule-based) với mô hình ngôn ngữ lớn (Gemini) để sửa lỗi diễn đạt, ngữ pháp học thuật, lập luận và bổ sung nội dung thiếu hụt theo rubric. Công cụ có khả năng **học dần phong cách sửa bài** của người dùng qua kho lưu trữ cặp bài đối chiếu và hướng dẫn phong cách (Style Guide) đồng bộ qua Google Drive giữa nhiều máy tính.

---

## Mục lục tài liệu

- [Kiến trúc hệ thống chi tiết](file:///E:/AssignmentFixer/docs/ARCHITECTURE.md) (`docs/ARCHITECTURE.md`)
- [Đặc tả chức năng & định dạng dữ liệu](file:///E:/AssignmentFixer/docs/SPECIFICATION.md) (`docs/SPECIFICATION.md`)
- [Hướng dẫn thiết lập & sử dụng](file:///E:/AssignmentFixer/docs/USER_GUIDE.md) (`docs/USER_GUIDE.md`)
- [Nhật ký tiến độ & trạng thái](file:///E:/AssignmentFixer/docs/PROGRESS.md) (`docs/PROGRESS.md`)

---

## 1. Đặc điểm nổi bật

1. **Sửa toàn diện cả Hình thức & Nội dung:**
   - **Hình thức:** Kiểm tra font chữ, cỡ chữ, khoảng cách dòng/đoạn, lề trang, heading hierarchy, tiêu đề bảng/biểu, định dạng trích dẫn in-text và tài liệu tham khảo (References) theo chuẩn APA 7 hoặc Harvard.
   - **Nội dung:** Đối chiếu cấu trúc bài làm với đề bài (brief) và thang chấm (rubric), chỉ ra các luận điểm yếu, thiếu số liệu chứng minh hoặc mục bị bỏ sót.
2. **Hai chế độ xử lý linh hoạt:**
   - **Chế độ sửa nhẹ (Track Changes):** Dành cho bài có nền tảng khá, chỉ sửa ngữ pháp, câu từ và bổ sung nhận xét chi tiết (Word Comments). Sử dụng cơ chế Microsoft Word Compare gốc để tạo bản Word sạch với lịch sử thay đổi chuẩn xác.
   - **Chế độ sửa nặng (Full Rewrite & Expansion):** Dành cho bài làm yếu hoặc thiếu mục lớn. Tạo file `.docx` mới chuẩn hóa định dạng, viết lại lập luận và **tự động bổ sung các phần còn thiếu theo rubric**. Các nội dung do AI bổ sung được **highlight trực quan**, trích dẫn bổ sung được đánh dấu `[VERIFY]` để người dùng đối soát.
3. **Cơ chế học tập tích lũy (Continual Learning Loop):**
   - Không cần fine-tune model ngay từ đầu khi tập dữ liệu còn nhỏ.
   - Khi người dùng nạp các cặp bài trước và sau khi duyệt (hoặc bài tự sửa thủ công), hệ thống tự bóc tách các điểm sửa đổi thành kho ví dụ mẫu (`corrections/`).
   - Tự động gợi ý cập nhật quy tắc sửa vào `style_guide.md` để người dùng duyệt.
   - Sử dụng mô hình Semantic Embedding chạy cục bộ trên CPU máy tính để tìm các ví dụ tương đồng truyền vào prompt của các bài sửa kế tiếp.
4. **Tối ưu chi phí và phần cứng:**
   - Hoạt động mượt mà trên cả laptop đồ họa rời (NVIDIA RTX) và laptop văn phòng CPU AMD không có GPU rời.
   - Backend chính tích hợp qua **Gemini CLI (Headless Mode)** tận dụng tài khoản **Google AI Pro** (hạn ngạch lên đến 1.500 requests/ngày, truy cập model Pro không mất phí token).
   - Có cơ chế Fallback sang **Gemini API Key (AI Studio)** hoặc các nhà cung cấp bên thứ 3 khác.
5. **Đồng bộ hóa cộng tác:**
   - Dữ liệu môn học, kho bài tập mẫu và Style Guide được lưu trữ dưới dạng các tập tin độc lập (YAML, JSONL, Markdown) trên thư mục Google Drive dùng chung, tránh xung đột khóa file database giữa hai người dùng.

---

## 2. Mô hình kiến trúc tổng quan

```
+---------------------------------------------------------------------------------+
|                                 GIAO DIỆN APP                                   |
|                (Desktop Native Webview / NiceGUI - Tiếng Việt)                  |
|     [Sửa bài mới]      [Thư viện môn & Rubric]     [Dạy tool]     [Cài đặt]     |
+---------------------------------------------------------------------------------+
                                       |
                                       v
+---------------------------------------------------------------------------------+
|                             CORE CONTROLLER & PIPELINE                          |
|  - Task Parser (DOCX / PDF to DOCX via Word COM)                                |
|  - Academic Rule-based Engine (Font, margins, citations, headings)              |
|  - Semantic Search Engine (Local ONNX CPU Embedding for Few-Shot Examples)       |
|  - LLM Orchestrator (Gemini CLI Headless Mode / Fallback API Key)               |
|  - Document Assembler & Word COM Engine (Compare, Track Changes, Highlight)     |
+---------------------------------------------------------------------------------+
                                       |
                   +-------------------+-------------------+
                   |                                       |
                   v                                       v
+-------------------------------------+   +-------------------------------------+
|        LOCAL MACHINE RESOURCES      |   |       SHARED GOOGLE DRIVE DATA      |
| - MS Word (WINWORD.EXE COM)         |   | - courses/<course_id>/preset.yaml   |
| - Local ONNX Embedding Model (CPU)  |   | - corrections/<id>.jsonl            |
| - Local SQLite Vector/FTS Index     |   | - pairs/<id>/ (original & final)    |
| - Local Settings & Auth Session     |   | - style_guide.md (Living guidelines)|
+-------------------------------------+   +-------------------------------------+
```

---

## 3. Cấu trúc thư mục mã nguồn

```
E:\AssignmentFixer/
├── docs/                           # Tài liệu kỹ thuật chi tiết
│   ├── ARCHITECTURE.md             # Kiến trúc hệ thống & thiết kế module
│   ├── SPECIFICATION.md            # Đặc tả kỹ thuật, schema dữ liệu & rules
│   └── USER_GUIDE.md               # Hướng dẫn cài đặt, cấu hình & sử dụng
├── app/                            # Mã nguồn ứng dụng
│   ├── config.py                   # Quản lý cấu hình cục bộ và đường dẫn GDrive
│   ├── main.py                     # Entrypoint khởi chạy ứng dụng GUI
│   ├── core/                       # Lõi xử lý logic
│   │   ├── word_com.py             # Giao tiếp MS Word (Compare, PDF convert, Comment)
│   │   ├── docx_parser.py          # Bóc tách cấu trúc DOCX (run, paragraph, table)
│   │   ├── rule_checker.py         # Kiểm tra quy chuẩn hình thức theo preset
│   │   ├── diff_extractor.py       # Trích xuất cặp sai/đúng từ tài liệu sửa
│   │   └── embedding_search.py     # Local Semantic Embedding & truy vấn few-shot
│   ├── llm/                        # Tương tác mô hình AI
│   │   ├── base.py                 # Interface trừu tượng cho LLM Backend
│   │   ├── gemini_cli.py           # Backend Gemini CLI (Google AI Pro headless)
│   │   └── gemini_api.py           # Backend Gemini API Key fallback
│   ├── pipeline/                   # Các luồng xử lý chính
│   │   ├── evaluation_flow.py      # Đánh giá tổng thể bài viết theo rubric
│   │   ├── light_fix_flow.py       # Luồng sửa nhẹ (Track changes + comments)
│   │   ├── heavy_fix_flow.py       # Luồng sửa nặng (Viết lại + bổ sung nội dung)
│   │   └── learning_flow.py        # Luồng nạp cặp bài đối chiếu & học style
│   └── ui/                         # Giao diện người dùng
│       ├── views/                  # Các màn hình chức năng
│       └── components/             # Các widget giao diện tái sử dụng
├── tests/                          # Kiểm thử đơn vị & tích hợp
├── requirements.txt                # Danh sách thư viện Python
└── README.md                       # Tài liệu tổng quan dự án
```

---

## 4. Bắt đầu nhanh (Quickstart)

### Điều kiện tiên quyết
1. Hệ điều hành: **Windows 10 / 11**.
2. **Microsoft Office Word** đã được cài đặt và kích hoạt bản quyền.
3. **Python 3.12+**.
4. **Node.js (LTS)** để chạy Gemini CLI.
5. Tài khoản **Google AI Pro** (hoặc Gemini API Key).

Xem hướng dẫn chi tiết từng bước cài đặt và cấu hình tại [USER_GUIDE.md](file:///E:/AssignmentFixer/docs/USER_GUIDE.md).
