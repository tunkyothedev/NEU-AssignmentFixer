# Đặc tả Kỹ thuật & Chức năng — AssignmentFixer

Tài liệu này quy định chi tiết về các yêu cầu chức năng, yêu cầu phi chức năng, lược đồ cấu trúc dữ liệu (data schemas), quy chuẩn hình thức văn bản học thuật và đặc tả các màn hình giao diện của hệ thống **AssignmentFixer**.

---

## 1. Yêu cầu Chức năng (Functional Requirements)

### 1.1. Quản lý Môn học & Rubric (Course & Rubric Management)
- **FR-1.1:** Cho phép người dùng tạo môn học mới bằng cách nhập tên môn, mã môn (ví dụ: `ECO301`, `MGT202`, `IB305`).
- **FR-1.2:** Hỗ trợ nạp tệp Đề bài (`brief.docx` / `brief.pdf`) và Thang chấm (`rubric.docx` / `rubric.pdf`).
- **FR-1.3:** Tự động sử dụng mô hình ngôn ngữ để trích xuất các thông tin sau từ tệp rubric và tạo thành tệp `preset.yaml`:
  - Quy chuẩn font chữ (Times New Roman, Arial, Calibri...).
  - Cỡ chữ nội dung (12pt, 13pt) và cỡ chữ tiêu đề các cấp.
  - Quy chuẩn giãn dòng (1.15, 1.3, 1.5 line spacing) và dãn đoạn (after 6pt).
  - Quy cách lề trang (chuẩn 1 inch / 2.54 cm).
  - Phong cách trích dẫn học thuật (APA 7th Edition hoặc Harvard Citation Style).
  - Danh sách các tiêu chí chấm điểm (Criteria), mô tả mức điểm và trọng số (%) tương ứng.
- **FR-1.4:** Cung cấp giao diện trực quan cho phép người dùng xem trước, chỉnh sửa bằng tay và lưu lại cấu hình `preset.yaml`.

---

### 1.2. Phân tích & Kiểm tra Quy chuẩn Hình thức (Rule-Based Checking)
- **FR-2.1:** Quét toàn bộ văn bản đầu vào để phát hiện các lỗi định dạng mà không cần phụ thuộc vào AI:
  - Sai font hoặc font không đồng bộ xuyên suốt các đoạn văn.
  - Sai kích cỡ chữ (so sánh với preset của môn).
  - Lề trang bị lệch chuẩn.
  - Tiêu đề bảng biểu (Table caption) hoặc hình vẽ (Figure caption) đặt sai vị trí hoặc thiếu nguồn trích dẫn.
  - Thiếu trang bìa tiêu chuẩn hoặc thông tin sinh viên/mã đề.
- **FR-2.2:** Quét và đối chiếu trích dẫn học thuật:
  - Quét danh sách các trích dẫn trong văn bản (In-text citations).
  - Đối chiếu với danh sách tài liệu tham khảo (References) ở cuối bài.
  - Liệt kê danh sách các nguồn được trích dẫn trong bài nhưng thiếu trong danh mục References.
  - Liệt kê danh sách các nguồn có trong References nhưng không hề được dẫn chứng trong nội dung bài (nguồn ảo).

---

### 1.3. Đánh giá Tổng thể & Phân loại Mức độ Lỗi (Global Evaluation & Triage)
- **FR-3.1:** Phân tích nội dung toàn bộ bài làm của sinh viên đối chiếu với Đề bài và Rubric của môn học:
  - Kiểm tra tính đầy đủ của các mục theo yêu cầu đề bài.
  - Đánh giá chất lượng lập luận, tính chặt chẽ về logic học thuật và mức độ đáp ứng của từng tiêu chí trong rubric.
  - Chấm điểm ước tính sơ bộ theo thang điểm của môn (kèm lý giải chi tiết bằng tiếng Anh).
- **FR-3.2:** Phân loại tự động mức độ bài làm để gợi ý chế độ sửa:
  - **Mức độ Nhẹ (Minor Errors / High Quality):** Cấu trúc bài tốt, đáp ứng trên 80% yêu cầu rubric, chỉ mắc lỗi ngữ pháp, câu từ học thuật hoặc thiếu sót nhỏ.
  - **Mức độ Nặng (Major Errors / Incomplete):** Lập luận yếu, thiếu nhiều mục quan trọng trong đề bài hoặc viết sơ sài, lạc đề so với rubric.
- **FR-3.3:** Cho phép người dùng tùy ý ghi đè (override) chế độ sửa: buộc chạy sửa nhẹ, buộc chạy sửa nặng hoặc chỉ xuất báo cáo phân tích.

---

### 1.4. Chế độ Sửa Nhẹ (Light Fix Mode)
- **FR-4.1:** Chia bài viết thành các cụm đoạn văn (chunks) theo từng đề mục (headings), giữ nguyên ngữ cảnh phân cấp.
- **FR-4.2:** Truy vấn kho dữ liệu cục bộ để lấy 3–5 ví dụ sửa bài tương đồng nhất trong quá khứ kèm theo các quy tắc từ `style_guide.md`.
- **FR-4.3:** Mô hình AI thực hiện sửa đổi:
  - Sửa lỗi ngữ pháp, chính tả, cấu trúc câu tiếng Anh lủng củng.
  - Nâng cấp câu văn sang giọng văn học thuật trang trọng (Formal Academic Tone).
  - Trả về cấu trúc JSON gồm: đoạn gốc, đoạn sửa, phân loại lỗi và lý do giải thích (rationale) bằng tiếng Anh.
- **FR-4.4:** Sử dụng Microsoft Word COM (`CompareDocuments`) để so khớp tài liệu gốc và tài liệu đã sửa:
  - Xuất ra file `.docx` mới có đầy đủ tính năng **Track Changes** bản địa của Word (mỗi chỗ sửa đều hiển thị chữ gạch ngang đỏ và chữ thêm mới gạch chân).
  - Chèn tự động các nhận xét (**Comments**) của Word tại đúng vị trí đoạn văn được sửa để giải thích lý do cho sinh viên.

---

### 1.5. Chế độ Sửa Nặng & Bổ sung Nội dung (Heavy Rewrite & Augmentation)
- **FR-5.1:** Tái cấu trúc lại toàn bộ bài viết dựa trên cấu trúc đề bài chuẩn.
- **FR-5.2:** Giữ nguyên các ý tưởng, quan điểm và số liệu ban đầu của sinh viên, viết lại toàn bộ diễn đạt bằng văn phong học thuật chuẩn mực.
- **FR-5.3:** **Tự động bổ sung các đề mục và nội dung còn thiếu hụt** theo yêu cầu của Rubric:
  - Tạo lập luận logic, dẫn chứng bổ trợ học thuật.
  - Toàn bộ các đoạn văn hoặc câu do AI viết thêm bắt buộc phải được **tô sáng màu vàng nhạt (Highlighting)** trong file Word để người dùng nhận diện ngay lập tức.
  - Bất kỳ tài liệu tham khảo nào do AI đề xuất mới bắt buộc phải gắn tiền tố `[VERIFY]` để nhắc nhở người dùng kiểm tra lại tính có thực của nguồn học liệu.
- **FR-5.4:** Áp dụng toàn bộ quy chuẩn hình thức của `preset.yaml` (font, cỡ chữ, dãn dòng, margins) vào tài liệu đầu ra.

---

### 1.6. Vòng Lặp Học Phong Cách (Continual Learning & Feedback Loop)
- **FR-6.1:** Cho phép người dùng nạp vào hệ thống cặp bài: bài chưa sửa (`original.docx`) và bài đã sửa hoàn thiện (`final.docx` sau khi người dùng đã duyệt/chỉnh sửa).
- **FR-6.2:** Tự động tách biệt từng điểm khác biệt giữa hai tài liệu:
  - Phân tích cú pháp văn bản hoặc trích xuất Track Changes có sẵn trong file Word.
  - Gọi AI để phân loại loại lỗi: `grammar`, `academic_tone`, `argument_structure`, `rubric_expansion` và giải thích lý do vì sao người dùng sửa như vậy.
  - Lưu trữ vào tập tin `corrections/<pair_id>.jsonl`.
- **FR-6.3:** Tự động cập nhật chỉ mục tìm kiếm ngữ nghĩa (Semantic Embedding Index) cục bộ.
- **FR-6.4:** Cứ sau mỗi 5 cặp bài được nạp, hệ thống kích hoạt chức năng tổng hợp và đề xuất các quy tắc mới vào `style_guide.md`, hiển thị bản đối chiếu diff cho người dùng xác nhận trước khi lưu.

---

## 2. Đặc tả Lược đồ Dữ liệu (Data Schemas)

### 2.1. Cấu hình Môn học (`courses/<course_code>/preset.yaml`)

```yaml
course_code: "IB305"
course_name: "International Business Strategy"
updated_at: "2026-11-13T16:00:00Z"
formatting:
  font_family: "Times New Roman"
  font_size_body: 12
  font_size_h1: 16
  font_size_h2: 14
  font_size_h3: 12
  line_spacing: 1.5
  space_after: 6 # points
  margins:
    top: 2.54 # cm
    bottom: 2.54
    left: 2.54
    right: 2.54
  citation_style: "APA_7" # APA_7 | HARVARD
  caption_above_tables: true
  caption_below_figures: true

rubric:
  total_marks: 100
  criteria:
    - id: "crit_intro"
      name: "Introduction & Company Overview"
      weight_percent: 15
      description: "Clear introduction of the chosen multinational enterprise, market background, and thesis statement."
    - id: "crit_analysis"
      name: "Strategic Environmental Analysis"
      weight_percent: 35
      description: "Comprehensive application of PESTLE and Porter's Five Forces models with updated real-world data."
    - id: "crit_strategy"
      name: "Strategic Evaluation & Recommendations"
      weight_percent: 30
      description: "Realistic international entry modes and expansion strategies justified with critical thinking."
    - id: "crit_presentation"
      name: "Academic Structure & Referencing"
      weight_percent: 20
      description: "Impeccable academic tone, strict adherence to APA 7th edition, and seamless report structure."
```

---

### 2.2. Dữ liệu Đoạn Sửa Mẫu (`corrections/<pair_id>.jsonl`)

Mỗi dòng trong tệp `.jsonl` là một đối tượng JSON hợp lệ:

```json
{
  "correction_id": "corr_7f9c2d11-382a",
  "pair_id": "pair_20261113_01",
  "course_code": "IB305",
  "section_context": "2.1 PESTLE Analysis - Economic Factor",
  "original_text": "The inflation of UK is very big in 2023, making people not want to buy luxury goods anymore.",
  "revised_text": "In 2023, the United Kingdom experienced unprecedented inflationary pressures, severely eroding consumer purchasing power within the premium retail sector.",
  "category": "academic_tone",
  "rationale": "Replaced informal, conversational phrasing ('very big', 'not want to buy') with formal academic vocabulary and precise economic terminology ('unprecedented inflationary pressures', 'eroding consumer purchasing power').",
  "machine_id": "laptop-rtx3050",
  "created_at": "2026-11-13T16:15:20Z"
}
```

---

### 2.3. Quy tắc Phong cách Học thuật (`style_guide.md`)

Tệp Markdown lưu trên thư mục Google Drive dùng chung, quy định các nguyên tắc sửa văn phong được hệ thống học dần:

```markdown
# Living Academic Style Guide — AssignmentFixer

## 1. Academic Vocabulary & Register
- **Cấm tuyệt đối ngôn ngữ văn nói/thân mật:** Không sử dụng các từ như `a lot of`, `very big`, `huge`, `good`, `bad`, `thing`.
  - Thay bằng: `substantial`, `pronounced`, `favorable`, `adverse`, `phenomenon / factor`.
- **Hạn chế dùng đại từ ngôi thứ nhất:** Tránh dùng `I think`, `In my opinion`, `We can see`.
  - Chuyển thành câu bị động hoặc khách quan: `The empirical evidence indicates that...`, `It can be inferred that...`.

## 2. Sentence Structure & Cohesion
- Phá vỡ các câu quá dài trên 45 từ (run-on sentences) thành 2 câu có liên từ nối mạch lạc (`Furthermore`, `Consequently`, `In contrast`).
- Đảm bảo mỗi đoạn văn có cấu trúc chuẩn PEEL (Point -> Evidence -> Explanation -> Link).

## 3. Rubric Alignment Principles
- Khi bổ sung nội dung thiếu trong bài phân tích chiến lược, luôn gắn liền với số liệu tài chính hoặc báo cáo thường niên giai đoạn 2022–2025.
- Đánh dấu `[VERIFY]` cho mọi nguồn tài liệu tham khảo được sinh mới.
```

---

## 3. Đặc tả Giao diện Người dùng (UI Specification)

Giao diện ứng dụng được thiết kế bằng tiếng Việt, trực quan, đơn giản, phù hợp cho môi trường văn phòng:

### Màn hình 1: Sửa Bài (Home / Fix Assignment)
- **Vùng chọn tài liệu:**
  - Hộp kéo thả file bài làm sinh viên (hỗ trợ `.docx`, `.pdf`).
  - Hộp dropdown chọn Môn học (tự động tải danh sách từ thư mục `courses/` trên Google Drive).
- **Lựa chọn Chế độ Sửa:**
  - Radio button: **Tự động đề xuất (Recommended)** | **Chế độ Sửa nhẹ (Track Changes)** | **Chế độ Viết lại & Bổ sung (Heavy Rewrite)**.
- **Nút hành động:** `Bắt đầu Sửa bài` (Nổi bật).
- **Vùng tiến độ thời gian thực (Live Progress Log):**
  - Hiển thị thanh tiến trình (%) và danh sách các bước đang chạy (Đang kiểm tra hình thức -> Đang đối chiếu rubric -> Đang hiệu đính ngữ pháp -> Đang tạo Track Changes).
- **Vùng kết quả đầu ra:**
  - Điểm ước tính tổng thể và nhận xét tóm tắt.
  - Nút `Mở file Word kết quả` (mở ngay trong ứng dụng Microsoft Word của Windows).
  - Nút `Xem Báo cáo Đánh giá chi tiết`.

### Màn hình 2: Thư viện Môn & Rubric (Course Library)
- Danh sách các môn học hiện có dưới dạng thẻ (cards) hoặc bảng.
- Nút `+ Thêm Môn Mới`:
  - Nhập mã môn, tên môn.
  - Nút đính kèm tệp Đề bài (`brief`) và tệp Thang chấm (`rubric`).
  - Nút `Tự động trích xuất Preset bằng AI`.
- Bảng form cho phép xem và sửa trực tiếp: font chữ, cỡ chữ, kiểu dãn dòng, chuẩn trích dẫn, danh sách các tiêu chí và tỷ trọng điểm (%).
- Nút `Lưu Preset lên Google Drive`.

### Màn hình 3: Dạy Tool & Quy tắc Sửa (Learn & Style Guide)
- **Mục Nạp Cặp Bài Mới:**
  - Chọn bài gốc (`original.docx`) và bài sau khi duyệt (`final.docx`).
  - Nút `Bóc tách & Học lỗi`: Hệ thống tự động so sánh, phân loại lỗi và lưu vào kho.
- **Mục Xem & Chỉnh sửa Style Guide:**
  - Trình soạn thảo văn bản hỗ trợ xem và chỉnh sửa trực tiếp nội dung tệp `style_guide.md`.
  - Nút `Tổng hợp quy tắc mới từ dữ liệu gần đây` (gọi AI quét các đoạn sửa mới để tìm pattern quy tắc).

### Màn hình 4: Cài đặt Hệ thống (Settings)
- **Cấu hình AI Backend:**
  - Trạng thái Gemini CLI: Kiểm tra xem lệnh `gemini` đã được cài đặt và đăng nhập chưa.
  - Nút `Kiểm tra kết nối Gemini CLI` (gọi thử một prompt nhỏ để xác thực tài khoản Google AI Pro).
  - Ô nhập `Gemini API Key Dự phòng` (dùng khi CLI gặp sự cố).
- **Cấu hình Lưu trữ & Đồng bộ:**
  - Ô chọn thư mục Google Drive dùng chung (`Đường dẫn thư mục GDrive: E:\Shared Drive\AssignmentFixerData`).
  - Nút `Đồng bộ & Lập lại Chỉ mục Cục bộ` (quét lại các tệp `.jsonl` mới từ đối tác để cập nhật SQLite embedding).

---

## 4. Yêu cầu Phi Chức năng (Non-Functional Requirements)

1. **Hiệu năng & Tốc độ xử lý:**
   - Thời gian kiểm tra quy chuẩn hình thức (Rule-based): dưới 3 giây cho bài viết 5.000 từ.
   - Thời gian sửa bài toàn diện qua Gemini CLI: trung bình từ 90 đến 180 giây cho bài viết 3.000–5.000 từ.
2. **Khả năng tương thích phần cứng:**
   - Không đòi hỏi card đồ họa rời (chạy mượt trên cả CPU AMD văn phòng).
   - Mức chiếm dụng RAM tối đa không vượt quá 1.5 GB.
3. **Độ ổn định & Xử lý sự cố:**
   - Tự động đóng và giải phóng tiến trình `WINWORD.EXE` khi gặp ngoại lệ, không để lại tiến trình Word treo ngầm làm đầy bộ nhớ hoặc khóa file.
   - Cơ chế tự động thử lại (Exponential Backoff Retry) với độ trễ 2s, 4s, 8s khi Gemini CLI hoặc API gặp lỗi nghẽn mạng hoặc quá tải tạm thời.
4. **Bảo mật & Quyền riêng tư:**
   - Thông tin phiên đăng nhập và khóa API chỉ lưu cục bộ trên từng máy, không tải lên thư mục chia sẻ Google Drive.
