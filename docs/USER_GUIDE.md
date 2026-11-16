# Hướng dẫn Cài đặt & Sử dụng — AssignmentFixer

Tài liệu này hướng dẫn chi tiết từng bước thiết lập môi trường và vận hành công cụ **AssignmentFixer** cho cả máy của bạn và laptop của đối tác (partner).

---

## 1. Yêu cầu Hệ thống & Chuẩn bị Môi trường

Thực hiện các bước sau trên cả 2 máy tính Windows:

### 1.1. Kiểm tra Microsoft Word
- Máy tính phải cài đặt sẵn **Microsoft Office Word 2016 trở lên** (Word 2019, 2021 hoặc Microsoft 365) và đã được kích hoạt bản quyền đầy đủ.
- Đảm bảo ứng dụng Word có thể mở và lưu file `.docx` bình thường.

### 1.2. Cài đặt Python 3.12+
- Tải và cài đặt Python từ [python.org](https://www.python.org/downloads/) (nếu máy chưa có).
- **Lưu ý quan trọng khi cài:** Tích chọn vào ô **"Add Python to PATH"**.

### 1.3. Cài đặt Node.js & Gemini CLI
Gemini CLI được sử dụng để kết nối trực tiếp với tài khoản **Google AI Pro** nhằm hưởng hạn ngạch 1.500 lượt gọi/ngày miễn phí thêm.

1. Tải và cài đặt phiên bản Node.js LTS từ [nodejs.org](https://nodejs.org/).
2. Mở cửa sổ dòng lệnh PowerShell hoặc Command Prompt và tiến hành cài đặt Gemini CLI:
   ```cmd
   npm.cmd install -g @google/gemini-cli
   ```
   *(Ghi chú: Nếu PowerShell báo lỗi script bị chặn do Execution Policy, dùng lệnh `npm.cmd` thay vì `npm`)*.
3. Xác thực tài khoản Google AI Pro:
   - Chạy lệnh sau trên terminal:
     ```cmd
     gemini
     ```
   - Trình duyệt web sẽ tự động mở trang đăng nhập Google. Đăng nhập bằng tài khoản **Google AI Pro** của bạn (hoặc của partner).
   - Sau khi cấp quyền thành công, terminal sẽ thông báo xác thực thành công. Gõ `/exit` để thoát phiên tương tác.

---

## 2. Thiết lập Thư mục Dữ liệu Dùng Chung (Google Drive)

Để hai máy đồng bộ dữ liệu kho bài và quy tắc sửa:

1. Tạo một thư mục trên Google Drive cá nhân, ví dụ: `AssignmentFixerData`.
2. Chia sẻ quyền chỉnh sửa (Editor) thư mục này cho tài khoản Google của partner.
3. Cài đặt ứng dụng **Google Drive for Desktop** trên cả 2 máy để thư mục này xuất hiện dưới dạng một ổ đĩa hoặc đường dẫn thư mục cục bộ trên Windows, ví dụ:
   - Trên máy 1: `G:\My Drive\AssignmentFixerData`
   - Trên máy 2: `G:\Shared drives\AssignmentFixerData`
4. Cấu trúc thư mục sẽ được ứng dụng tự động khởi tạo khi chạy lần đầu:
   ```
   AssignmentFixerData/
   ├── courses/          # Chứa đề bài, rubric và preset theo từng môn
   ├── pairs/            # Chứa các cặp bài gốc và bài duyệt để dạy tool
   ├── corrections/      # Chứa các bản ghi sửa đổi chi tiết (.jsonl)
   └── style_guide.md    # Bộ quy tắc phong cách học thuật living
   ```

---

## 3. Cài đặt và Khởi chạy Ứng dụng

1. Tải hoặc clone mã nguồn về máy:
   ```cmd
   git clone <repository_url> E:\AssignmentFixer
   cd /d E:\AssignmentFixer
   ```
2. Cài đặt các thư viện phụ thuộc của Python:
   ```cmd
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```
3. Khởi chạy ứng dụng:
   ```cmd
   python -m app.main
   ```
   Giao diện ứng dụng desktop sẽ tự động hiển thị trong một cửa sổ riêng biệt.

---

## 4. Hướng dẫn Sử dụng Từng Chức năng

### 4.1. Thiết lập Cấu hình Ban đầu (Màn hình "Cài đặt")
1. Vào tab **Cài đặt**.
2. Nhập hoặc duyệt chọn đường dẫn đến thư mục `AssignmentFixerData` trên Google Drive máy bạn.
3. Bấm nút **Kiểm tra kết nối Gemini CLI**:
   - Nếu hiển thị tích xanh và thông báo kết nối thành công: Hệ thống đã sẵn sàng sử dụng tài khoản AI Pro.
4. *(Tùy chọn)* Nhập **Gemini API Key** từ Google AI Studio vào ô API Key dự phòng (để hệ thống tự chuyển đổi khi CLI gặp sự cố).
5. Bấm **Lưu cấu hình**.

---

### 4.2. Thêm Môn học Mới & Thiết lập Rubric (Màn hình "Thư viện Môn")
Khi sinh viên có môn học mới hoặc assignment mới:
1. Vào tab **Thư viện Môn**, bấm **Thêm Môn Mới**.
2. Điền **Mã môn** (ví dụ: `MGT202`) và **Tên môn** (ví dụ: `Business Communication`).
3. Đính kèm tệp Đề bài (`brief.docx` hoặc `.pdf`) và tệp Thang chấm (`rubric.docx` hoặc `.pdf`).
4. Bấm **Tự động trích xuất cấu hình bằng AI**:
   - AI sẽ tự đọc rubric để lấy: Font chữ yêu cầu, cỡ chữ, dãn dòng, chuẩn trích dẫn (APA 7 hay Harvard), các tiêu chí chấm điểm và tỷ trọng %.
5. Kiểm tra lại thông tin trên biểu mẫu, chỉnh sửa nếu cần, sau đó bấm **Lưu Môn học**. Môn học này sẽ lập tức khả dụng trên cả máy của bạn và partner qua Google Drive.

---

### 4.3. Sửa Bài Assignment (Màn hình "Sửa bài")
1. Chọn bài làm của sinh viên cần sửa (hỗ trợ cả `.docx` và `.pdf`).
2. Chọn Môn học tương ứng từ danh sách thả xuống.
3. Chọn Chế độ sửa:
   - **Tự động đề xuất (Khuyên dùng):** AI phân tích bài trước; nếu bài làm tốt sẽ chọn sửa nhẹ, nếu bài quá sơ sài sẽ chọn sửa nặng.
   - **Chế độ Sửa nhẹ (Track Changes):** Giữ nguyên sườn bài, chỉ sửa câu từ, ngữ pháp và chèn nhận xét (Word Comments).
   - **Chế độ Sửa nặng (Viết lại & Bổ sung):** Viết lại bài học thuật hoàn chỉnh, bổ sung các phần còn thiếu theo rubric.
4. Bấm **Bắt đầu Sửa bài**.
5. Theo dõi tiến trình xử lý trên màn hình (thông thường mất khoảng 1.5 – 3 phút tùy độ dài bài).
6. Khi hoàn tất:
   - Xem điểm ước tính và các nhận xét sơ bộ ngay trên app.
   - Bấm **Mở file Word kết quả**: Ứng dụng sẽ tự mở tệp kết quả trực tiếp trong Microsoft Word trên máy tính của bạn.

---

### 4.4. Quy trình Duyệt & Kiểm tra Kết quả trong Word
- **Đối với tệp Sửa nhẹ:**
  - Mở file trong Microsoft Word, bạn sẽ thấy toàn bộ các điểm sửa đổi hiển thị dưới dạng **Track Changes** (chữ gạch đỏ là chữ sinh viên viết sai, chữ gạch chân là chữ AI đã sửa theo phong cách học thuật).
  - Phía lề phải là các bong bóng nhận xét (**Comments**) giải thích lỗi sai.
  - Bạn có thể nhấn `Accept` hoặc `Reject` từng chỗ sửa, hoặc chỉnh sửa thêm theo ý mình, sau đó bấm Save.
- **Đối với tệp Sửa nặng:**
  - File mới được định dạng chuẩn mực theo quy định của môn.
  - **Lưu ý các đoạn văn bản được bôi màu vàng nhạt (Highlight):** Đây là các luận điểm hoặc đề mục do AI viết thêm để bù đắp các phần sinh viên bị thiếu theo rubric.
  - **Lưu ý các trích dẫn có gắn `[VERIFY]`:** Đây là các nguồn học liệu do AI đề xuất thêm, bạn hãy đối soát hoặc kiểm chứng lại trên Google Scholar để đảm bảo tính xác thực.

---

### 4.5. Dạy Tool & Nâng cao Phong cách Sửa (Màn hình "Dạy tool")
Càng nạp nhiều bài đã duyệt, hệ thống sẽ sửa bài càng ngày càng giống ý bạn:
1. Vào tab **Dạy tool**.
2. Chọn tệp bài gốc của sinh viên (`original.docx`).
3. Chọn tệp bài sau khi bạn đã kiểm duyệt và hoàn thiện (`final.docx`).
4. Bấm **Bóc tách & Học lỗi**:
   - Ứng dụng sẽ bóc tách các điểm sửa, phân loại lỗi và lưu vào thư mục `corrections/` trên Google Drive.
   - Chỉ mục tìm kiếm ngữ nghĩa cục bộ sẽ tự động cập nhật để các bài sửa tiếp theo sẽ lấy các mẫu này làm ví dụ đối chiếu.
5. Kiểm tra mục **Living Style Guide**: Bạn có thể đọc và sửa đổi trực tiếp các quy tắc trong văn bản `style_guide.md`. Bấm **Tổng hợp quy tắc mới** khi muốn AI tự phát hiện xu hướng sửa đổi chung từ các bài gần đây.

---

## 5. Xử lý Sự cố Thường gặp (Troubleshooting & FAQ)

### Q1: Bị lỗi "Word is busy" hoặc tiến trình Word bị treo?
- **Nguyên nhân:** Có thể một tiến trình Word đang mở hộp thoại thông báo bản quyền hoặc popup lưu file ngầm.
- **Cách khắc phục:**
  1. Đóng toàn bộ các cửa sổ Microsoft Word đang mở trên máy tính.
  2. Mở Task Manager (Ctrl + Shift + Esc), tìm mục `Microsoft Word` trong danh sách Processes và chọn `End task`.
  3. Chạy lại tác vụ sửa bài trên ứng dụng.

### Q2: Gemini CLI báo lỗi xác thực hoặc hết hạn phiên?
- **Cách khắc phục:**
  1. Mở cửa sổ dòng lệnh và gõ `gemini`.
  2. Gõ lệnh `/login` hoặc đăng nhập lại tài khoản Google AI Pro qua trình duyệt.
  3. Hoặc vào màn hình Cài đặt của app và điền Gemini API Key dự phòng để tiếp tục công việc ngay lập tức.

### Q3: Tệp PDF sinh viên gửi có nhiều bảng biểu phức tạp bị vỡ khung khi mở?
- **Cách khắc phục:** Khi nạp file PDF, ứng dụng sử dụng công cụ chuyển đổi PDF sang DOCX của Word. Với các bài có biểu đồ quá phức tạp, khuyên bạn nên yêu cầu sinh viên gửi file gốc dạng `.docx` để đảm bảo giữ nguyên vẹn cấu trúc bảng biểu và hình ảnh.

### Q4: Máy của partner có xem được các bài tôi đã dạy không?
- **Có.** Vì toàn bộ các cặp bài và tệp `.jsonl` được lưu trên Google Drive chung. Khi partner mở app, ứng dụng sẽ tự động phát hiện các tệp mới và lập chỉ mục vào bộ nhớ cục bộ của máy partner mà không cần thao tác thủ công.
