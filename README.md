# AI Office Translate Plugin & Backend Server

Dự án cung cấp giải pháp dịch thuật tài liệu Microsoft Office (Word, Excel, PowerPoint) sử dụng mô hình AI của OpenAI (như GPT-4o-mini hoặc GPT-4o). Hệ thống bao gồm:
1. **Node.js Express Backend API Server**: Xử lý logic nghiệp vụ, gọi API OpenAI, tích hợp bảo mật, giới hạn tần suất (rate limiting) và chạy dưới giao thức HTTPS sử dụng chứng chỉ tự sinh.
2. **Microsoft Office Add-in (Frontend)**: Chạy trực tiếp bên trong Word, Excel, PowerPoint dưới dạng Taskpane hiện đại với hiệu ứng kính mờ (glassmorphism) sang trọng.

---

## Các Tính Năng Nổi Bật

- **Bảo toàn định dạng văn bản (Style Preservation)**: 
  - *Word/PPT*: Lấy nội dung HTML, tự động dọn dẹp và tối giản hóa (loại bỏ 90% thuộc tính CSS rác của Word) trước khi dịch. Điều này giúp LLM dịch chuẩn xác theo ngữ cảnh của câu và giữ nguyên định dạng in đậm, in nghiêng, gạch chân, liên kết, bảng biểu, danh sách của tài liệu.
  - *Excel*: Dịch các ô được chọn bằng mảng dữ liệu văn bản thô thâu tóm trong một API duy nhất và ghi đè lại giá trị. Excel tự giữ lại định dạng ô (màu nền, viền, font chữ).
- **Bảo mật & Tránh quá tải**:
  - Xác thực bằng **Client Access Token** được lưu cấu hình cả ở server `.env` và client `localStorage`.
  - Hạn chế spam và chống quá tải hệ thống nội bộ bằng **Rate Limiting** (giới hạn số lượng request từ mỗi IP).
  - Giới hạn payload truyền tải tối đa **5MB** tránh tràn bộ nhớ.
- **Hoạt động offline hoàn toàn (Local Air-gapped Ready)**:
  - Loại bỏ hoàn toàn các liên kết CDN và font trực tuyến từ Internet.
  - Sử dụng font chữ hệ thống **Segoe UI** có sẵn trên Windows.
  - Thư viện **Office.js** được phục vụ cục bộ từ `node-modules` ngay trên server nội bộ của bạn.
- **Tương thích 3 chế độ dịch**:
  - Dịch phần bôi đen (Translate Selection).
  - Dịch đối tượng xem hiện tại (Dịch trang hiện tại trong Word, sheet hiện tại trong Excel, slide hiện tại trong PowerPoint).
  - Dịch toàn bộ tài liệu (Translate All).

---

## Hướng Dẫn Cài Đặt và Chạy Thử (Local Testing)

### 1. Cài đặt các gói phụ thuộc (Dependencies)
Ở thư mục gốc dự án, cài đặt các package Node.js:
```bash
npm install
```

### 2. Định cấu hình môi trường (.env)
Dự án đã có tệp `.env` được tạo sẵn. Bạn có thể mở ra để cấu hình lại các khóa:
- `CLIENT_ACCESS_TOKEN`: Token bảo mật khách hàng gửi lên (mặc định là `secure-token-123`).
- `OPENAI_API_KEY`: Điền OpenAI API Key của bạn để sử dụng thực tế. Nếu giữ nguyên `mock`, server sẽ chạy ở **chế độ mô phỏng dịch thuật (Mock Mode)**, tự động thêm tiền tố ngôn ngữ vào văn bản để bạn kiểm thử luồng hoạt động mà không tốn chi phí API.
- `OPENAI_MODEL`: Mẫu AI muốn sử dụng (mặc định `gpt-4o-mini` để tối ưu chi phí và tốc độ).

### 3. Tạo chứng chỉ SSL/TLS tự ký (Chạy HTTPS)
Bởi vì Microsoft Office bắt buộc các Web Add-in phải chạy qua kết nối HTTPS an toàn, bạn cần tạo khóa SSL nội bộ bằng cách chạy lệnh:
```bash
npm run certs
```
*Lệnh này sử dụng `node-forge` tạo tự động hai tệp `backend/key.pem` và `backend/cert.pem` tương thích với trình duyệt.*

### 4. Khởi chạy Server
Khởi động backend server và serve thư mục tĩnh chứa add-in:
```bash
npm start
```
Server sẽ chạy tại địa chỉ: `https://localhost:3000`.

### 5. Xác thực tin cậy Chứng chỉ SSL trên Trình duyệt (RẤT QUAN TRỌNG)
Vì đây là chứng chỉ SSL tự ký (self-signed) chạy trên localhost, các trình duyệt web và Office sẽ chặn tải tài nguyên iframe nếu bạn chưa tin cậy nó:
1. Mở trình duyệt (Chrome/Edge/Firefox).
2. Truy cập đường dẫn: **`https://localhost:3000/addon/taskpane.html`**
3. Bạn sẽ thấy cảnh báo bảo mật màu đỏ (Your connection is not private).
4. Click chọn **Advanced** (Nâng cao) -> Chọn **Proceed to localhost (unsafe)** (Tiếp tục truy cập localhost).
5. Khi trình duyệt tải được giao diện màu tím mờ của add-in nghĩa là chứng chỉ đã được chấp nhận và tin cậy trên máy tính của bạn.

---

## Hướng Dẫn Sideload Add-in vào Office trên Windows

Không cần thiết lập chia sẻ thư mục mạng (Network Share) phức tạp, chúng tôi cung cấp script đăng ký thông tin trực tiếp vào Windows Registry:

### Bước 1: Đăng ký Registry
1. Mở cửa sổ **PowerShell** (không cần quyền Administrator).
2. Di chuyển vào thư mục dự án và chạy lệnh sau:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\register-addin.ps1
   ```
3. Màn hình console sẽ báo "Đăng ký thành công!".

### Bước 2: Kích hoạt Add-in trong Office
1. Khởi động ứng dụng **Microsoft Word**, **Excel** hoặc **PowerPoint**.
2. Trên thanh Ribbon, chọn tab **Insert** (Chèn) -> chọn **My Add-ins** (Add-in của tôi).
3. Trong cửa sổ hiện ra, chọn tab **Developer Add-ins** (Add-in dành cho nhà phát triển).
4. Nhấp đúp vào **AI Office Translate**.
5. Bạn sẽ thấy một tab menu mới xuất hiện trên thanh công cụ tên là **Translate** chứa 3 tính năng dịch thuật.

### Bước 3: Cấu hình Token trên Add-in
1. Click vào bất kỳ tính năng dịch nào trên thanh Ribbon để mở bảng Taskpane bên phải.
2. Click vào **biểu tượng bánh răng cài đặt** (Settings Gear) ở góc trên bên phải của Taskpane.
3. Nhập **Client Access Token** trùng khớp với token cấu hình trong tệp `.env` (mặc định là `secure-token-123`).
4. Ấn **Lưu Cài Đặt**. Bây giờ bạn đã sẵn sàng sử dụng!

---

## Cách Gỡ Cài Đặt (Uninstall)
Khi muốn gỡ cài đặt add-in khỏi Office và xóa khóa đăng ký Registry, chạy lệnh sau trong PowerShell:
```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\unregister-addin.ps1
```
Sau đó khởi động lại các ứng dụng Office.
