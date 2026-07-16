# AI Office Translate - VSTO Version

Phiên bản VSTO (Visual Studio Tools for Office) của Add-in dịch thuật sử dụng AI dành cho Microsoft Excel, Word và PowerPoint. Phiên bản này được thiết kế tối ưu cho môi trường doanh nghiệp có chính sách bảo mật nghiêm ngặt (GPO), cho phép cài đặt và kích hoạt thành công mà **không cần quyền Quản trị viên (Admin)**.

---

## 1. Hướng Dẫn Biên Dịch (Dành cho Lập trình viên)

### Yêu Cầu Hệ Thống (Môi trường Build)
*   **Hệ điều hành:** Windows 10 / 11.
*   **Visual Studio:** Phiên bản 2019 / 2022 trở lên.
*   **Workload bắt buộc:** Gói công cụ **Office/SharePoint development** (Phát triển Office/SharePoint) phải được cài đặt trên Visual Studio.

### Các Bước Biên Dịch ra file `setup.exe` duy nhất:
1.  Mở PowerShell tại thư mục `vsto-version/`.
2.  Chạy script build tự động:
    ```powershell
    powershell -ExecutionPolicy Bypass -File .\BuildVSTO.ps1
    ```
3.  Sau khi build thành công, bạn sẽ nhận được một file cài đặt độc lập duy nhất tại đường dẫn:
    `vsto-version/setup.exe`

---

## 2. Hướng Dẫn Khắc Phục Lỗi Khi Build (Troubleshooting)

Nếu khi chạy `BuildVSTO.ps1` bạn gặp lỗi biên dịch liên quan đến:
*   `warning MSB3245: Could not locate the assembly "Microsoft.Office.Tools..."`
*   `error CS0234: The type or namespace name 'ExcelAddInBase' / 'WordAddInBase' / 'AddInBase' does not exist...`

> [!IMPORTANT]
> **Nguyên nhân:** Máy tính của bạn đang thiếu gói Workload phát triển Office trong Visual Studio.

### Cách xử lý từng bước:
1.  Mở chương trình **Visual Studio Installer** trên máy tính của bạn.
2.  Tìm phiên bản Visual Studio đang sử dụng và nhấn nút **Modify** (Sửa).
3.  Tại tab **Workloads**, cuộn xuống nhóm **Other Toolsets** (Các bộ công cụ khác).
4.  Tích chọn vào ô **Office/SharePoint development** (Phát triển Office/SharePoint).
5.  Nhấn nút **Modify** (Sửa) ở góc dưới bên phải để bắt đầu tải xuống và cài đặt bổ sung.
6.  Sau khi hoàn tất, mở lại PowerShell và chạy lại lệnh build `.\BuildVSTO.ps1`. Lỗi sẽ biến mất hoàn toàn.

---

## 3. Hướng Dẫn Cài Đặt (Dành cho Người dùng cuối / Máy công ty)

Ưu điểm lớn nhất của trình cài đặt `setup.exe` tự động này là hoạt động hoàn toàn dưới quyền User thường (**không cần Run as Administrator**):

1.  **Sao chép:** Copy duy nhất file `setup.exe` (đã được biên dịch từ bước 1) sang máy tính nội bộ của công ty.
2.  **Chạy cài đặt:** Nhấp đúp chuột để mở `setup.exe`.
3.  **Thực hiện:**
    *   Nhấn nút **INSTALL** để cài đặt. Trình cài đặt sẽ tự động:
        *   Giải nén mã nguồn nhị phân vào thư mục `%APPDATA%\AITranslateVSTO`.
        *   Tự động đăng ký chứng chỉ bảo mật tự ký đi kèm vào khu vực Tin cậy (`Trusted Publishers` và `Root` của CurrentUser).
        *   Ghi Registry Manifest VSTO cho Excel, Word, PowerPoint.
    *   Nhấn nút **UNINSTALL** nếu bạn muốn gỡ cài đặt sạch sẽ khỏi máy tính.
4.  **Hoàn thành:** Khởi động lại Excel, Word và PowerPoint để bắt đầu sử dụng Add-in dịch thuật AI.
