# Hướng Dẫn Chẩn Đoán Lỗi "Not Loaded" Của Add-in Trên Máy Tính Nội Bộ

Tài liệu này hướng dẫn chi tiết các bước kiểm tra, xác định nguyên nhân và cách khắc phục tình trạng AI Office Translate Add-in hiện trạng thái **"Not loaded"** (Không tải được) trên các máy tính trong mạng nội bộ công ty (thường do chính sách bảo mật GPO thiết lập).

---

## BƯỚC 1: Xác Định Thông Báo Lỗi Chi Tiết Trong Office

Khi Office báo trạng thái "Not loaded", hãy xem dòng **Load Behavior** ở dưới cùng để biết lý do cụ thể:

1. Mở **Excel**, **Word** hoặc **PowerPoint**.
2. Chọn **File** -> **Options** -> **Add-ins**.
3. Tại dòng **Manage** (ở dưới cùng), chọn **COM Add-ins** -> nhấn **Go...**.
4. Chọn **AI Office Translate** và đọc dòng **Load Behavior**:
   * **Trường hợp A:** *"Not loaded. A runtime error occurred during the loading of the COM Add-in."* -> Lỗi do không tìm thấy file, thiếu môi trường .NET, hoặc bị hệ thống chặn thực thi file DLL.
   * **Trường hợp B:** *"Not loaded. The user selected to disable..."* hoặc trống thông báo -> Office đã tắt add-in này do phát hiện khởi động chậm hoặc crash trước đó.
   * **Trường hợp C:** Báo lỗi liên quan đến *"Signature"*, *"Trust Center"*, hoặc *"Trusted Publisher"* -> Bị chặn bởi chính sách bảo mật chứng chỉ số.

---

## BƯỚC 2: Kiểm Tra Tệp Tin và Biên Dịch DLL

Add-in của chúng ta chạy bằng cách dùng compiler của hệ thống (`csc.exe`) biên dịch mã nguồn C# thành file DLL tại chỗ. Hãy kiểm tra xem file DLL này có được tạo thành công không:

1. Nhấn `Win + R`, gõ `%APPDATA%\AITranslateAddin` và nhấn **Enter**.
2. Kiểm tra xem thư mục có các file này không:
   * [x] `AITranslateAddin.dll` (Đây là file quan trọng nhất)
   * [x] `settings.txt`
   * [x] `Extensibility.dll`
   * [x] `Office.dll`

> [!WARNING]
> **Nếu không thấy file `AITranslateAddin.dll`:**
> Trình cài đặt `setup.exe` đã bị chặn biên dịch hoặc không tìm thấy `csc.exe` trên hệ thống. 
> * **Giải pháp:** Kiểm tra xem Windows Defender hoặc phần mềm diệt virus của công ty có đưa `setup.exe` hoặc thư mục AppData vào danh sách cách ly hay không.

---

## BƯỚC 3: Sửa Lỗi LoadBehavior Trong Windows Registry

Khi Office phát hiện Add-in load chậm hoặc gặp sự cố nhỏ, nó sẽ tự động đổi giá trị Registry của add-in từ `3` (Luôn load) thành `0` hoặc `2` (Không load). Bạn cần chỉnh thủ công:

1. Nhấn `Win + R`, gõ `regedit` và nhấn **Enter** để mở Registry Editor.
2. Tìm đến đường dẫn của các ứng dụng Office:
   * **Excel:** `HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\AITranslate.Excel`
   * **Word:** `HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\AITranslate.Word`
   * **PowerPoint:** `HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\AITranslate.PPT`
3. Nhấp đúp vào khóa **`LoadBehavior`** ở khung bên phải:
   * Đổi giá trị thành `3` (ở chế độ Base `Hexadecimal` hoặc `Decimal`).
   * Nếu khóa này không tồn tại hoặc không sửa được, có thể tài khoản của bạn bị GPO khóa quyền ghi registry nhánh này.

---

## BƯỚC 4: Kiểm Tra Danh Sách Disabled Items Trong Office

Nếu bạn đã sửa Registry thành `3` nhưng khi mở lại Office, giá trị này lại tự động quay về `0` hoặc `2`, nghĩa là add-in đã bị Office đưa vào danh sách đen (Hard Disabled Items):

1. Mở Word/Excel -> Chọn **File** -> **Options** -> **Add-ins**.
2. Ở mục **Manage** cuối trang, chọn **Disabled Items** -> nhấn **Go...**.
3. Nếu thấy **AI Office Translate** nằm trong danh sách này, chọn nó và nhấn **Enable**.
4. Khởi động lại Office.

---

## BƯỚC 5: Kiểm Tra Chính Sách AppLocker / Chặn Thực Thi DLL (%APPDATA%)

Đây là nguyên nhân rất phổ biến trên các máy tính Enterprise. Chính sách bảo mật của công ty thường cấm chạy bất kỳ file `.dll` hay `.exe` nào nằm trong thư mục người dùng như `C:\Users\<User>\AppData\`.

### Cách kiểm tra:
1. Mở **Event Viewer** của Windows (`Win + R` -> gõ `eventvwr.msc` -> Enter).
2. Vào **Applications and Services Logs** -> **Microsoft** -> **Windows** -> **AppLocker** -> **MSI and Script** hoặc **DLL**.
3. Tìm xem có sự kiện màu vàng/đỏ nào báo chặn tệp `AITranslateAddin.dll` hay không.

### Cách xử lý nếu bị AppLocker chặn:
Bạn phải chuyển Add-in sang cài đặt ở một thư mục được phép thực thi (thường là thư mục bạn tạo trực tiếp ở ổ đĩa `C:\` ví dụ `C:\AITranslateAddin` hoặc thư mục dự án của bạn).
1. Di chuyển thủ công toàn bộ thư mục `AITranslateAddin` từ `%APPDATA%` ra `C:\AITranslateAddin`.
2. Mở Registry (`regedit`) và cập nhật đường dẫn mới:
   * Tìm tất cả các khóa chứa đường dẫn cũ `%APPDATA%\AITranslateAddin\AITranslateAddin.dll` và thay thế bằng `C:\AITranslateAddin\AITranslateAddin.dll`.
   * Đặc biệt là khóa: `HKEY_CURRENT_USER\Software\Classes\CLSID\{GUID}\InprocServer32` (ở đây sẽ có các giá trị `CodeBase` trỏ tới file DLL).

---

## BƯỚC 6: Kiểm Tra Chính Sách Ký Số (Trust Center Settings)

Nếu máy công ty bắt buộc các Add-in phải được ký số:

1. Vào `File` -> `Options` -> `Trust Center` -> `Trust Center Settings...` -> `Add-ins`.
2. Kiểm tra xem dòng **"Require Application Add-ins to be signed by Trusted Publisher"** có đang bị tích chọn hay không.
3. Nếu có và bạn không thể bỏ tích (do bị IT khóa chính sách), bạn sẽ cần:
   * Tạo chứng chỉ tự ký (Self-signed certificate) bằng PowerShell.
   * Ký số vào file `AITranslateAddin.dll`.
   * Thêm chứng chỉ đó vào phân vùng **Trusted Root Certification Authorities** và **Trusted Publishers** trên máy tính nội bộ.

---

## Script PowerShell Chẩn Đoán Nhanh (Dành cho máy nội bộ)

Bạn có thể lưu đoạn mã sau thành file `.ps1` (ví dụ `diagnose.ps1`) và chạy bằng PowerShell trên máy công ty để kiểm tra nhanh trạng thái đăng ký của Add-in:

```powershell
Write-Host "--- KIỂM TRA TRẠNG THÁI AI OFFICE TRANSLATE ADD-IN ---" -ForegroundColor Cyan

# 1. Kiểm tra file DLL
$dllPath = "$env:APPDATA\AITranslateAddin\AITranslateAddin.dll"
if (Test-Path $dllPath) {
    Write-Host "[OK] Tìm thấy file DLL tại: $dllPath" -ForegroundColor Green
} else {
    Write-Host "[ERR] Không tìm thấy file DLL tại: $dllPath" -ForegroundColor Red
}

# 2. Kiểm tra LoadBehavior trong Registry
$apps = @("Excel", "Word", "PowerPoint")
foreach ($app in $apps) {
    $regPath = "HKCU:\Software\Microsoft\Office\$app\Addins\AITranslate.$($app -eq 'PowerPoint' -and 'PPT' -or $app)"
    if (Test-Path $regPath) {
        $val = (Get-ItemProperty -Path $regPath -Name "LoadBehavior" -ErrorAction SilentlyContinue).LoadBehavior
        if ($val -eq 3) {
            Write-Host "[OK] $app: LoadBehavior đã được đặt là 3 (Luôn load)." -ForegroundColor Green
        } else {
            Write-Host "[WARN] $app: LoadBehavior hiện tại là $val (Cần sửa lại thành 3)." -ForegroundColor Yellow
        }
    } else {
        Write-Host "[ERR] $app: Không tìm thấy khóa đăng ký add-in trong Registry." -ForegroundColor Red
    }
}

# 3. Kiểm tra đăng ký COM Class
$clsidExcel = "HKCU:\Software\Classes\CLSID\{8A16298B-77E5-46D9-B582-7EF2C3F6F6B1}"
if (Test-Path $clsidExcel) {
    Write-Host "[OK] Đã đăng ký COM Class cho Excel Add-in." -ForegroundColor Green
} else {
    Write-Host "[ERR] Chưa đăng ký COM Class trong registry. Cần chạy lại setup.exe." -ForegroundColor Red
}

Write-Host "----------------------------------------------------" -ForegroundColor Cyan
```
