# KHNV Import / Export Rules

Tài liệu này ghi lại các ràng buộc đã chốt cho luồng import file local, đọc lại mẫu và xuất DOCX.

## 1. Import file local

- Người dùng chọn file từ máy tính qua giao diện `index.php`.
- File import bắt buộc phải là `.xlsx`.
- Tên file phải bắt đầu bằng `CTKHNV`.
- Nếu import thành công, file mới sẽ thay thế file nguồn trong `INPUT/test.xlsx`.
- Sau khi import xong phải hiển thị thông báo thành công.
- Nút `Đọc lại mẫu` phải nạp lại dữ liệu từ file vừa import.

## 2. Dữ liệu hiển thị trên `index.php`

- Header của bảng không được hard-code cố định theo số nhóm cột.
- Các tiêu đề nhóm phải đọc từ dữ liệu Excel.
- Các cột tiêu đề con phải sinh động theo workbook đang mở.
- Nếu tiêu đề đọc được là `0` hoặc rỗng/null thì không dùng giá trị đó để render tiêu đề.

## 3. Quy tắc xuất DOCX

- DOCX phải dùng đúng dữ liệu đang import hiện tại.
- Tiêu đề DOCX phải bám theo workbook.
- Các placeholder trong template DOCX phải được thay bằng nội dung thật.
- Không được để sót các chuỗi dạng `{{...}}` trong file xuất ra.
- `STT` trong DOCX phải là số thứ tự hiển thị hợp lệ, không lấy sai từ dữ liệu công thức.

## 4. Lọc dữ liệu khi xuất

- Nếu `Điều chỉnh tăng trưởng` bằng `0` thì không đưa dòng đó vào DOCX.
- Nếu `Điều chỉnh tăng trưởng` là rỗng hoặc `NULL` thì không đưa dòng đó vào DOCX.
- Chỉ những dòng có `Điều chỉnh tăng trưởng` khác `0` mới được xuất.
- Nếu một nhóm không còn dòng nào hợp lệ thì không tạo khối xuất cho nhóm đó.
- Không được sinh page trống đầu file khi không có nhóm hợp lệ.

## 5. Mẫu DOCX

- Mẫu 1: `OUTPUT/MAU.docx`
- Mẫu 2: `OUTPUT/MAU_NEW.docx`
- Mẫu 3: `OUTPUT/1.docx`
- Mẫu đang active để export hiện tại: `OUTPUT/1.docx`
- Mẫu 2 phải giữ đủ các thẻ sau để export thay thế:
  - `{{TIEU_DE_1}}`
  - `{{PHONG_GIAO_DICH}}`
  - `{{GHI_CHU_1}}`
  - `{{GHI_CHU_2}}`
  - `{{DON_VI}}`
  - `{{tenxa}}`
  - `{{DIEU_CHINH_TANG_TRUONG}}`
  - `{{CHI_TIEU_KE_HOACH}}`
- `{{tenxa}}` dùng cho tên xã/phường trong khối bảng của mẫu 2.
- Khi chuyển mẫu, phải giữ nguyên bố cục bảng và kiểu chữ của DOCX gốc.

## 6. Nguyên tắc sửa tiếp

- Luôn đọc file này trước khi chỉnh `index.php`, `data.php`, `import.php` hoặc `export.php`.
- Khi có yêu cầu mới, ưu tiên giữ nguyên các ràng buộc trên trừ khi người dùng nói rổ muốn đổi.
- Không được tự ý nới điều kiện lọc `0/null` nếu chưa được yêu cầu.
- Không được ghi đè các quy tắc đã chốt bằng logic hard-code mới.

## 7. Cách tổ chức code

- `CODE/data.php` chỉ giữ các hàng số chung, hàm tiện ích và `require_once` sang 2 module chức năng.
- `CODE/import.php` chứa luồng đọc workbook, import file local, lưu Excel và tiền xử lý dữ liệu.
- `CODE/export.php` chứa luồng xuất DOCX, thay placeholder và render bảng theo workbook.
- Khi sửa import/export, luôn kiểm tra hàm đang nằm trong `data.php`, `import.php` hay `export.php` để tránh vá nhầm chỗ.

## 8. Quy đị text tiếng Việt

- Tất cả chuỗi hiển thị ra giao diện hoặc tài liệu phải là UTF-8 chuẩn.
- Không được để lại chuỗi mojibake; hãy kiểm tra lại toàn bộ tiếng Việt có dấu trước khi chốt.
- Nếu sửa thông báo hoặc tiêu đề mới, phải kiểm tra lại bằng tiếng Việt có dấu trước khi chốt.
## 9. Ghi nhớ bố cục

- `PHONG_GIAO_DICH` phải nằm ngay dưới `TIEU_DE_1`, trước các dòng ghi chú và `DON_VI`.

## 10. Ghi nhớ mới

- Màn hình `CHITIEU/CODE/index.php` phải bám theo workbook đang import:
- Nếu Excel có 3 dòng dữ liệu thì chỉ hiển thị đúng 3 dòng dữ liệu.
- Nếu Excel có 2 nhóm/cột `Cho vay ...` thì hiển thị đúng 2 nhóm/cột; nếu có 5 nhóm/cột thì hiển thị đúng 5 nhóm/cột.
- Không được render thêm danh sách PGD hoặc cột rỗng/`0` do logic hard-code cũ.
- Khi export DOCX, mỗi PGD phải ra đúng block của chính nó; không được dồn danh sách tên PGD lên đầu trang đầu.

## 11. Ghi nhớ file mẫu

- File mẫu đang active để export hiện tại là `CHITIEU/OUTPUT/1.docx`.
- Hai file mẫu còn lại đang có trong thư mục là `CHITIEU/OUTPUT/MAU.docx` và `CHITIEU/OUTPUT/MAU_NEW.docx`.
- Khi người dùng nói chỉnh file mẫu, ưu tiên kiểm tra 3 file trên trước vì đây là các file mẫu cần giữ lại.

## 12. Ghi nhớ lựa chọn mẫu sau này

- Dự kiến bổ sung 3 lựa chọn xuất:
- Mẫu 1: `DP`
- Mẫu 2: `TW`
- Mẫu 3: `ALL`
- Khi người dùng yêu cầu sau này, hiểu đây là yêu cầu chia export theo 3 chế độ chọn mẫu ở trên.
