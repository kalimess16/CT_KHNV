# KHNV Import / Export Rules

Tài liệu này ghi lại các ràng buộc đã chốt cho luồng import file local, đọc lại mẫu và xuất DOCX.

## 1. Import file local

- Người dùng chọn file từ máy tính qua giao diện `index.php`.
- File import bắt buộc phải là `.xlsx`.
- Tên file phải bắt đầu bằng `CTKHNV_TW` hoặc `CTKHNV_DP`.
- Nếu file import bắt đầu bằng `CTKHNV_TW` thì khi import thành công sẽ thay thế file nguôn trong `INPUT/TW.xlsx`.
- Nếu file import bắt đầu bằng `CTKHNV_DP` thì khi import thành công sẽ thay thế file nguôn trong `INPUT/DP.xlsx`.
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


## 4. Lọc dữ liệu khi xuất

- Nếu `Điều chỉnh tăng trưởng` bằng `0` thì không đưa dòng đó vào DOCX.
- Nếu `Điều chỉnh tăng trưởng` là rỗng hoặc `NULL` thì không đưa dòng đó vào DOCX.
- Chỉ những dòng có `Điều chỉnh tăng trưởng` khác `0` mới được xuất.
- Nếu một nhóm không còn dòng nào hợp lệ thì không tạo khối xuất cho nhóm đó.
- Không được sinh page trống đầu file khi không có nhóm hợp lệ.

## 5. Mẫu DOCX

- Mẫu PGD_XA: `OUTPUT/Dieu_chinh_chi_tieu.docx`
- Mẫu TT_PGD: `OUTPUT/To_trinh.docx`
- Mẫu PGD_XA phải giữ đủ các thẻ sau để export thay thế:
  - `{{PHONG_GIAO_DICH}}`
  - `{{ten_xa_tw}}`
  - `{{ten_xa_dp}}`
  - `{{stt_tw}}`
  - `{{stt_dp}}`
  - `{{chuong_trinh_tw}}`
  - `{{chuong_trinh_dp}}`
  - `{{dieu_chinh_tw}}`
  - `{{dieu_chinh_dp}}`
  - `{{ctkh_tw}}`
  - `{{ctkh_dp}}`
  - `{{don_vi}}`
- Mẫu TT_PGD phải giữ đủ các thẻ sau để export thay thế:
  - `{{ten_pgd_tw}}`
  - `{{ten_pgd_dp}}`
  - `{{stt_tw}}`
  - `{{stt_dp}}`
  - `{{tong_denghi_tw}}`
  - `{{tong_denghi_dp}}`
  - `{{denghi_ct_tw}}`
  - `{{denghi_ct_dp}}`
  - `{{chuong_trinh_ct_tw}}`
  - `{{chuong_trinh_ct_dp}}`
  - `{{ten_xa_tw}}`
  - `{{ten_xa_dp}}`
  - `{{don_vi}}`
  - `{{dagiao_pgd_tw}}`
  - `{{dagiao_pgd_dp}}`
  - `{{dagiao_xa_tw}}`
  - `{{dagiao_xa_dp}}`
  - `{{denghi_pgd_tw}}`
  - `{{denghi_pgd_dp}}`
  - `{{denghi_xa_tw}}`
  - `{{denghi_xa_dp}}`
  - `{{ctkh_pgd_tw}}`
  - `{{ctkh_pgd_dp}}`
  - `{{ctkh_xa_tw}}`
  - `{{ctkh_xa_dp}}`

- `{{ten_xa_dp}}`; `{{ten_xa_tw}}` dùng cho tên xã/phường trong khối bảng của mẫu PGD_XA và mẫu TT_PGD.
- `{{PHONG_GIAO_DICH}}` mẫu PGD_XA bằng với `{{ten_pgd_tw}}`; `{{ten_pgd_dp}}` tại mẫu TT_PGD.
- `{{don_vi}}` mặc định là "Triệu Đồng"
- Khi chuyển mẫu, phải giữ nguyên bố cục bảng và kiểu chữ của DOCX gốc.
- đối với mẫu thì TT_PGD chỉnh sửa lại table cho gọn

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

## 10. Ghi nhớ mới

- Màn hình `CHITIEU/CODE/index.php` phải bám theo workbook đang import:
- Nếu Excel có 3 dòng dữ liệu thì chỉ hiển thị đúng 3 dòng dữ liệu.
- Nếu Excel có 2 nhóm/cột `Cho vay ...` thì hiển thị đúng 2 nhóm/cột; nếu có 5 nhóm/cột thì hiển thị đúng 5 nhóm/cột.
- Không được render thêm danh sách PGD hoặc cột rỗng/`0` do logic hard-code cũ.
- Khi export DOCX, mỗi PGD phải ra đúng block của chính nó; không được dồn danh sách tên PGD lên đầu trang đầu.

## 11. Ghi nhớ file mẫu

- File mẫu đang active để export hiện tại là  `OUTPUT/Dieu_chinh_chi_tieu.docx` và `OUTPUT/To_trinh.docx`


## 12. Ghi nhớ lựa chọn mẫu sau này

- Dự kiến bổ sung 3 lựa chọn xuất:
- Loại 1: `TT` : xuất Mẫu PGD_XA: `OUTPUT/Dieu_chinh_chi_tieu.docx`.
- Loại 2: `DMDN` : xuất Mẫu TT_PGD: `OUTPUT/To_trinh.docx`.
- Loại 3: `ALL` : xuất 1 lần 2 mẫu trên
- Khi người dùng yêu cầu sau này, hiểu đây là yêu cầu chia export theo 3 chế độ chọn loại trên.


