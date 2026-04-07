# KHNV Import / Export Rules

Tài liệu này ghi lại các ràng buộc đã chốt cho luồng import file local, xem Excel theo workbook và xuất DOCX trong module `CHITIEU`.

## 1. Import file local

- Người dùng chỉ được import file `.xlsx`.
- Tên file import phải bắt đầu bằng `CTKHNV_TW` hoặc `CTKHNV_DP`.
- `CTKHNV_TW*.xlsx` khi import thành công phải thay thế file nguồn `INPUT/TW.xlsx`.
- `CTKHNV_DP*.xlsx` khi import thành công phải thay thế file nguồn `INPUT/DP.xlsx`.
- Trên giao diện phần import phải được gom thành 1 khối thao tác gọn.
- Bên ngoài khối import chỉ hiển thị 1 nút gọn dạng button `Import Excel`; khi bấm nút này mới mở layout con để chọn `TW/DP`, chọn file và thực hiện import.
- Khi panel import chưa được mở, các control bên trong panel phải ở trạng thái disabled; chỉ enable khi người dùng mở đúng panel import.
- Trong khối import đó vẫn phải có lựa chọn rõ ràng `TW` hoặc `DP` để người dùng chọn workbook cần thay.
- Sau khi import xong phải hiển thị thông báo thành công và nói rõ workbook nào đã được thay.
- Nút `Đọc lại mẫu` phải nạp lại workbook đang xem trên màn hình.

## 2. Dữ liệu hiển thị trên `CHITIEU/CODE/index.php`

- Màn hình phải có lựa chọn xem `TW` hoặc `DP`.
- Khi chuyển lựa chọn `TW/DP`, bảng Excel phải nạp đúng workbook tương ứng:
  - `TW` -> `INPUT/TW.xlsx`
  - `DP` -> `INPUT/DP.xlsx`
- Header của bảng không được hard-code cố định theo số nhóm cột.
- Các tiêu đề nhóm phải đọc từ dữ liệu Excel.
- Các cột tiêu đề con phải sinh động theo workbook đang mở.
- Nếu tiêu đề đọc được là `0`, rỗng hoặc `null` thì không dùng giá trị đó để render tiêu đề.
- Nút `Lưu cập nhật` chỉ được ghi vào workbook đang xem.
- Trạng thái view Excel trên web phải giữ nguyên kiểu hiển thị hiện có: bảng động theo workbook, không thêm cột rỗng và không hard-code số dòng.
- Khi vừa tải màn hình, panel import/export không được tự mở theo URL/query; chỉ mở khi người dùng bấm trực tiếp vào nút thao tác tương ứng.

## 3. Quy tắc xuất DOCX

- Luồng export phải đọc đồng thời cả `INPUT/TW.xlsx` và `INPUT/DP.xlsx`.
- Nếu người dùng đang sửa workbook hiện tại mà chưa bấm lưu, payload của view đang mở vẫn phải được áp vào context export trước khi tạo file.
- Phải hỗ trợ 3 chế độ export:
  - `TT` -> xuất mẫu `OUTPUT/Dieu_chinh_chi_tieu.docx`
  - `DMDN` -> xuất mẫu `OUTPUT/To_trinh.docx`
  - `ALL` -> xuất 1 file `.zip` chứa cả 2 mẫu trên
- Phần export trên giao diện phải được gom theo cùng kiểu với import: 1 khối thao tác, trong đó người dùng chọn loại mẫu rồi bấm xuất.
- Bên ngoài khối export chỉ hiển thị 1 nút gọn dạng button `Xuất DOCX`; khi bấm nút này mới mở layout con để chọn `TT`, `DMDN` hoặc `ALL`.
- Khi panel export chưa được mở, các control chọn mẫu/nút xuất bên trong panel phải ở trạng thái disabled; chỉ enable khi người dùng mở đúng panel export.
- DOCX phải dùng đúng dữ liệu import hiện tại.
- Tất cả placeholder trong template DOCX phải được thay bằng nội dung thật.
- Không được để sót chuỗi dạng `{{...}}` trong file xuất ra.

## 4. Lọc dữ liệu khi xuất

- Nếu `Điều chỉnh tăng trưởng` bằng `0` thì không đưa dòng đó vào DOCX.
- Nếu `Điều chỉnh tăng trưởng` là rỗng hoặc `NULL` thì không đưa dòng đó vào DOCX.
- Nếu giá trị điều chỉnh là lỗi kiểu `#REF!` hoặc không phải số thì xem như không hợp lệ để xuất.
- Chỉ những dòng có `Điều chỉnh tăng trưởng` khác `0` mới được xuất.
- Nếu một nhóm không còn dòng hợp lệ thì không tạo khối xuất cho nhóm đó.
- Không được sinh page trống đầu file khi không có dữ liệu hợp lệ.

## 5. Mẫu DOCX

- Mẫu `TT`: `OUTPUT/Dieu_chinh_chi_tieu.docx`
- Mẫu `DMDN`: `OUTPUT/To_trinh.docx`

### 5.1. Placeholder của mẫu `Dieu_chinh_chi_tieu.docx`

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
- `{{nam_ke_hoach}}`

### 5.2. Placeholder của mẫu `To_trinh.docx`

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

### 5.3. Quy ước thay dữ liệu vào mẫu

- `{{phong_giao_dich}}` và `{{PHONG_GIAO_DICH}}` của mẫu `TT` phải bằng tên PGD thật của block đang render.
- `{{ten_pgd_tw}}` và `{{ten_pgd_dp}}` của mẫu `DMDN` phải bằng tên PGD thật.
- `{{don_vi}}` mặc định là `Triệu Đồng`.
- Khi render mẫu, phải giữ nguyên bố cục bảng và kiểu chữ của DOCX gốc.
- Mẫu `TT` render theo từng PGD, trong mỗi block có 2 section `TW` và `DP`.
- Mẫu `DMDN` render theo cấu trúc `chương trình -> PGD -> xã` cho từng nguồn `TW` và `DP`.

## 6. Nguyên tắc sửa tiếp

- Luôn đọc file này trước khi chỉnh `CHITIEU/CODE/index.php`, `data.php`, `import.php` hoặc `export.php`.
- Khi có yêu cầu mới, ưu tiên giữ nguyên các ràng buộc trên trừ khi người dùng nói rõ muốn đổi.
- Không được tự ý nới điều kiện lọc `0/null`.
- Không được ghi đè các quy tắc đã chốt bằng logic hard-code mới.

## 7. Cách tổ chức code

- `CODE/data.php` chỉ giữ hằng số chung, helper dùng chung và `require_once` sang 2 module chức năng.
- `CODE/import.php` chứa luồng đọc workbook, import file local, lưu Excel và tiền xử lý dữ liệu.
- `CODE/export.php` chứa luồng xuất DOCX, dựng context từ `TW/DP`, thay placeholder và đóng gói file tải xuống.
- Khi sửa import/export, luôn kiểm tra hàm đang nằm đúng module để tránh vá nhầm chỗ.

## 8. Quy định text tiếng Việt

- Tất cả chuỗi hiển thị ra giao diện hoặc tài liệu phải là UTF-8 chuẩn.
- Không được để lại chuỗi mojibake.
- Nếu sửa thông báo hoặc tiêu đề mới, phải kiểm tra lại bằng tiếng Việt có dấu trước khi chốt.

## 9. Ghi nhớ màn hình Excel

- Màn hình `CHITIEU/CODE/index.php` phải bám theo workbook đang xem.
- Nếu Excel có 3 dòng dữ liệu thì chỉ hiển thị đúng 3 dòng dữ liệu.
- Nếu Excel có 2 nhóm/cột `Cho vay ...` thì hiển thị đúng 2 nhóm/cột; nếu có 5 nhóm/cột thì hiển thị đúng 5 nhóm/cột.
- Không được render thêm danh sách PGD hoặc cột rỗng do logic cũ.

## 10. Ghi nhớ file mẫu đang active

- File mẫu export hiện tại là:
  - `OUTPUT/Dieu_chinh_chi_tieu.docx`
  - `OUTPUT/To_trinh.docx`

## 11. Ghi nhớ lựa chọn mẫu sau này

- `TT`: xuất `OUTPUT/Dieu_chinh_chi_tieu.docx`
- `DMDN`: xuất `OUTPUT/To_trinh.docx`
- `ALL`: xuất 1 lần cả 2 mẫu trên dưới dạng file `.zip`
- Khi người dùng nhắc tới `TT`, `DMDN` hoặc `ALL`, hiểu đây là lựa chọn loại export của module `CHITIEU`.
