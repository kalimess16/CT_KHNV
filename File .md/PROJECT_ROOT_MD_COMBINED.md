# CT_KHNV Root Markdown Combined

Tài liệu này gộp nội dung từ các file `.md` ở thư mục gốc của project `CT_KHNV`.

## Phạm vi gộp

- Chỉ gộp các file `.md` nằm trực tiếp trong thư mục gốc `CT_KHNV/`.
- Bỏ qua toàn bộ file `.md` nằm trong các thư mục con.
- Không thay thế hoặc xóa các file nguồn gốc.

## Nguồn đã gộp

- `YEU_CAU_CODEX.md`
- `HOME_INDEX_RULES.md`
- `IMPORT_EXPORT_RULES.md`
- `CHITIEU_DOCX_PLACEHOLDER_RULES.md`
- `CHITIEU_CLEAR_DOCX_RULES.md`
- `ACCESS_CONTROL_RULES.md`

## 1. Quy ước làm việc với Codex

### 1.1. Điểm vào mặc định

- Khi bạn bảo `doc file yeu cau`, Codex đọc `YEU_CAU_CODEX.md` trước.
- Sau đó Codex ưu tiên đọc các file `.md` liên quan trong dự án để lấy đúng quy tắc và bối cảnh trước khi sửa code.
- Nếu đụng đến trang chủ hoặc giao diện tổng, ưu tiên đọc `HOME_INDEX_RULES.md`.
- Nếu đụng đến luồng import/export, DOCX, Excel hoặc module `CHITIEU`, ưu tiên đọc `IMPORT_EXPORT_RULES.md`.
- Khi có yêu cầu mới, có thể sửa nội dung trong `YEU_CAU_CODEX.md` rồi dùng câu: `Hay doc YEU_CAU_CODEX.md va lam theo.`
- Sau khi hoàn thành một yêu cầu hoặc tính năng mới chưa có trong tài liệu, cần bổ sung vào file `.md` liên quan để lưu lại.

### 1.2. Cách ghi yêu cầu

- Mỗi nhóm công việc viết theo từng thư mục để Codex dễ theo dõi.
- Mỗi mục nên có đủ: `Trang thai`, `Muc do uu tien`, `File can`, `Yeu cau`, `Cach kiem tra`, `Ghi chu them`.
- Có thể thêm nhiều mục trong cùng một thư mục nếu có nhiều việc riêng.

### 1.3. Giá trị nên dùng

- `Trang thai`: `TODO`
- `Muc do uu tien`: `cao` | `trung binh` | `thap`

### 1.4. Mẫu yêu cầu theo thư mục

#### Thư mục gốc `CT_KHNV/`

- Trang thai: chua lam
- Muc do uu tien: cao
- File can:
  - Doc truoc: `YEU_CAU_CODEX.md`
  - Doc them neu lien quan: `HOME_INDEX_RULES.md`
  - File sua/tao: `index.php`, `assets/home.css`
- Yeu cau:
  - Mô tả rõ cần sửa ở trang chủ, layout, nút bấm, khu vực nhúng module hoặc logic điều hướng.
- Cach kiem tra:
  - Mở `index.php` trên trình duyệt.
  - Chạy `php -l index.php`.
  - Bấm thử các nút `Chi tieu`, `Ke hoach`, `Trang chu`.
- Ghi chu them:
  - Ghi rõ ràng buộc nào không được đổi.
  - Nếu cần giữ nguyên UI hoặc luồng cũ thì ghi rõ.

#### Thư mục `assets/`

- Trang thai: chua lam
- Muc do uu tien: trung binh
- File can:
  - Doc truoc: `YEU_CAU_CODEX.md`
  - Doc them neu lien quan trang chu: `HOME_INDEX_RULES.md`
  - File sua/tao: `assets/home.css` hoặc file asset liên quan
- Yeu cau:
  - Mô tả rõ cần đổi CSS, hình ảnh, icon, khoảng cách, màu sắc, responsive.
- Cach kiem tra:
  - Tải lại trang và kiểm tra trên desktop/mobile.
  - Kiểm tra có vỡ layout, lệch chữ, mất nút, tràn nội dung hay không.
- Ghi chu them:
  - Nếu chỉ được sửa giao diện thì ghi rõ `khong sua logic`.

#### Thư mục `CHITIEU/`

- Trang thai: chua lam
- Muc do uu tien: cao
- File can:
  - Doc truoc: `YEU_CAU_CODEX.md`
  - Doc bat buoc neu sua module nay: `IMPORT_EXPORT_RULES.md`
  - Doc them neu module nay dang duoc nhung trong trang chu: `HOME_INDEX_RULES.md`
  - File sua/tao: `CHITIEU/index.php`, `CHITIEU/CODE/index.php`, `CHITIEU/CODE/data.php`, `CHITIEU/CODE/import.php`, `CHITIEU/CODE/export.php`, `CHITIEU/CODE/style.css`, file trong `INPUT/`, `OUTPUT/` nếu có nêu rõ
- Yeu cau:
  - Mô tả rõ đang sửa import, export, đọc workbook, render bảng, file mẫu DOCX, bộ lọc dữ liệu hoặc giao diện module.
- Cach kiem tra:
  - Chạy `php -l CHITIEU\\CODE\\index.php`
  - Import file `CTKHNV*.xlsx`
  - Bấm `Doc lai mau`
  - Export DOCX và kiểm tra placeholder `{{...}}` đã được thay hết
  - Kiểm tra số dòng và số nhóm cột đúng theo workbook
- Ghi chu them:
  - Nếu có file mẫu đang active thì ghi rõ tên file.
  - Nếu có quy tắc `khong duoc sua` về filter, placeholder, bố cục thì ghi rõ.

#### Thư mục `kehoach/`

- Trang thai: chua lam
- Muc do uu tien: thap
- File can:
  - Doc truoc: `YEU_CAU_CODEX.md`
  - Doc them file `.md` liên quan nếu sau này dự án bổ sung
  - File sua/tao: ghi rõ file cụ thể trong `kehoach/`
- Yeu cau:
  - Mô tả rõ module cần tạo mới, trang tạm, UI thông báo hoặc logic riêng.
- Cach kiem tra:
  - Mô tả cách bấm thử, mở đường dẫn hoặc test nghiệp vụ cần xác nhận.
- Ghi chu them:
  - Nếu tạm thời chỉ hiện `Dang phat trien cho` thì ghi rõ để tránh nối nhầm sang `CHITIEU`.

### 1.5. Nội dung yêu cầu hiện có trong `YEU_CAU_CODEX.md`

#### Thu mục `CHITIEU/`

- Trang thai: `TODO`
- Muc do uu tien: `cao`
- File can:
  - Doc truoc: `YEU_CAU_CODEX.md`
  - Doc bat buoc neu sua module nay: `IMPORT_EXPORT_RULES.md`
  - Doc them neu module nay dang duoc nhung trong trang chu: `HOME_INDEX_RULES.md`
  - File sua/tao: `CHITIEU/index.php`, `CHITIEU/CODE/index.php`, `CHITIEU/CODE/data.php`, `CHITIEU/CODE/import.php`, `CHITIEU/CODE/export.php`, `CHITIEU/CODE/style.css`, file trong `INPUT/`, `OUTPUT/` nếu có nêu rõ
- Yeu cau:
  - Xem ảnh và xử lý ô đỏ: chỗ này hàm `SUM` nếu dương thì có số, nhưng khi âm lại bị đưa về `0`; yêu cầu là vẫn phải thể hiện tổng âm.
- Cach kiem tra:
  - Vào xuất DOCX và tải file về để kiểm tra.
  - Nhấn `Xóa` thì số liệu phải mất hết.
- Ghi chu them:
  - Nếu có file mẫu đang active thì ghi rõ tên file.
  - Nếu có quy tắc `khong duoc sua` về filter, placeholder, bố cục thì ghi rõ.

## 2. Quy tắc trang chủ `CT_KHNV/index.php`

### 2.1. Mục đích trang chủ

- `index.php` ở thư mục gốc `CT_KHNV` là trang chủ điều hướng nghiệp vụ.
- Trang này không chứa logic import/export Excel hay DOCX.
- Trang chủ chỉ đóng vai trò hiển thị các thẻ chức năng và điều hướng người dùng.

### 2.2. Chức năng `Kế hoạch`

- Thẻ có `kicker` là `Kế hoạch` là một chức năng riêng biệt.
- Chức năng `Kế hoạch` hiện tại không liên quan tới thư mục `CHITIEU`.
- Không được gắn link của `Kế hoạch` sang `CHITIEU/index.php` hay `CHITIEU/CODE/index.php`.
- Khi người dùng bấm vào thẻ `Kế hoạch`, hệ thống chỉ hiển thị thông báo `Đang phát triển chờ`.
- Nếu chưa có module riêng cho `Kế hoạch`, phải giữ nguyên hành vi này.

### 2.3. Chức năng `Chỉ tiêu`

- Thẻ `Chỉ tiêu` là chức năng đang hoạt động.
- Thẻ này được phép điều hướng sang module trong thư mục `CHITIEU`.
- Luồng hiện tại dùng `CHITIEU/index.php`.
- Khi sửa tiếp, không được làm ảnh hưởng tới luồng import/export hiện có trong `CHITIEU`.

### 2.4. Quy ước giao diện trang chủ

- Font giao diện trang chủ đang dùng: `Times New Roman`.
- Trang chủ dùng CSS tại `assets/home.css`.
- Header đang hiển thị full chiều ngang màn hình.
- Bên dưới header có vùng hero/showcase full chiều ngang để hiển thị logo hoặc nội dung chức năng.
- Hero dùng hiệu ứng sáng nhẹ theo màu logo và có slogan `Thấu hiểu lòng dân, tận tâm, phục vụ`.
- Nền hero phải còn sắc hồng nhận diện của logo, không bị bạc sang gần trắng.
- Cuối trang có footer ngắn hiển thị dòng credit `Creat by @TinHoc_DN`.
- Cụm logo + tiêu đề ở đầu trang là điểm bấm để quay về trạng thái trang chủ.
- Cụm logo + tiêu đề không hiển thị gạch chân kiểu link.
- Ưu tiên giữ cụm này là phần tử điều hướng nhẹ, không để browser áp style native kiểu form-control làm xuất hiện nền trắng trên header.
- Nếu cụm này dùng `button` để reset về trang chủ thì phải reset style mặc định của button.
- Không dùng selector quá rộng kiểu `.brand-home-trigger *` để ép style toàn bộ phần con.
- Khi dọn hoặc sửa giao diện, ưu tiên giữ bố cục gọn, dễ nhìn và không để khối tiêu đề quá lớn gây vỡ layout.

### 2.5. Quy ước tải chức năng trong trang chủ

- Khi người dùng bấm vào chức năng đang hoạt động, nội dung phải tải ngay trong vùng nội dung phía dưới của `CT_KHNV/index.php`.
- Không chuyển toàn bộ trình duyệt sang trang khác nếu chức năng đó hỗ trợ nhúng trong trang chủ.
- Với module `CHITIEU`, luồng nhúng trong trang chủ phải dùng đúng nguồn `CHITIEU/index.php`.
- Nếu cần chế độ nhúng riêng cho module thì phải giữ nguyên nghiệp vụ hiện có và chỉ tinh chỉnh hiển thị.
- Khi chức năng được mở trong vùng nhúng, không cần thêm thanh tiêu đề riêng kiểu `Chức năng đang mở`.

### 2.6. Quy ước hiển thị module `CHITIEU` khi nhúng

- Khi `CHITIEU` mở trong trang chủ, không hiển thị lại header riêng, topbar riêng hay khối hero cũ của module.
- Trong chế độ nhúng, phần đầu của `CHITIEU` chỉ giữ lại đúng câu `Nhập file local theo chuẩn CTKHNV_DP/TW*.xlsx`.
- Câu này phải nằm trong khu toolbar/trạng thái của module, không nằm ở một khối riêng phía trên.
- Trong chế độ nhúng, toolbar chỉ nên hiện 2 nút gọn `Import Excel` và `Xuất DOCX`; panel con chỉ mở khi người dùng bấm.
- Module không được tự mở panel import/export khi vừa tải trang.
- Trên desktop, cụm `Đọc lại mẫu` và `Lưu cập nhật` nên nằm cùng hàng với 2 nút gọn để tiết kiệm chiều cao.
- Trên desktop, cụm `Import Excel` và `Xuất DOCX` nên bám sát trạng thái bên trái; cụm `TW/DP` và `Đọc lại mẫu/Lưu cập nhật` nên gom về cột phải.
- Trên desktop, 4 nút `Import Excel`, `Xuất DOCX`, `Đọc lại mẫu`, `Lưu cập nhật` phải nằm ngay vùng đầu toolbar.
- Khi bấm `Trang chủ KHNV` từ `CHITIEU` trong chế độ nhúng, không được mở lồng một trang chủ mới trong khung nhúng.
- Vùng bảng của `CHITIEU` khi nhúng phải hiển thị theo layout rộng ở phần dưới của trang chủ, không mất chức năng, không sinh khoảng trắng lớn bất thường phía dưới.
- Ưu tiên tăng không gian theo chiều dọc của vùng cuộn bên phải, không kéo lệch bố cục theo chiều ngang.
- Phần nhúng hiện đang dùng `body.embedded .table-wrap { max-height: calc(100vh - 175px); }`.
- Phần nhúng hiện đang tụt vào hai bên `5px` mỗi bên.
- Khi sửa tiếp, không để nội dung bị cắt ở mép trái làm mất chữ hoặc mất một phần nút/chức năng.

### 2.7. Kiểm tra sau khi sửa

- Khi sửa `CT_KHNV/index.php`, luôn kiểm tra:
  - `Kế hoạch` còn là chức năng riêng.
  - `Kế hoạch` không trỏ vào `CHITIEU`.
  - `Kế hoạch` vẫn báo `Đang phát triển chờ` nếu chưa có module riêng.
  - `Chỉ tiêu` vẫn trỏ đúng vào module `CHITIEU`.
  - Chức năng hoạt động được tải trong vùng nội dung bên dưới thay vì điều hướng toàn trang.
- Khi sửa `CHITIEU/CODE/index.php` hoặc `CHITIEU/CODE/style.css`, luôn kiểm tra:
  - Chế độ nhúng vẫn hiển thị đủ toolbar, nút thao tác và bảng.
  - Không sinh khoảng trắng lớn phía dưới bảng.
  - Không làm mất chữ ở mép trái khi mở trong trang chủ.
- Sau khi sửa `index.php`, nên chạy:
  - `php -l index.php`
  - `php -l CHITIEU\CODE\index.php`

## 3. Quy tắc module `CHITIEU`

### 3.1. Import file local

- Người dùng chỉ được import file `.xlsx`.
- Tên file import phải bắt đầu bằng `CTKHNV_TW` hoặc `CTKHNV_DP`.
- `CTKHNV_TW*.xlsx` khi import thành công phải thay thế `INPUT/TW.xlsx`.
- `CTKHNV_DP*.xlsx` khi import thành công phải thay thế `INPUT/DP.xlsx`.
- Phần import trên giao diện phải gom thành một khối thao tác gọn.
- Bên ngoài khối import chỉ hiển thị 1 nút `Import Excel`; khi bấm mới mở layout con để chọn `TW/DP`, chọn file và thực hiện import.
- Khi panel import chưa mở, các control bên trong phải disabled.
- Trong khối import vẫn phải có lựa chọn rõ ràng `TW` hoặc `DP`.
- Sau khi import xong phải hiển thị thông báo thành công và nói rõ workbook nào đã được thay.
- Nút `Đọc lại mẫu` phải nạp lại workbook đang xem trên màn hình.

### 3.2. Dữ liệu hiển thị trên `CHITIEU/CODE/index.php`

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

### 3.3. Quy tắc xuất DOCX

- Luồng export phải đọc đồng thời cả `INPUT/TW.xlsx` và `INPUT/DP.xlsx`.
- Nếu người dùng đang sửa workbook hiện tại mà chưa bấm lưu, payload của view đang mở vẫn phải được áp vào context export trước khi tạo file.
- Phải hỗ trợ 3 chế độ export:
  - `TT` -> `OUTPUT/Dieu_chinh_chi_tieu.docx`
  - `DMDN` -> `OUTPUT/To_trinh.docx`
  - `ALL` -> 1 file `.zip` chứa cả 2 mẫu trên
- Phần export trên giao diện phải gom theo cùng kiểu với import: một khối thao tác, trong đó người dùng chọn loại mẫu rồi bấm xuất.
- Bên ngoài khối export chỉ hiển thị 1 nút `Xuất DOCX`; khi bấm mới mở layout con để chọn `TT`, `DMDN` hoặc `ALL`.
- Khi panel export chưa mở, các control chọn mẫu/nút xuất bên trong phải disabled.
- DOCX phải dùng đúng dữ liệu import hiện tại.
- Tất cả placeholder trong template DOCX phải được thay bằng nội dung thật.
- Không được để sót chuỗi dạng `{{...}}` trong file xuất ra.

### 3.4. Lọc dữ liệu khi xuất

- Nếu `Điều chỉnh tăng trưởng` bằng `0` thì không đưa dòng đó vào DOCX.
- Nếu `Điều chỉnh tăng trưởng` là rỗng hoặc `NULL` thì không đưa dòng đó vào DOCX.
- Nếu giá trị điều chỉnh là lỗi kiểu `#REF!` hoặc không phải số thì xem như không hợp lệ để xuất.
- Chỉ những dòng có `Điều chỉnh tăng trưởng` khác `0` mới được xuất.
- Nếu một nhóm không còn dòng hợp lệ thì không tạo khối xuất cho nhóm đó.
- Không được sinh page trống đầu file khi không có dữ liệu hợp lệ.

### 3.5. Mẫu DOCX và placeholder

- Mẫu `TT`: `OUTPUT/Dieu_chinh_chi_tieu.docx`
- Mẫu `DMDN`: `OUTPUT/To_trinh.docx`

#### Placeholder của mẫu `Dieu_chinh_chi_tieu.docx`

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

#### Placeholder của mẫu `To_trinh.docx`

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

#### Quy ước thay dữ liệu vào mẫu

- `{{phong_giao_dich}}` và `{{PHONG_GIAO_DICH}}` của mẫu `TT` phải bằng tên PGD thật của block đang render.
- `{{ten_pgd_tw}}` và `{{ten_pgd_dp}}` của mẫu `DMDN` phải bằng tên PGD thật.
- `{{don_vi}}` mặc định là `Triệu Đồng`.
- Khi render mẫu, phải giữ nguyên bố cục bảng và kiểu chữ của DOCX gốc.
- Mẫu `TT` render theo từng PGD, trong mỗi block có 2 section `TW` và `DP`.
- Mẫu `DMDN` render theo cấu trúc `chương trình -> PGD -> xã` cho từng nguồn `TW` và `DP`.

### 3.6. Placeholder bị tách run

- Placeholder trong `word/document.xml` có thể bị tách qua nhiều `w:r` hoặc `w:t`.
- Logic thay placeholder phải xử lý được cả khi root đang render là toàn bộ `body`, `table` hoặc một `paragraph` clone riêng.
- Không được để sót `{{phong_giao_dich}}`, `{{PHONG_GIAO_DICH}}` hoặc placeholder tương tự chỉ vì template tách chuỗi thành nhiều run.

### 3.7. Lưu ý riêng cho mẫu `TT`

- Mẫu `TT` active là `CHITIEU/OUTPUT/Dieu_chinh_chi_tieu.docx`.
- `{{phong_giao_dich}}` phải lấy đúng tên PGD của từng block đang render.
- Sau khi render xong một PGD, luồng xuất phải tiếp tục render block PGD tiếp theo trong cùng file DOCX.
- Khi render mẫu `TT`, `{{phong_giao_dich}}` và `{{PHONG_GIAO_DICH}}` phải đưa tên PGD về chữ hoa.
- Sau khi export `TT`, không được còn placeholder dạng `{{...}}` trong file DOCX sinh ra.
- Trường `{{phong_giao_dich}}` phải hiện tên PGD thật trong tiêu đề từng block.

### 3.8. Layout dòng DOCX

- Các dòng dữ liệu clone vào bảng DOCX không được giữ `w:trHeight` cố định nếu làm dòng bị dư chiều cao.
- Dòng nội dung ngắn phải tự co theo text; dòng dài vẫn được phép cao lên theo nội dung thực tế.

### 3.9. Xóa số liệu workbook

- Màn hình `CHITIEU/CODE/index.php` có nút `Xóa số liệu` cho workbook đang xem (`TW` hoặc `DP`).
- Khi xóa, tất cả ô số liệu từ vùng dữ liệu trở xuống phải về rỗng.
- Công thức Excel phải được giữ lại.
- Nếu toàn bộ ô nguồn của công thức đều rỗng thì giá trị hiển thị của ô công thức cũng phải rỗng.

### 3.10. Tổng âm ở ô công thức tổng hợp

- Các ô tổng hợp dùng `SUM(...)` hoặc cộng trừ nhiều ô trong module `CHITIEU` nếu tính ra âm thì phải giữ nguyên tổng âm, không được ép về `0`.
- Khi công thức tổng hợp có chen thành phần lỗi cũ như `#REF!` nhưng vẫn còn các hạng hợp lệ khác, ưu tiên tính trên phần hợp lệ còn lại để không làm mất tổng âm.

### 3.11. Nguyên tắc sửa tiếp

- Luôn đọc tài liệu này trước khi chỉnh `CHITIEU/CODE/index.php`, `data.php`, `import.php` hoặc `export.php`.
- Khi có yêu cầu mới, ưu tiên giữ nguyên các ràng buộc trên trừ khi người dùng nói rõ muốn đổi.
- Không được tự ý nới điều kiện lọc `0/null`.
- Không được ghi đè các quy tắc đã chốt bằng logic hard-code mới.

### 3.12. Cách tổ chức code

- `CODE/data.php` chỉ giữ hằng số chung, helper dùng chung và `require_once` sang 2 module chức năng.
- `CODE/import.php` chứa luồng đọc workbook, import file local, lưu Excel và tiền xử lý dữ liệu.
- `CODE/export.php` chứa luồng xuất DOCX, dựng context từ `TW/DP`, thay placeholder và đóng gói file tải xuống.
- Khi sửa import/export, luôn kiểm tra hàm đang nằm đúng module để tránh vá nhầm chỗ.

### 3.13. Quy định text tiếng Việt

- Tất cả chuỗi hiển thị ra giao diện hoặc tài liệu phải là UTF-8 chuẩn.
- Không được để lại chuỗi mojibake.
- Nếu sửa thông báo hoặc tiêu đề mới, phải kiểm tra lại bằng tiếng Việt có dấu trước khi chốt.

### 3.14. Ghi nhớ màn hình Excel

- Màn hình `CHITIEU/CODE/index.php` phải bám theo workbook đang xem.
- Nếu Excel có 3 dòng dữ liệu thì chỉ hiển thị đúng 3 dòng dữ liệu.
- Nếu Excel có 2 nhóm/cột `Cho vay ...` thì hiển thị đúng 2 nhóm/cột; nếu có 5 nhóm/cột thì hiển thị đúng 5 nhóm/cột.
- Không được render thêm danh sách PGD hoặc cột rỗng do logic cũ.

### 3.15. Ghi nhớ lựa chọn mẫu

- `TT`: xuất `OUTPUT/Dieu_chinh_chi_tieu.docx`
- `DMDN`: xuất `OUTPUT/To_trinh.docx`
- `ALL`: xuất 1 lần cả 2 mẫu trên dưới dạng file `.zip`
- Khi người dùng nhắc tới `TT`, `DMDN` hoặc `ALL`, hiểu đây là lựa chọn loại export của module `CHITIEU`.

## 4. Quy tắc access control

### 4.1. Phạm vi áp dụng

- Thư mục gốc `CT_KHNV/` có `.htaccess` để chặn web request từ IP không nằm trong whitelist.
- Trang chủ gốc `index.php` phải chạy kiểm tra IP trước khi render giao diện.
- Module `CHITIEU` khi vào qua `CHITIEU/index.php` hoặc truy cập thẳng `CHITIEU/CODE/index.php` đều phải bị ràng buộc bởi cùng whitelist.
- Không được đưa logic whitelist vào CSS hay JavaScript giao diện.
- Các thư mục dữ liệu `CHITIEU/INPUT` và `CHITIEU/OUTPUT` không được phép tải trực tiếp qua web.

### 4.2. Danh sách IP được phép

- `10.64.0.108`: máy Tin học
- `10.64.0.62`: máy Ms Trang
- `10.64.0.60`: máy Ms Tư
- `10.64.0.83`: máy Mr Doanh
- `10.64.0.234`: máy ptp tin học

### 4.3. Hành vi truy cập

- Apache ưu tiên chặn request ngay từ đầu; whitelist gồm 5 IP được phép và `local` để máy chủ có thể vào bằng `localhost`.
- Nếu người dùng vào bằng `localhost`, hệ thống PHP ưu tiên đổi sang IP LAN của máy chủ rồi mới kiểm tra whitelist.
- Nếu `REMOTE_ADDR` không nằm trong danh sách được phép, hệ thống trả về `403 Forbidden`.
- Trang bị chặn phải hiển thị IP hiện tại để dễ đối chiếu khi cần cập nhật whitelist.

### 4.4. Cách cập nhật IP

- Để cho phép IP mới: thêm IP vào hàm `khnv_access_allowed_clients()` trong `access_control.php` và thêm IP đó vào file `.htaccess`.
- Để chặn IP đang được phép: xóa IP khỏi hàm `khnv_access_allowed_clients()` trong `access_control.php` và xóa IP đó khỏi file `.htaccess`.

### 4.5. Cách kiểm tra

- Mở `index.php` từ một máy có IP nằm trong whitelist và xác nhận trang vẫn hoạt động bình thường.
- Mở `index.php` từ một máy không nằm trong whitelist và xác nhận nhận trang `403`.
- Thử truy cập trực tiếp một file tĩnh như `logo.png` và xác nhận web server vẫn áp dụng whitelist.
- Truy cập trực tiếp `CHITIEU/index.php` và `CHITIEU/CODE/index.php` để xác nhận không bypass được whitelist.
- Thử truy cập trực tiếp file trong `CHITIEU/INPUT` hoặc `CHITIEU/OUTPUT` và xác nhận web server từ chối.
