# CT_KHNV Home Index Rules

Tài liệu này ghi nhớ các quy ước hiện tại của trang chủ `CT_KHNV/index.php` để tránh sửa nhầm luồng chức năng.

## 1. Mục đích trang chủ

- `index.php` ở thư mục gốc `CT_KHNV` là trang chủ điều hướng nghiệp vụ.
- Trang này không chứa logic import/export Excel hay DOCX.
- Trang chủ chỉ đóng vai trò hiển thị các thẻ chức năng và điều hướng người dùng.

## 2. Chức năng Kế hoạch

- Thẻ có `kicker` là `Kế hoạch` là một chức năng riêng biệt.
- Chức năng `Kế hoạch` hiện tại không liên quan tới thư mục `CHITIEU`.
- Không được gắn link của `Kế hoạch` sang `CHITIEU/index.php` hay `CHITIEU/CODE/index.php`.
- Khi người dùng bấm vào thẻ `Kế hoạch`, hệ thống chỉ hiển thị thông báo:
  - `Đang phát triển chờ`
- Nếu chưa có module riêng cho `Kế hoạch`, phải giữ nguyên hành vi thông báo này.

## 3. Chức năng Chỉ tiêu

- Thẻ `Chỉ tiêu` là chức năng đang hoạt động.
- Thẻ này được phép điều hướng sang module trong thư mục `CHITIEU`.
- Luồng hiện tại đang dùng:
  - `CHITIEU/index.php`
- Khi sửa tiếp, không được làm ảnh hưởng tới luồng import/export hiện có trong `CHITIEU`.

## 4. Quy ước giao diện

- Font giao diện trang chủ đang dùng: `Times New Roman`.
- Trang chủ đang dùng CSS tại `assets/home.css`.
- Header trang chủ đang hiển thị full chiều ngang màn hình.
- Phần dưới header hiện có một vùng hero/showcase full chiều ngang để hiển thị logo hoặc nội dung chức năng.
- Hero trang chủ đang dùng hiệu ứng sáng nhẹ theo màu logo và có slogan:
  - `Thấu hiểu lòng dân, tận tâm, phục vụ`
- Nền hero phải còn sắc hồng nhận diện của logo, không để bị bạc sang gần-trắng làm mất màu tổng thể của trang chủ.
- Cuối trang chủ hiện có footer ngắn để hiển thị dòng credit:
  - `Creat by @TinHoc_DN`
- Cụm logo + tiêu đề ở đầu trang hiện là điểm bấm để quay về trạng thái trang chủ.
- Cụm logo + tiêu đề này không hiển thị gạch chân kiểu link.
- Ưu tiên giữ cụm này ở dạng phần tử điều hướng nhẹ, không để browser áp style native kiểu form-control làm xuất hiện nền trắng trên header.
- Nếu cụm này dùng `button` để reset về trang chủ, phải reset giao diện mặc định của button để không bị nền trắng chèn lên header.
- Không dùng selector quá rộng kiểu `.brand-home-trigger *` để ép style toàn bộ phần con; ưu tiên chỉ target phần chữ để không ảnh hưởng màu/logo.
- Khi dọn hoặc sửa giao diện, ưu tiên giữ bố cục gọn, dễ nhìn và không để khối tiêu đề quá lớn gây vỡ layout.

## 5. Quy ước tải chức năng trong trang chủ

- Khi người dùng bấm vào chức năng đang hoạt động, nội dung phải được tải ngay trong vùng nội dung phía dưới của `CT_KHNV/index.php`.
- Không chuyển toàn bộ trình duyệt sang một trang khác nếu chức năng đó hỗ trợ nhúng trong trang chủ.
- Với module `CHITIEU`, luồng nhúng trong trang chủ phải dùng đúng nguồn:
  - `CHITIEU/index.php`
- Nếu cần chế độ nhúng riêng cho module, phải giữ nguyên nghiệp vụ hiện có và chỉ tinh chỉnh hiển thị để phù hợp khi mở trong trang chủ.
- Khi chức năng được mở trong vùng nhúng, không cần thêm thanh tiêu đề riêng kiểu `Chức năng đang mở`; phần nhúng nên đi thẳng vào nội dung để gọn chiều cao.

## 6. Quy ước hiển thị module CHITIEU khi nhúng

- Khi `CHITIEU` mở trong trang chủ, không hiển thị lại header riêng, topbar riêng hay khối hero cũ của module.
- Trong chế độ nhúng, phần đầu của `CHITIEU` chỉ giữ lại đúng câu:
  - `Nhập file local theo chuẩn CTKHNV_DP/TW*.xlsx`
- Câu trên phải nằm trong khu toolbar/trạng thái của module, không nằm ở một khối riêng phía trên.
- Trong chế độ nhúng, toolbar chỉ nên hiện 2 nút gọn `Import Excel` và `Xuất DOCX`; panel con chỉ mở khi người dùng bấm.
- Module `CHITIEU` không được tự mở panel import/export khi vừa tải trang; mặc định các panel phải đóng và chỉ mở khi người dùng bấm nút tương ứng.
- Trên desktop, cụm `Đọc lại mẫu` và `Lưu cập nhật` nên nằm cùng hàng với 2 nút gọn này để tiết kiệm chiều cao.
- Trên desktop, cụm `Import Excel` và `Xuất DOCX` nên bám sát phần trạng thái bên trái; cụm `TW/DP` và `Đọc lại mẫu/Lưu cập nhật` nên gom về cột phải để phần đầu toolbar gọn hơn.
- Trên desktop, 4 nút `Import Excel`, `Xuất DOCX`, `Đọc lại mẫu`, `Lưu cập nhật` phải nằm ngay ở vùng đầu toolbar, không bị tụt xuống thành một dải thấp hơn gây nhiều khoảng trắng trống phía trên.
- Khi bấm `Trang chủ KHNV` từ `CHITIEU` trong chế độ nhúng, không được mở lồng một trang chủ mới bên trong khung nhúng.
- Vùng bảng của `CHITIEU` khi nhúng phải hiển thị theo layout rộng ở phần dưới của trang chủ, không để mất chức năng, không để khoảng trắng lớn bất thường phía dưới.
- Ưu tiên tăng không gian theo chiều dọc của vùng cuộn bên phải, không kéo lệch bố cục theo chiều ngang.
- Phần nhúng `CHITIEU` hiện đang dùng:
  - `body.embedded .table-wrap { max-height: calc(100vh - 175px); }`
- Phần nhúng `CHITIEU` hiện đang tụt vào hai bên:
  - `5px` mỗi bên
- Khi sửa tiếp, không để nội dung bị cắt ở mép trái làm mất chữ hoặc mất một phần nút/chức năng.

## 7. Quy ước sửa tiếp

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
- Sau khi sửa `index.php`, nên kiểm tra lại cú pháp bằng:
  - `php -l index.php`
  - `php -l CHITIEU\CODE\index.php`
