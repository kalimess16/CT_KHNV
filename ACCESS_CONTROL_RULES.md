# CT_KHNV Access Control Rules

Tai lieu nay ghi lai quy tac whitelist IP dang ap dung cho du an `CT_KHNV`.

## 1. Pham vi ap dung

- Thu muc goc `CT_KHNV/` co `.htaccess` de chan web request tu IP khong nam trong whitelist.
- Trang chu goc `index.php` phai chay kiem tra IP truoc khi render giao dien.
- Module `CHITIEU` khi vao qua `CHITIEU/index.php` hoac truy cap thang `CHITIEU/CODE/index.php` deu phai bi rang buoc boi cung whitelist.
- Khong duoc dua logic whitelist vao CSS hay JavaScript giao dien.
- Cac thu muc du lieu `CHITIEU/INPUT` va `CHITIEU/OUTPUT` khong duoc phep tai truc tiep qua web.

## 2. Danh sach IP duoc phep

- `10.64.0.108`: may Tin hoc
- `10.64.0.62`: may Ms Trang
- `10.64.0.60`: may Ms Tu
- `10.64.0.83`: may Mr Doanh
- `10.64.0.234`: may ptp tin hoc

## 3. Hanh vi truy cap

- Apache uu tien chan request ngay tu dau; whitelist gom 5 IP duoc phep va `local` de may chu co the vao bang `localhost`.
- Neu nguoi dung vao bang `localhost`, he thong PHP uu tien doi sang IP LAN cua may chu roi moi kiem tra whitelist.
- Neu `REMOTE_ADDR` khong nam trong danh sach duoc phep, he thong tra ve `403 Forbidden`.
- Trang bi chan phai hien IP hien tai de de doi chieu khi can cap nhat whitelist.

## 4. Cach cap nhat IP

- De cho phep IP moi: them IP vao ham `khnv_access_allowed_clients()` trong `access_control.php` va them IP do vao file `.htaccess`.
- De chan IP dang duoc phep: xoa IP khoi ham `khnv_access_allowed_clients()` trong `access_control.php` va xoa IP do khoi file `.htaccess`.

## 5. Cach kiem tra

- Mo `index.php` tu mot may co IP nam trong whitelist va xac nhan trang van hoat dong binh thuong.
- Mo `index.php` tu mot may khong nam trong whitelist va xac nhan nhan trang `403`.
- Thu truy cap truc tiep mot file tinh nhu `logo.png` va xac nhan web server van ap dung whitelist.
- Truy cap truc tiep `CHITIEU/index.php` va `CHITIEU/CODE/index.php` de xac nhan khong bypass duoc whitelist.
- Thu truy cap truc tiep file trong `CHITIEU/INPUT` hoac `CHITIEU/OUTPUT` va xac nhan web server tu choi.
