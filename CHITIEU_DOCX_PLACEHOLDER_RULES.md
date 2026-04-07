# CHITIEU DOCX Placeholder Rules

Tai lieu nay ghi lai luu y da chot cho luong xuat DOCX cua module `CHITIEU`.

## 1. Placeholder bi tach run

- Placeholder trong `word/document.xml` co the bi tach qua nhieu `w:r` hoac `w:t`.
- Logic thay placeholder phai xu ly duoc ca khi root dang render la toan bo `body`, `table`, hoac mot `paragraph` clone rieng.
- Khong duoc de sot `{{phong_giao_dich}}`, `{{PHONG_GIAO_DICH}}`, hoac placeholder tuong tu chi vi template tach chuoi thanh nhieu run.

## 2. Mau TT

- Mau `TT` dang active la `CHITIEU/OUTPUT/Dieu_chinh_chi_tieu.docx`.
- `{{phong_giao_dich}}` phai lay dung ten PGD cua tung block dang render.
- Sau khi render xong mot PGD, luong xuat phai tiep tuc render block PGD tiep theo trong cung file DOCX.

## 3. Kiem tra

- Sau khi export `TT`, khong duoc con placeholder dang `{{...}}` trong file DOCX sinh ra.
- Truong `{{phong_giao_dich}}` phai hien ten PGD that trong tieu de tung block.
