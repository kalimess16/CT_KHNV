# CHITIEU Clear And DOCX Rules

Tai lieu nay ghi lai 3 quy tac moi vua them cho module `CHITIEU`.

## 1. DOCX row layout

- Cac dong du lieu clone dong vao bang DOCX khong duoc giu `w:trHeight` co dinh neu no lam dong bi du chieu cao.
- Dong noi dung ngan phai tu co theo text; dong dai van duoc phep cao len theo noi dung thuc te.

## 2. Placeholder PGD

- Khi render mau `TT`, `{{phong_giao_dich}}` va `{{PHONG_GIAO_DICH}}` phai dua ten PGD ve chu hoa.
- Khong duoc de sot placeholder trong `word/document.xml` sau khi xuat.

## 3. Xoa so lieu workbook

- Man hinh `CHITIEU/CODE/index.php` co nut `Xoa so lieu` cho workbook dang xem (`TW` hoac `DP`).
- Khi xoa, tat ca o so lieu tu vung du lieu tro xuong phai ve rong.
- Cong thuc Excel phai duoc giu lai; neu toan bo o nguon cua cong thuc deu rong thi gia tri hien thi cua o cong thuc cung phai rong.

## 4. Tong am o cong thuc tong hop

- Cac o tong hop dung `SUM(...)` hoac cong tru nhieu o trong module `CHITIEU` neu tinh ra am thi phai giu nguyen tong am, khong duoc ep ve `0`.
- Khi cong thuc tong hop co chen thanh phan loi cu nhu `#REF!` nhung van con cac hang hop le khac, uu tien tinh tren phan hop le con lai de khong lam mat tong am.
