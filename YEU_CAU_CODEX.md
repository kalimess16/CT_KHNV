# Yeu cau cho Codex

Khi ban nhan cho Codex, hay uu tien dung file nay de viet yeu cau thay vi nhan tin ngan trong khung chat.

## Quy uoc mac dinh bat buoc

- Khi ban bao `doc file yeu cau`, Codex se doc file nay truoc.
- Sau do, Codex phai uu tien doc cac file `.md` lien quan trong du an de lay dung quy tac va bo canh truoc khi sua code.
- Neu dung den trang chu hoac giao dien tong, uu tien doc `HOME_INDEX_RULES.md`.
- Neu dung den luong import/export, DOCX, Excel, hoac module `CHITIEU`, uu tien doc `IMPORT_EXPORT_RULES.md`.
- Neu co yeu cau moi, ban chi can sua noi dung ben duoi roi nhan: "Hay doc `YEU_CAU_CODEX.md` va lam theo."
- Sau khi chay xong yeu cau/tinh nang/dieu kien moi ma chua co trong cac file `.md` thi bo sung vao hoac tao moi file `.md` lien quan de luu tru lai.

## Cach ghi yeu cau

- Moi nhom cong viec viet theo tung thu muc de Codex de theo doi.
- Moi thu muc nen co day du: `Trang thai`, `Muc do uu tien`, `File can`, `Yeu cau`, `Cach kiem tra`, `Ghi chu them`.
- Co the them nhieu muc cung mot thu muc neu co nhieu viec rieng.

## Quy uoc gia tri nen dung

- `Trang thai`: TODO
- `Muc do uu tien`: `cao` | `trung binh` | `thap`

## Mau yeu cau theo tung thu muc

### 1. Thu muc goc `CT_KHNV/`

- Trang thai: chua lam
- Muc do uu tien: cao
- File can:
  - Doc truoc: `YEU_CAU_CODEX.md`
  - Doc them neu lien quan: `HOME_INDEX_RULES.md`
  - File sua/tao: `index.php`, `assets/home.css`
- Yeu cau:
  - Mo ta ro can sua o trang chu, layout, nut bam, khu vuc nhung module, hoac logic dieu huong.
- Cach kiem tra:
  - Vi du: mo `index.php` tren trinh duyet.
  - Vi du: chay `php -l index.php`.
  - Vi du: bam thu cac nut `Chi tieu`, `Ke hoach`, `Trang chu`.
- Ghi chu them:
  - Ghi ro rang buoc nao khong duoc doi.
  - Neu can giu nguyen UI/luong cu thi ghi ro tai day.

### 2. Thu muc `assets/`

- Trang thai: chua lam
- Muc do uu tien: trung binh
- File can:
  - Doc truoc: `YEU_CAU_CODEX.md`
  - Doc them neu lien quan trang chu: `HOME_INDEX_RULES.md`
  - File sua/tao: `assets/home.css` hoac file asset lien quan
- Yeu cau:
  - Mo ta ro can doi CSS, hinh anh, icon, khoang cach, mau sac, responsive...
- Cach kiem tra:
  - Tai lai trang va kiem tra tren desktop/mobile.
  - Kiem tra co vo layout, lech chu, mat nut, tran noi dung hay khong.
- Ghi chu them:
  - Neu chi duoc sua giao dien, ghi ro "khong sua logic".

### 3. Thu muc `CHITIEU/`

- Trang thai: chua lam
- Muc do uu tien: cao
- File can:
  - Doc truoc: `YEU_CAU_CODEX.md`
  - Doc bat buoc neu sua module nay: `IMPORT_EXPORT_RULES.md`
  - Doc them neu module nay dang duoc nhung trong trang chu: `HOME_INDEX_RULES.md`
  - File sua/tao: `CHITIEU/index.php`, `CHITIEU/CODE/index.php`, `CHITIEU/CODE/data.php`, `CHITIEU/CODE/import.php`, `CHITIEU/CODE/export.php`, `CHITIEU/CODE/style.css`, file trong `INPUT/`, `OUTPUT/` neu co noi ro
- Yeu cau:
  - Mo ta ro dang sua import, export, doc workbook, render bang, file mau DOCX, bo loc du lieu, hoac giao dien module.
- Cach kiem tra:
  - Vi du: chay `php -l CHITIEU\\CODE\\index.php`
  - Vi du: import file `CTKHNV*.xlsx`
  - Vi du: bam `Doc lai mau`
  - Vi du: export DOCX va kiem tra placeholder `{{...}}` da duoc thay het
  - Vi du: kiem tra so dong/so nhom cot dung theo workbook
- Ghi chu them:
  - Neu co file mau dang active thi ghi ro ten file.
  - Neu co quy tac "khong duoc sua" ve filter, placeholder, bo cuc thi ghi ro tai day.

### 4. Thu muc `kehoach/`

- Trang thai: chua lam
- Muc do uu tien: thap
- File can:
  - Doc truoc: `YEU_CAU_CODEX.md`
  - Doc them file `.md` lien quan neu sau nay du an bo sung
  - File sua/tao: ghi ro file cu the trong `kehoach/`
- Yeu cau:
  - Mo ta ro module can tao moi, trang tam, UI thong bao, hay logic rieng.
- Cach kiem tra:
  - Mo ta cach bam thu, mo duong dan, hoac test nghiep vu can xac nhan.
- Ghi chu them:
  - Neu tam thoi chi hien `Dang phat trien cho` thi ghi ro de Codex khong noi nham sang `CHITIEU`.

## Noi dung yeu cau

### Thu muc `CHITIEU/`

- Trang thai: TODO
- Muc do uu tien: cao
- File can:
  - Doc truoc: `YEU_CAU_CODEX.md`
  - Doc bat buoc neu sua module nay: `IMPORT_EXPORT_RULES.md`
  - Doc them neu module nay dang duoc nhung trong trang chu: `HOME_INDEX_RULES.md`
  - File sua/tao: `CHITIEU/index.php`, `CHITIEU/CODE/index.php`, `CHITIEU/CODE/data.php`, `CHITIEU/CODE/import.php`, `CHITIEU/CODE/export.php`, `CHITIEU/CODE/style.css`, file trong `INPUT/`, `OUTPUT/` neu co noi ro
- Yeu cau:
  - trên index khi tôi thay đổi cột ` điều chỉnh tăng trưởng ` thì cột ` chỉ tiêu kế hoạch năm 2026` = ` kế hoạch năm 2026 đã giao` + ` điều chỉnh tăng trưởng `.
   Cach kiem tra:
  - vào xuất docs tải về xem
- Ghi chu them:
  - Neu co file mau dang active thi ghi ro ten file.
  - Neu co quy tac "khong duoc sua" ve filter, placeholder, bo cuc thi ghi ro tai day.