# SCAF → Bộ chứng từ xuất khẩu

Web tĩnh (chạy trên GitHub Pages) để chuyển **file Data xuất từ SCAF** thành **bộ chứng từ Excel** theo form mẫu, gồm các sheet:

| Sheet | Nội dung |
|---|---|
| `<Customs code>` (vd `4AC089-977`) | Dán dữ liệu SCAF: Order No, Reference, Color, Size, CTN From, CTN To, Number of Box, Real QTY + bảng **tổng hợp theo Customs code + Size** (cột L:O) |
| `PLIST` | Packing list, lấy từ sheet Customs code (công thức liên kết) |
| `CI` | Commercial Invoice: số INV theo SCAF, các dòng item lấy từ bảng tổng hợp của sheet Customs code (không có bảng tổng hợp theo HS code) |
| `INV-3RD` | Invoice bên thứ 3 (buyer phát hành), lấy từ `CI`, **có sẵn dấu/chữ ký buyer** (dùng cho hồ sơ CO). Đổi/bỏ dấu ngay trên web |
| `INV - delta` | Lấy thông tin từ sheet `CI` bằng công thức |

Các sheet liên kết với nhau bằng công thức như form gốc, nên sửa trong Excel (giá, số lượng…) sẽ tự cập nhật.

## Cách dùng
1. Mở trang web, chọn file `.xlsx` xuất từ SCAF (báo cáo Packing List).
2. Nhập **Order No** cho từng lô (mặc định lấy theo SCAF, sửa lại cho đúng), **Customs code** cho từng cặp Reference + Màu (vd `4AC089-977`), ORDER NO (style, vd `AOFYDDQ11`), HS code, đơn giá. Lần sau web tự nhớ.
3. Kiểm tra thông tin chứng từ (consignee, tàu, CBM/NW/GW, container…).
4. Chọn quy ước size: `90(S)` → **S** (mặc định), hoặc **90**, hoặc giữ nguyên.
5. Bấm **Xuất file Excel**.

Bảng mã có thể **Lưu (.json)** và **Nạp** để chia sẻ cho người khác trong phòng.

Toàn bộ xử lý chạy trên trình duyệt, dữ liệu không gửi đi đâu.

## Đưa lên GitHub Pages
```bash
git init && git add . && git commit -m "SCAF docs generator"
git branch -M main
git remote add origin https://github.com/<tai-khoan>/scaf-docs.git
git push -u origin main
```
Trên GitHub: **Settings → Pages → Build and deployment → Source: Deploy from a branch → Branch: `main` / `(root)` → Save**.
Sau 1–2 phút, web chạy tại `https://<tai-khoan>.github.io/scaf-docs/`.

> Nếu repo để **Private**, GitHub Pages cần gói GitHub Pro/Team/Enterprise.

## Cấu trúc
```
index.html, style.css   giao diện
app.js                  xử lý giao diện, nhớ bảng mã (localStorage)
core.js                 đọc file SCAF + tạo workbook (dùng chung trình duyệt & Node)
vendor/exceljs.min.js   thư viện ExcelJS 4.4.0
tests/node-test.js      chạy thử bằng Node: npm i && npm test
sample/                 file data mẫu
```
