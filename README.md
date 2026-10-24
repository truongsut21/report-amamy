# Amamy Dashboard

Dashboard HTML/CSS/JS + Chart.js, đọc trực tiếp file `data/amamy.xlsx` (xuất từ Google Sheet "Tổng hợp đơn hàng Amamy").
Các tab và biểu đồ tương ứng với Apps Script `scrip.js`.

## Cấu trúc

```
index.html         Trang chính
css/style.css      Giao diện (tự đổi sáng/tối)
js/worker.js       Đọc xlsx trong Web Worker (SheetJS) → tính số liệu
js/charts.js       Helper Chart.js + định dạng số VNĐ
js/app.js          Tabs, KPI, biểu đồ, bảng
data/amamy.xlsx    Dữ liệu
```

## Chạy thử trên máy

Trình duyệt không cho đọc file khi mở `index.html` trực tiếp (file://), nên cần một server tĩnh:

```bash
python3 -m http.server 8000
# mở http://localhost:8000
```

## Đưa lên GitHub Pages

```bash
git init && git add . && git commit -m "Amamy dashboard"
git branch -M main
git remote add origin https://github.com/<tài-khoản>/<repo>.git
git push -u origin main
```

Trên GitHub: **Settings → Pages → Source: Deploy from a branch → main / (root)**. Sau ~1 phút trang có ở `https://<tài-khoản>.github.io/<repo>/`.

## Cập nhật dữ liệu

1. Google Sheet → Tệp → Tải xuống → Microsoft Excel (.xlsx)
2. Đổi tên thành `amamy.xlsx`, chép đè vào `data/`
3. `git add data/amamy.xlsx && git commit -m "Cập nhật dữ liệu" && git push`

Tên các tab trong file phải giữ nguyên (Khách hàng, Chi tiết đơn, Theo line, …). Các tab `📊 …` không cần thiết.

Ngoài ra có thể bấm **Mở file xlsx** hoặc kéo thả file vào trang để xem một file khác — file chỉ được đọc trong trình duyệt, không tải lên đâu cả.

## Tốc độ

- Parse xlsx chạy trong Web Worker → giao diện không bị đứng.
- Kết quả đã tính được lưu vào IndexedDB của trình duyệt; lần sau nếu file không đổi thì trang hiện ngay (~10ms), không parse lại.
- Bảng dữ liệu gốc mỗi tab phân trang 50 dòng, có tìm kiếm / sắp xếp / tải CSV.
- Biểu đồ chỉ được vẽ khi cuộn tới (lazy).

## ⚠️ Dữ liệu cá nhân

File xlsx có tên, SĐT, email khách hàng. Repo GitHub **public** + GitHub Pages = ai có link cũng tải được `data/amamy.xlsx`.
Nên dùng repo **private** (GitHub Pages cho repo private cần gói trả phí), hoặc không commit file data và dùng nút **Mở file xlsx** để xem.
