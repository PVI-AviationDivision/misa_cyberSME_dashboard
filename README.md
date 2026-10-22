# Báo cáo Doanh thu hàng ngày – Bảo hiểm An Tâm Số (BH PVI × MISA)

Dashboard theo dõi doanh phí chương trình **Bảo hiểm An Tâm Số (Cyber Shield SME)** phân phối qua MISA.
Link online: https://pvi-aviationdivision.github.io/misa_cyberSME_dashboard/

## Dùng hằng ngày (≈ 30 giây)

1. Copy file báo cáo raw (`bhkh-YYYY-MM-DD....xlsx`) vào thư mục **`inbox\`**.
2. Nháy đúp **`update.bat`**. Script sẽ:
   - đọc Excel, gộp với các ngày trước, loại đơn test / nội bộ, tạo `data/data.js`;
   - **chép sẵn tin nhắn báo cáo vào clipboard**;
   - tạo **ảnh báo cáo** khổ điện thoại `output\bao-cao-YYYY-MM-DD.png` và mở thư mục chứa ảnh;
   - đẩy lên GitHub và mở dashboard.
3. Vào nhóm Zalo: **Ctrl+V** để dán tin nhắn, kéo thả ảnh từ thư mục `output` vào → gửi.

Nếu lỡ copy thứ khác: bấm nút **“Tin nhắn Zalo”** trên dashboard hoặc mở `data\message.txt`. Nút **“Ảnh báo cáo”** mở lại ảnh dạng trang web; nút **“In / PDF”** để in hoặc lưu PDF.

**Mẫu tin nhắn tự đổi theo ngày:** thứ Hai dùng mẫu *tổng kết tuần* (`message_template_weekly.txt`), ngày 1 hằng tháng dùng mẫu *tổng kết tháng* (`message_template_monthly.txt`), các ngày còn lại dùng `message_template.txt`. Sửa câu chữ trực tiếp trong các file này, giữ nguyên các ô `{...}`; dòng có `{GOI_...}` tự lặp cho từng gói; dòng dự báo tự ẩn khi chưa đủ 3 ngày dữ liệu.

**Chống chạy nhầm:** script hỏi lại (Y/N) khi file đã từng xử lý, file có ngày cũ hơn báo cáo gần nhất, hoặc báo cáo mới không có thay đổi nào. Chọn N thì không cập nhật gì.

**Kiểm tra dữ liệu:** khi có đơn chưa duyệt, phí khác biểu phí, STBH không khớp gói, thiếu file đính kèm hoặc hiệu lực trước ngày cấp, cửa sổ `update.bat` hiện cảnh báo vàng và dashboard liệt kê ở mục “Kiểm tra dữ liệu”. Các cảnh báo này **không** đưa vào tin nhắn gửi sếp.

Có thể bỏ nhiều file một lúc. Nếu cùng một số đơn xuất hiện ở nhiều file, bản ở file mới nhất được giữ lại.

## Cài đặt lần đầu (một lần duy nhất)

1. Cài **Git for Windows**: https://git-scm.com/download/win (để mặc định khi cài).
2. Thư mục dashboard nằm tại `D:\4. Báo cáo Misa`.
   *Máy khác:* `git clone https://github.com/PVI-AviationDivision/misa_cyberSME_dashboard.git`.
3. Nháy đúp `update.bat`. Lần đầu script tự kết nối thư mục với repo; Git mở trình duyệt để đăng nhập GitHub – dùng tài khoản có quyền ghi vào repo.
4. Bật GitHub Pages: repo → **Settings → Pages** → *Source: Deploy from a branch* → Branch **main**, thư mục **/(root)** → Save. Sau 1–2 phút link online hoạt động.

## Quy tắc tính số liệu

| Chỉ tiêu | Cách tính |
|---|---|
| Doanh phí | Tổng `Phí BH` của GCN **cấp mới**, trạng thái “Đã duyệt” |
| SĐBS | Số đơn chứa `SĐBS` hoặc đuôi `/E01…` – hiển thị riêng, **không** cộng doanh phí (đổi bằng `count_endorsement_premium`) |
| Bản ghi test / nội bộ | Tên khách hàng chứa `TEST` hoặc `BẢO HIỂM PVI`, hoặc Ngày CT trước `count_from` (go-live 21/10/2026) – loại khỏi số liệu (xem lại bằng ô “Hiện bản ghi test / nội bộ”) |
| Gói | Xác định theo `Số tiền BH`: 50/100/200/500 triệu → An Tâm Số 50/100/200/500 |
| So sánh | Hôm qua vs ngày trước · 7 ngày qua vs 7 ngày trước đó · tháng này vs cùng kỳ tháng trước |
| Dự báo cuối tháng | Doanh phí tháng ÷ số ngày đã qua × số ngày của tháng (tính từ go-live); chỉ hiện khi có ≥ 3 ngày |
| Hôm qua / Tháng này / Tháng trước | Theo `Ngày CT`; “hôm qua” = ngày báo cáo − 1 (báo cáo sáng 22/10 → số ngày 21/10); tháng này tính đến hết hôm qua |

## Cấu hình – `config.json`

- `mask_customer_names`: `true` = che tên khách hàng (repo đang **public**, ai có link đều xem được). Chỉ đặt `false` nếu repo chuyển sang private.
- `count_from`: ngày bắt đầu tính số liệu (go-live). `exclude_name_patterns`: các cụm từ trong tên khách hàng để loại là test / nội bộ.
- `packages`: danh sách gói / phí.

## Cấu trúc thư mục

```
index.html          dashboard
data/data.js        dữ liệu đã xử lý (tự sinh – đừng sửa tay)
config.json         cấu hình
message_template*.txt mẫu tin nhắn Zalo (ngày / tuần / tháng)
report.html         ảnh báo cáo khổ điện thoại
output/             ảnh báo cáo PNG              (không đẩy lên GitHub)
update.bat          chạy hằng ngày
scripts/build.ps1   đọc Excel (PowerShell, không cần cài Python/Excel)
inbox/              bỏ báo cáo mới vào đây      (không đẩy lên GitHub)
archive/            báo cáo đã xử lý            (không đẩy lên GitHub)
```
