# Báo cáo Doanh thu hàng ngày – Bảo hiểm An Tâm Số (BH PVI × MISA)

Dashboard theo dõi doanh phí chương trình **Bảo hiểm An Tâm Số (Cyber Shield SME)** phân phối qua MISA.
Link online: https://pvi-aviationdivision.github.io/misa_cyberSME_dashboard/

## Dùng hằng ngày (≈ 30 giây)

1. Copy file báo cáo raw (`bhkh-YYYY-MM-DD....xlsx`) vào thư mục **`inbox\`**.
2. Nháy đúp **`update.bat`**.
   - Script đọc Excel → gộp với các ngày trước → tạo `data/data.js` → **chép sẵn tin nhắn báo cáo vào clipboard** → đẩy lên GitHub → mở dashboard.
   - File đã xử lý tự chuyển sang `archive\` (chỉ lưu trên máy, **không** đẩy lên GitHub).
3. Mở nhóm Zalo → **Ctrl+V** → gửi. (Nếu lỡ copy thứ khác: bấm nút **“Tin nhắn Zalo”** trên dashboard, hoặc mở `data\message.txt`.)

Tin nhắn gồm: ngày hôm qua, tháng này, tháng trước, lũy kế chương trình và số GCN theo từng gói. Muốn sửa câu chữ → sửa file **`message_template.txt`** (giữ nguyên các ô `{...}`; dòng nào có `{GOI_...}` sẽ tự lặp cho từng gói).

Có thể bỏ nhiều file một lúc. Nếu cùng một số đơn xuất hiện ở nhiều file, bản ở file mới nhất (theo ngày trong tên file, rồi hậu tố `_2`, `_3`…) được giữ lại.

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
| Bản ghi test | Khách hàng có tên chứa `TEST` – loại khỏi số liệu (bật lại bằng ô “Hiện bản ghi TEST”) |
| Gói | Xác định theo `Số tiền BH`: 50/100/200/500 triệu → An Tâm Số 50/100/200/500 |
| Hôm qua / Tháng này / Tháng trước | Theo `Ngày CT`; “hôm qua” = ngày báo cáo − 1 (báo cáo sáng 22/10 → số ngày 21/10); tháng này tính đến hết hôm qua |

## Cấu hình – `config.json`

- `mask_customer_names`: `true` = che tên khách hàng (repo đang **public**, ai có link đều xem được). Chỉ đặt `false` nếu repo chuyển sang private.
- `monthly_target_premium`: điền kế hoạch doanh phí tháng (đồng) để hiện thanh tiến độ và % trong tin nhắn; `0` = ẩn.
- `packages`: danh sách gói / phí.

## Cấu trúc thư mục

```
index.html          dashboard
data/data.js        dữ liệu đã xử lý (tự sinh – đừng sửa tay)
config.json         cấu hình
message_template.txt mẫu tin nhắn Zalo
update.bat          chạy hằng ngày
scripts/build.ps1   đọc Excel (PowerShell, không cần cài Python/Excel)
inbox/              bỏ báo cáo mới vào đây      (không đẩy lên GitHub)
archive/            báo cáo đã xử lý            (không đẩy lên GitHub)
```
