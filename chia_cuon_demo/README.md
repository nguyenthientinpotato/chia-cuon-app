# Chia cuộn màng (demo)

Tìm phương án chia cuộn màng khổ lớn thành các cuộn khổ nhỏ sao cho đủ hàng, ít phế, ít dư nhất.
Bản Python thay cho file VBA `CTCC2.xlsm` cũ.

## Dùng thử

1. Mở link web của app.
2. Thanh bên trái → **Mở kịch bản đã lưu (.json)** → chọn 1 file trong thư mục `kich_ban/`:
   - `1_mot_cuon_NVL_khong_gioi_han.json`: 1 khổ NVL, không giới hạn KL (giống file VBA cũ).
   - `2_hai_lo_NVL_co_stock.json`: 2 lô NVL, có stock lấp biên.
   - `3_nhieu_lo_NVL.json`: 3 lô NVL khác khổ / biên phế, giới hạn KL và số cuộn.
   - `4_ca_lon_nhieu_kho.json`: ca nhiều khổ, số công thức rất lớn.
3. Bấm **Tính phương án** → xem sơ đồ dao, đáp ứng theo đơn, tải **Excel lệnh chia**.
4. Tab **Liệt kê công thức**: danh sách công thức theo biên phế tăng dần (giống bảng FormulaFastCal cũ);
   công thức trong phương án vừa tính được tô màu theo lệnh.

Có thể sửa trực tiếp bảng NVL / đơn hàng / stock rồi bấm tính lại.

## Điểm mới so với VBA

| VBA cũ | Bản này |
|---|---|
| Liệt kê 1 công thức, Planner tự chọn số bộ | Giải cả kế hoạch: chọn công thức, số bộ, số mét/bộ |
| Tối đa 5–6 khổ, dễ sót tổ hợp | Không giới hạn số khổ, có chứng minh tối ưu khi ca vừa |
| 1 khổ NVL | Nhiều lô NVL (khổ, KL, số cuộn, biên phế riêng) |
| Không xét dung sai số mét, nối cuộn, khổ đôi | ±% số mét, nối 2 cuộn, khổ đôi/ba/tư, stock theo ưu tiên |

## Chạy trên máy (không cần link web)

Cài Python 3.12+ (tick "Add Python to PATH"), rồi double-click `run_app.bat`.
