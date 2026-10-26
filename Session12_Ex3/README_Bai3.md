# Bài 3: Chuẩn hóa quy trình giao vận và tra cứu giá cước (FastMove)

## Các file nộp
| File | Nội dung |
|---|---|
| `FastMove_QuyTrinhKho.docx` | Quy trình vận hành kho, Styles (Title, Heading 1/2, Body) đồng bộ màu xanh `#0A3D91` và cam `#F57C00` |
| `FastMove_GiaCuoc.xlsx` | Sheet `VanDon` (500 dòng), `GiaCuoc` (bảng tra), `KPI` |
| `FastMove_BaoCao_KPI.pptx` | 3 slide: KPI bar chart highlight, hướng dẫn shipper, giải pháp tra cứu |

## Phần 1. Mapping 4 tham số VLOOKUP vào mô hình "Nhân viên kho tra hồ sơ"
`=VLOOKUP(lookup_value, table_array, col_index_num, range_lookup)`

| Tham số | Trong mô hình nhân viên kho | Trong file Excel |
|---|---|---|
| `lookup_value` | Số hiệu trên phiếu hàng đang cầm trên tay | `TRIM(A2)` (mã vận đơn) |
| `table_array` | Tủ hồ sơ chứa các phiếu giá, số hiệu nằm ở cột đầu | `GiaCuoc!$A$2:$B$501` |
| `col_index_num` | "Lấy thông tin ở cột số mấy trong phiếu" | `2` (cột giá cước) |
| `range_lookup` | Chỉ lấy đúng phiếu trùng số hiệu, không lấy phiếu gần giống | `FALSE` |

## Phần 2. Giải pháp công nghệ cho FastMove
- **Công thức một lần cho 500 dòng:** `=IFERROR(VLOOKUP(TRIM(A2),GiaCuoc!$A$2:$B$501,2,FALSE),"Không tìm thấy")`.
- **Dấu `$` (Dollar sign):** `$A$2:$B$501` khóa cố định vùng bảng cước, nên kéo công thức xuống 500 dòng vùng tra cứu không bị trượt. Mã vận đơn `A2` không có `$` nên tự đổi theo từng dòng.
- **5 Quy tắc Vàng:** có một dòng header duy nhất, mỗi ô một giá trị, không merge cell, không hàng trống, mỗi cột một loại dữ liệu.
- **KPI:** sheet `KPI` tính tỷ lệ giao đúng hạn bằng `COUNTIFS`. Quận 7 thấp nhất (79%), nên được highlight màu cam trên Bar Chart, các quận khác màu xám.
- **Tiết kiệm thời gian:** thay tra cứu thủ công bằng công thức, loại bỏ khoảng 16 giờ làm việc mỗi ngày (con số từ đề bài).
- Dữ liệu vận đơn, giá cước và KPI là dữ liệu giả định phục vụ bài tập.

## Phần 3. Kiểm thử: lỗi #N/A khi mã vận đơn có ký tự lạ
**Nguyên nhân:** mã như `FM0037#` hoặc `FM0411?` không có trong bảng cước, `VLOOKUP` với `FALSE` trả `#N/A`.

**Cách xử lý:**
1. Bọc `IFERROR`: `=IFERROR(VLOOKUP(TRIM(A38),GiaCuoc!$A$2:$B$501,2,FALSE),"Không tìm thấy")`. Khi có lỗi, ô hiện chữ "Không tìm thấy" thay vì `#N/A`, các hàm như `SUM` vẫn chạy bình thường.
2. Dùng `TRIM` bên trong để loại khoảng trắng thừa.
3. Lọc cột giá cước theo "Không tìm thấy" (hoặc `COUNTIF`) để lấy danh sách mã cần kiểm tra lại. File mẫu có 2 mã như vậy (sheet `KPI`, ô `G3`).
4. Sửa mã gốc, hoặc bổ sung vào bảng cước nếu là đơn mới.

**Lưu ý:** `IFERROR` che mọi loại lỗi, nên chỉ dùng sau khi đã hiểu nguyên nhân. Ký tự `*` và `?` là ký tự đại diện trong tìm kiếm khớp chính xác: mã `FM0205*` vẫn tra ra giá của `FM0205`, còn `FM0411?` thì không. Nên làm sạch ký tự này trong cột mã trước khi tra cứu.
