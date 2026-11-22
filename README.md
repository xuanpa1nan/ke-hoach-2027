# Kế hoạch 2027

Dự án tự động hóa việc tải dữ liệu từ cổng nội bộ của công ty và đưa dữ liệu vào tệp Excel tính toán hiện có. Mục tiêu là giảm thao tác tải tệp, giải nén và sao chép dữ liệu thủ công giữa các workbook.

## Luồng dự kiến

1. Mở trình duyệt và truy cập cổng nội bộ bằng tài khoản được cấp quyền. Người vận hành xác nhận nếu hệ thống yêu cầu đăng nhập lại hoặc xác thực nhiều yếu tố.
2. Chọn báo cáo và kỳ dữ liệu, sau đó tải các tệp Excel hoặc ZIP về máy.
3. Giải nén các tệp ZIP vào thư mục xử lý và kiểm tra tệp đầu vào, tên sheet và cột bắt buộc.
4. Tạo bản sao của workbook tính toán. Đưa dữ liệu thô vào đúng sheet và vùng dữ liệu theo cấu hình, giữ nguyên công thức và định dạng của mẫu.
5. Tính lại workbook, đối chiếu số dòng và kiểm tra lỗi công thức trước khi lưu kết quả theo kỳ. Nếu dữ liệu không khớp cấu hình, ghi rõ tệp và nguyên nhân để người vận hành kiểm tra.

## Dữ liệu dự án

| Thành phần | Vai trò |
| --- | --- |
| Tệp Excel hoặc ZIP tải từ cổng nội bộ | Dữ liệu thô theo kỳ |
| Workbook tính toán | Mẫu chứa các sheet đích, công thức và định dạng |
| Cấu hình ánh xạ | Quy định tệp và sheet nguồn, cột cần lấy, sheet đích và ô bắt đầu |
| Workbook kết quả | Bản sao của mẫu đã nhận dữ liệu và được tính lại |

## Thông tin cần chốt

- Các báo cáo cần tải, cách chọn kỳ và quy tắc đặt tên tệp.
- Cấu trúc tệp ZIP và workbook nguồn, cùng quy tắc xử lý khi thiếu hoặc trùng dữ liệu.
- Ánh xạ từng nguồn vào workbook tính toán và cách cập nhật dữ liệu cũ khi chạy lại cùng kỳ.
- Cách tính lại công thức, kiểm tra kết quả và nơi lưu đầu ra.

Dữ liệu nội bộ, workbook chứa dữ liệu thật và thông tin đăng nhập được lưu ngoài Git.

## Chạy thử HQ

Tool tính HQ cho VinaPro và VinaKey1 từ tệp IMS73 và BRIS đã xuất. Chuẩn bị danh mục SMCS dạng Excel với các cột `Mã điểm bán`, `Số eLoad`, `Loại kênh`, `Đối tượng`, `Mã HRM` và `Mã đơn vị`. Bổ sung HRM của VinaPro và mã đơn vị quản lý từ danh mục nhân viên hoặc đơn vị nếu tệp SMCS thiếu các trường này. Có thể chỉ định sheet danh mục bằng `--smcs-sheet`.

```powershell
python -m pip install -r requirements.txt
python scripts/tinh_hq.py --ims local/RPVNP0173.xls --bris local/bris.xlsx --smcs local/smcs.xlsx --bris-from 2026-11-01 --bris-to 2026-11-03 --output local/hq
```

Lệnh dùng toàn bộ tệp doanh thu IMS theo kỳ đã xuất. Hai tham số ngày ở ví dụ lọc ngày kích hoạt BRIS. Để tính tái sử dụng, truyền thêm `--td49` cho workbook có sheet `ALL_TD49` và `--sim-tsd` cho workbook có sheet `Sim TSD`. Khi thiếu nguồn tái sử dụng của kênh, cột `reuse_revenue` để trống và doanh thu HQ của kênh chỉ gồm PTM.

Kết quả gồm `hq.csv` theo đơn vị và kênh, `points.csv` theo điểm, `exceptions.csv` theo dòng nguồn và `run.json` mô tả phạm vi chạy. Dùng `--compare` với workbook tuần 46 để xuất thêm bảng chênh lệch. Nếu thiếu danh mục SMCS trong bộ thử tuần 46, `--week46-catalog` dùng danh mục điểm từ workbook mẫu thay cho `--smcs`. Tham số `--diagnose-reference-bris` tạo một phép đối chiếu riêng dùng khóa BRIS đã lưu trong workbook mẫu. Kết quả chính vẫn dùng BRIS thô.

Xem [quy trình quản trị kênh bán hàng](docs/quan-tri-kenh-ban-hang.md) để biết nguồn dữ liệu, công thức và các điểm cần chốt.

## Xuất bộ báo cáo thử

Sau khi chạy HQ, chuẩn bị dữ liệu cho ba đầu ra và xuất Excel/Word. Bộ thử lấy danh mục CTV, nhân viên và kế hoạch BSC từ workbook tuần 46. Mẫu Excel tuần 47 cung cấp tên sheet và tiêu đề. Các chỉ tiêu ngoài phạm vi tính được đánh dấu thiếu trong báo cáo.

```powershell
python scripts/du_lieu_bao_cao_tuan.py --hq-dir local/hq --reference local/mau-tuan46.xlsx --ims local/RPVNP0173.xls --excel-template local/mau-bao-cao.xlsx --week 46 --output local/du-lieu-bao-cao.json
node scripts/xuat_bao_cao_tuan.mjs local/du-lieu-bao-cao.json local/bao-cao-tuan46
& ./scripts/dat-kho-in.ps1 "local/bao-cao-tuan46/HQ và BSC - Tuần 46.xlsx" "local/bao-cao-tuan46/BC SXKD - Tuần 46.xlsx"
```

Bước xuất dùng Node.js với `@oai/artifact-tool` và `docx`. Trong môi trường Codex, đặt `REPORT_NODE_MODULES` thành đường dẫn thư mục thư viện Node.js do trình nạp phụ thuộc cung cấp. Chạy bước cuối trong PowerShell 7 với Microsoft Excel trên Windows để đặt các sheet `In doanh thu`, `In sản lượng` và `In BSC` ở khổ A4 ngang, mỗi sheet vừa một trang.

## Chạy giao diện trên máy

Sau khi cài thư viện Python và thiết lập Node.js như trên, chạy `python scripts/giao_dien_hq.py` rồi mở `http://127.0.0.1:8765`. Nếu Node.js chưa có trên PATH, thêm `--node <đường dẫn node.exe>`.

Chọn IMS73, BRIS và danh mục SMCS. Để tính BSC, chọn thêm workbook tham chiếu đúng kỳ có bố cục `HQ`, `BSC`, `TH` và các bảng CTV giống mẫu. Chọn mẫu Excel tuần để xuất báo cáo. Các kết quả được tách theo lần chạy trong `local/giao-dien-hq/`.

Để hiện nút chạy bộ mẫu tuần 46, truyền `--demo-dir <thư mục dữ liệu>`, `--demo-reference <workbook tuần 46>` và `--demo-template <mẫu Excel tuần>`. Chế độ này dùng danh mục điểm trong workbook mẫu khi chưa có file SMCS.
