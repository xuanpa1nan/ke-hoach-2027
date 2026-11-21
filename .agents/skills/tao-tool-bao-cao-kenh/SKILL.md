---
name: tao-tool-bao-cao-kenh
description: Phát triển hoặc sửa tool báo cáo quản trị kênh bán hàng di động từ IMS73, BRIS và SMCS, tính HQ/BSC và xuất báo cáo tuần Excel và Word. Dùng khi thêm nguồn, điều chỉnh công thức, xây đầu ra hoặc đối soát một kỳ báo cáo.
---

# Tạo tool báo cáo quản trị kênh

## Chọn tài liệu theo việc cần làm

| Công việc | Đọc trước |
| --- | --- |
| Thu thập, nhận diện tệp, ánh xạ dữ liệu và tính chỉ tiêu | [Quy trình và công thức](../../../docs/quan-tri-kenh-ban-hang.md) |
| Xuất workbook HQ/BSC, điền báo cáo tuần và kiểm tra bản in | [Đặc tả đầu ra và đối soát](../../../docs/dau-ra-bao-cao-tuan.md) |
| Tái sử dụng phần tính đã có | [Tool HQ](../../../scripts/tinh_hq.py) và [hướng dẫn chạy](../../../README.md) |

## Khi phát triển hoặc sửa tool

1. Chốt kỳ, nguồn bắt buộc và đơn vị tính cho từng chỉ tiêu. Kiểm tra tiêu đề và khóa trong tệp thật. Chọn báo cáo theo dữ liệu trong tệp, vì tên tệp và ngày xuất có thể khác kỳ nghiệp vụ.
2. Dùng IMS73 trên máy có IP VNPT. Sau khi tải số liệu, chờ bảng ổn định rồi xuất. Nhận diện BRIS theo tiền tố `hqkd_kenh_ct_ddtt` và cấu trúc cột. Dùng danh mục Điểm Ủy quyền của SMCS để tra mã `NAN…` sang eLoad. Giữ khóa gốc, dòng nguồn và mã đơn vị giao dịch để đối soát.
3. Chuẩn hóa theo khóa và tổng hợp một lần theo **kỳ, đơn vị quản lý, kênh, điểm**. VinaPro dùng HRM, VinaKey1 dùng eLoad. Giữ các điểm có sản lượng bằng 0 trong mẫu số. Lưu riêng doanh thu PTM, tái sử dụng, doanh thu đánh giá, số điểm PSSL và số thuê bao PTM. Dùng đơn vị giao dịch hoặc đơn vị quản lý theo định nghĩa của chỉ tiêu.
4. Đối chiếu theo **điểm, đơn vị và toàn tỉnh**. Kiểm tra tổng, số dòng, khóa thiếu hoặc trùng. Thử bằng dữ liệu giả lập khi kỳ mẫu không có tình huống như mã `NAN…` hoặc khóa khớp nhiều điểm. Ghi rõ nguồn thay thế và khác biệt với công thức cũ.
5. Xuất ba đầu ra theo phạm vi được yêu cầu. Ô thiếu đầu vào ghi **Thiếu dữ liệu**. Hạng mục hoãn ghi **Bổ sung sau**. Chỉ áp dụng tỷ lệ, kế hoạch hoặc quy tắc gán sản lượng chưa rõ sau khi xác nhận với người dùng.

Repo là công khai. Giữ bản tải, workbook thật, kết quả chạy thử và dữ liệu đối soát trong `local/`. Chỉ đưa mã nguồn, quy tắc nghiệp vụ và ví dụ giả lập vào Git.
