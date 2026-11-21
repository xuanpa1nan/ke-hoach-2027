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

Repo hiện ở giai đoạn mô tả quy trình. Dữ liệu nội bộ, workbook chứa dữ liệu thật và thông tin đăng nhập được lưu ngoài Git.
