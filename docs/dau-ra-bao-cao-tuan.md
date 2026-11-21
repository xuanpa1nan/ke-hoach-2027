# Đầu ra và đối soát báo cáo tuần

Phần quản trị kênh bán hàng thuộc mảng **Di động** trong báo cáo tuần. Bộ báo cáo còn có BRCĐ, MyTV, CNTT và các dịch vụ khác. Mẫu tuần 47 dùng để xác định bố cục, tên chỉ tiêu và cách trình bày. Số liệu đưa vào báo cáo phải lấy theo kỳ đang tính.

## Đầu ra và nơi sử dụng số liệu

| Tệp | Vai trò và vùng cần điền |
| --- | --- |
| **Workbook HQ/BSC** | `HQ` tổng hợp theo điểm, kênh và đơn vị. `BSC` tính theo kế hoạch kỳ và các chỉ tiêu I.8.1, I.8.2, I.8.3. Ghi rõ nguồn thay thế và chỉ tiêu chưa đủ điều kiện tính. |
| **Báo cáo SXKD tuần Excel** | `KQTH Kênh PRO, KEY1` trình bày chi tiết theo đơn vị. VinaPro dùng `J` doanh thu đánh giá, `L` điểm PSSL, `M` thuê bao PTM, `O` điểm đạt điều kiện chi lương. VinaKey1 dùng `S`, `U`, `V`, `X` tương ứng. Trong `CTKD DT DD và KBH`, vùng `O:Q` của CT2.1 dùng doanh thu PTM theo **đơn vị giao dịch**, quy đổi triệu đồng. Xác nhận định nghĩa điểm đạt chuẩn và kế hoạch trước khi điền các cột tỷ lệ hoặc xếp hạng. |
| **Báo cáo Điều hành SXKD tuần Word** | Phần **Hiệu quả kênh ngoài** nêu doanh thu, sản lượng, điểm PSSL và tình hình theo đơn vị. Chỉ viết mức tăng giảm hoặc so với kế hoạch khi có số cùng định nghĩa, cùng kỳ đối chiếu và có kế hoạch tương ứng. |

## Quy tắc trình bày

- **Điểm PSSL** là số điểm bán có sản lượng phát sinh. **Sản lượng PTM** là số thuê bao. **Điểm phát sinh doanh thu** và **điểm đạt điều kiện chi lương** là hai chỉ tiêu riêng. Tổng điểm PSSL hiện tại chỉ gồm VinaPro và VinaKey1, ghi rõ phạm vi ở tiêu đề. CTV liên kết sẽ bổ sung sau.
- Giữ riêng **doanh thu PTM** và **doanh thu đánh giá**. HQ theo điểm dùng đơn vị quản lý, còn chỉ tiêu CT2.1 lấy đơn vị giao dịch khi theo cấu trúc mẫu. Chỉ dùng tỷ lệ tái sử dụng đã xác nhận cho kênh và kỳ báo cáo.
- Giữ bảng chi tiết để truy vết. Bản in tách theo **doanh thu**, **sản lượng** và **BSC**, có tiêu đề kỳ, đơn vị đo, dòng **Toàn tỉnh**, cột số căn phải và tỷ lệ định dạng phần trăm. Ưu tiên A4 ngang với mỗi nhóm vừa một trang. Kiểm tra xem trước để phát hiện cột bị cắt hoặc chữ quá nhỏ.
- Giá trị **0** là kết quả đã tính. Dùng **Thiếu dữ liệu** khi chưa có đầu vào và **Bổ sung sau** cho hạng mục hoãn.

## Đối soát trước khi xuất

1. Kiểm tra tổng doanh thu trước và sau chuẩn hóa tài khoản. Đối chiếu số dòng theo kỳ, tách mã không khớp hoặc trùng, và kiểm tra số điểm trong danh mục gồm cả điểm chưa phát sinh.
2. So kết quả theo điểm và đơn vị, rồi mới so tổng toàn tỉnh. Trong bản mẫu đã kiểm tra, tổng VinaKey1 có thể khớp dù một số công thức đọc eLoad ở dòng khác. Phần TD49 cũng có thể lệch vì vùng công thức cũ dừng trước dòng cuối của dữ liệu.
3. Kiểm tra số trong Word khớp chỉ tiêu và đơn vị đo trong Excel. Giữ số kỳ trước, kế hoạch và tỷ lệ chấm theo từng kỳ để giải thích chênh lệch.
4. Chỉ công bố số đã đủ nguồn. I.8.1 phụ thuộc đơn hàng CTV, danh sách quản lý và kế hoạch. I.8.2 dùng số điểm có phát sinh và danh mục đầy đủ. I.8.3 phụ thuộc doanh thu đánh giá và kế hoạch. Nếu dùng giá trị lưu sẵn trong workbook để thử, ghi rõ đó là số tham chiếu.

Quy tắc lấy tệp, cột nguồn và các điểm chưa chốt nằm trong [quy trình quản trị kênh bán hàng](quan-tri-kenh-ban-hang.md).
