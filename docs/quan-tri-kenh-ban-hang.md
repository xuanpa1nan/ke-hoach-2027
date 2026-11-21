# Quy trình quản trị kênh bán hàng di động

Quy trình này chuẩn bị số liệu cho phần **Di động → Quản trị kênh bán hàng** của báo cáo kết quả kinh doanh tuần. `HQ` tổng hợp hiệu quả theo điểm bán, kênh và đơn vị. `BSC` dùng kết quả đó cùng số liệu cộng tác viên và kế hoạch để chấm điểm.

Tài liệu ghi lại quy trình vận hành và đặc tả tính toán từ workbook mẫu tuần 46. [Tool tính HQ](../scripts/tinh_hq.py) xử lý các tệp đã xuất cho VinaPro và VinaKey1. Các bước tải từ hệ thống và tính BSC nằm ngoài phạm vi của bản thử.

## 1. Nguồn dữ liệu và cách lấy

| Nguồn | Cách lấy và nhận diện | Trường dùng để tính |
| --- | --- | --- |
| **IMS73** | Dùng máy có IP VNPT và đăng nhập bằng tài khoản được cấp. Vào **VNP → Báo cáo → 73. Doanh thu bán hàng di động**. Chọn kỳ, bấm **Tải dữ liệu**, chờ bảng hiển thị ổn định rồi bấm **Xuất**. Tệp chi tiết là `RPVNP0173.xls`. | Sheet `temp1`, tiêu đề dòng 1, dữ liệu từ dòng 2. `B` đơn vị giao dịch, `D` số máy, `F` giá gói, `K` loại thuê bao, `L` tài khoản đăng ký, `O` đối tượng. Các cột này lần lượt tương ứng `E`, `G`, `I`, `N`, `O`, `R` của `IMS73_CT`. |
| **BRIS / P01** | Mở báo cáo chi tiết phát triển thuê bao trả trước, chọn đơn vị và kỳ, lấy dữ liệu rồi xuất Excel. Nhận diện theo mẫu `hqkd_kenh_ct_ddtt*.xlsx` và kiểm tra tiêu đề cột. | Tệp mẫu có dữ liệu trong `Sheet1`, tiêu đề dòng 4 và chi tiết từ dòng 5. `F` là số thuê bao, `J` là ngày kích hoạt, `K` là User/eLoad đăng ký thông tin thuê bao, tương ứng `P01_SL PTM!L` trước các bước điều chỉnh của workbook. |
| **SMCS** | Vào **Báo cáo → Danh sách kênh bán hàng** và xuất danh mục. Đối tượng **Điểm Ủy quyền** tương ứng **VinaKey1** trong nghiệp vụ này. | Mã điểm bán, số eLoad, loại kênh, đối tượng. Dùng cùng danh mục đơn vị và nhân viên để xác định đơn vị quản lý và chủ thể. |
| **Nguồn bổ sung** | Đơn hàng CTV, danh sách CTV và nhân viên quản lý, kế hoạch BSC theo kỳ, bảng `ALL_TD49` và danh sách `Sim TSD`. | Cần để tái lập đầy đủ các nhánh tái sử dụng, VinaKey2, CTV liên kết và điểm BSC. |

Chọn kỳ bằng ngày nghiệp vụ trên từng nguồn. Ngày tạo tệp và tên thư mục tuần chỉ hỗ trợ nhận diện. Giữ nguyên bản tải gốc để đối chiếu. Lưu dữ liệu thật và workbook kết quả trong `local/`.

## 2. Chuẩn hóa tài khoản và gán kênh

1. Đọc mã điểm, eLoad, HRM và tài khoản dưới dạng văn bản để giữ số 0 đầu. Loại khoảng trắng thừa và giữ giá trị gốc để đối soát.
2. Với IMS73, nếu `TAIKHOAN_DKY` bắt đầu bằng `NAN`, tra khớp chính xác mã điểm bán trong danh mục SMCS của Điểm Ủy quyền và lấy eLoad tương ứng. Số eLoad, tài khoản thường và mã HRM giữ giá trị hiện có. Nếu dùng workbook mẫu, ghi tài khoản đã chuẩn hóa vào `IMS73_CT!O`.
3. Ghép khóa chuẩn hóa với danh mục đối tượng để xác định cá nhân hoặc điểm bán, loại kênh và đơn vị quản lý. Giữ riêng **đơn vị giao dịch** từ IMS và **đơn vị quản lý** từ danh mục.
4. Với BRIS, chuẩn bị khóa tương đương `P01_SL PTM!L`. Workbook mẫu dùng cả giá trị trực tiếp, công thức tham chiếu HRM và tra cứu eLoad. Giữ riêng các dòng chưa áp dụng được quy tắc đã xác nhận.
5. Ghi nhận các dòng thiếu khóa, không tìm thấy mã hoặc có nhiều kết quả mâu thuẫn. Chỉ phân bổ số liệu vào một đối tượng khi xác định được khóa duy nhất.

## 3. Tính hiệu quả kênh tại HQ

| Chỉ tiêu | VinaPro | VinaKey1 |
| --- | --- | --- |
| Khóa điểm và đơn vị | HRM trong `Hiệu quả_VINAPRO!C`, đơn vị ở `D`. | eLoad trong `Hiệu quả_VINAKEY1!E`, đơn vị ở `C`. |
| Doanh thu PTM theo điểm | Cộng `IMS73_CT!I` theo tài khoản `O` khớp HRM và loại thuê bao `N = KHM`. | Cộng `IMS73_CT!I` theo tài khoản `O` khớp eLoad và `N = KHM`. |
| Doanh thu tái sử dụng | `Hiệu quả_VINAPRO!V` cộng `IMS73_CT!I` theo cùng HRM, `N = HIEN HUU`, `Y = TD49`. Cột `Y` tra số máy `G` trong `ALL_TD49!A:B`. | `Hiệu quả_VINAKEY1!W` cộng `Sim TSD!J` theo khóa `Sim TSD!F` khớp eLoad. |
| Doanh thu đánh giá | PTM + 80% doanh thu tái sử dụng, tại `Hiệu quả_VINAPRO!K`. Tổng hợp theo đơn vị ở `HQ!J`. | PTM + 50% doanh thu tại `W`, tại `Hiệu quả_VINAKEY1!K`. Tổng hợp theo đơn vị ở `HQ!S`. |
| Sản lượng PTM | Đếm dòng `P01_SL PTM!L` khớp HRM. Kết quả ở `Hiệu quả_VINAPRO!N`, tổng theo đơn vị ở `HQ!M`. | Đếm dòng `P01_SL PTM!L` khớp eLoad. Kết quả ở `Hiệu quả_VINAKEY1!J`, tổng theo đơn vị ở `HQ!V`. |

`HQ!A` là mã đơn vị. Với từng kênh, đếm tổng số điểm từ danh mục, số điểm có doanh thu trên 0 và số điểm có sản lượng trên 0. Tính tỷ lệ theo tổng số điểm của đơn vị và kênh, bao gồm cả điểm chưa phát sinh. `HQ!F = M + V` là tổng sản lượng PTM của hai kênh. VinaKey2 và CTV liên kết lấy phát sinh từ nhánh đơn hàng CTV, qua mã CTV ở `CTV XHH!N`.

Giữ riêng chỉ tiêu **doanh thu PTM theo đơn vị giao dịch**. `HQ!E` tra `TH!X` và đổi từ triệu đồng sang đồng. `TH!V:W` cộng `IMS73_CT!I` theo đơn vị giao dịch ở `E`, `N = KHM` và `R` là VinaPro hoặc VinaKey1. Các cột `HQ!J` và `HQ!S` dùng đơn vị quản lý điểm từ danh mục, nên hai cách tổng hợp có thể khác nhau theo đơn vị.

## 4. Tính BSC theo đơn vị

| Chỉ tiêu | Tử số, mẫu số và kế hoạch | Trọng số |
| --- | --- | --- |
| **I.8.1 — Nhân viên có CTV liên kết phát sinh** | Dùng mã nhân viên `BSC!V` khớp người quản lý ở `All_CTV XHH!H`. Tính tỷ lệ CTV có phát sinh (`All_CTV XHH!A > 0`) trong tổng CTV người đó quản lý. Nhân viên đạt khi tỷ lệ ít nhất 20%. Tỷ lệ nhân viên đạt của đơn vị bằng số người đạt chia cho số nhân viên trong danh sách BSC của đơn vị. So kết quả với kế hoạch ở `BSC!E`. | 30% |
| **I.8.2 — Điểm có phát sinh sản lượng** | Chia tổng điểm VinaPro và VinaKey1 có sản lượng (`HQ!L + HQ!U`) cho tổng số điểm hai kênh (`HQ!H + HQ!Q`). So tỷ lệ với kế hoạch ở `BSC!K`. | 30% |
| **I.8.3 — Doanh thu PTM và tái sử dụng** | Cộng `HQ!J + HQ!S`, đổi sang triệu đồng rồi chia kế hoạch ở `BSC!N`. Workbook mẫu lấy kế hoạch từ sheet `PL2.1_chỉ số KBH` của workbook ngoài. | 40% |

Với tỷ lệ hoàn thành kế hoạch là `t` ở dạng thập phân, workbook chấm **I.8.1** bằng 1 điểm khi dưới 30%, `10 × t − 2` khi từ 30% đến dưới 70%, và 5 điểm khi từ 70% trở lên. **I.8.2** và **I.8.3** dùng 1 điểm khi dưới 80%, `10 × t − 5` khi từ 80% đến dưới 100%, và 5 điểm khi đạt từ 100%. Điểm chung bằng `30% × I.8.1 + 30% × I.8.2 + 40% × I.8.3`.

## 5. Cách tự động hóa và đối soát

1. Kiểm tra cấu trúc tệp, tiêu đề, kỳ và mã định danh. Giữ tên tệp, sheet và dòng nguồn cho mỗi bản ghi.
2. Chuẩn hóa và ghép theo khóa, rồi tạo bảng **kỳ × đơn vị quản lý × kênh × đối tượng**. Lưu riêng doanh thu PTM, tái sử dụng, doanh thu đánh giá và sản lượng.
3. Ghép kết quả vào toàn bộ danh mục điểm để giữ các điểm chưa phát sinh. Tính `HQ`, sau đó ghép dữ liệu CTV, nhân viên và kế hoạch để tính `BSC`.
4. Đối chiếu số dòng và tổng doanh thu trước và sau chuẩn hóa. Đối chiếu kết quả ở cấp điểm, đơn vị và toàn tỉnh. Báo riêng dòng thiếu ánh xạ, kỳ không khớp và các chỉ tiêu thiếu đầu vào.
5. Xuất kết quả theo bố cục `HQ` và `BSC`. Khi cần cập nhật workbook mẫu, làm trên bản sao và chỉ ghi vùng dữ liệu đã xác định, giữ các cột tính và bảng tra.

## 6. Các điểm chưa chốt

| Nội dung | Cần xác nhận |
| --- | --- |
| Tệp danh mục SMCS | Đường dẫn menu đầy đủ, tên sheet, tên cột trong tệp xuất và cách chọn bản danh mục đúng kỳ khi điểm đổi đơn vị hoặc eLoad. |
| Kỳ và dấu hiệu xuất báo cáo | Mốc ngày theo từng nguồn, trường ngày dùng để lọc IMS73 và điều kiện nhận biết bảng IMS đã tải xong trước khi xuất. |
| Khóa sản lượng BRIS | Điều kiện dùng trực tiếp User/eLoad, dùng HRM hoặc tra eLoad. Workbook mẫu có công thức tra ngoài trong cột `L` và có 11 dòng chi tiết chưa khớp với tệp BRIS đối chiếu. |
| Sản lượng và đơn vị | Workbook đếm dòng theo khóa tài khoản. Cần xác nhận cách xử lý trùng thuê bao và quy tắc sử dụng đơn vị giao dịch hay đơn vị quản lý khi lập từng chỉ tiêu. |
| Tái sử dụng và bảng tra | Nguồn `Sim TSD`, `ALL_TD49` theo kỳ, dữ liệu minh chứng khi có tái sử dụng VinaKey1 và hiệu lực của tỷ lệ 80%/50% theo kỳ. |
| BSC đầy đủ | Danh sách nhân viên thuộc phạm vi chấm, quy tắc xử lý người không quản lý CTV, nguồn kế hoạch và cách xử lý kế hoạch bằng 0 hoặc thiếu. |
| Khác biệt với workbook cũ | Một số công thức doanh thu VinaKey1 tham chiếu eLoad ở dòng khác. `HQ!AG`, `AH` và `AJ` có tham chiếu `#REF!`. Cần đối soát kết quả theo khóa đúng ở cấp điểm và đơn vị trước khi dùng làm số báo cáo. |
