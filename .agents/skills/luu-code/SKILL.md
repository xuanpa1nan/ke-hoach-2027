---
name: luu-code
description: Kiểm tra thay đổi trong repo hiện tại, commit phần thuộc phạm vi yêu cầu và đẩy nhánh lên GitHub. Dùng khi người dùng yêu cầu lưu code lên GitHub hoặc gọi skill luu-code.
---

# Lưu code lên GitHub

1. Xác định repo, nhánh hiện tại, upstream và remote đích. Kiểm tra URL fetch và push để bảo đảm đích là repo GitHub dự kiến. Giữ nhánh hiện tại. Nếu đang detached HEAD hoặc đích chưa rõ, hỏi người dùng trước khi push.
2. Kiểm tra `git status --short --branch`, diff đã stage, diff chưa stage và danh sách tệp mới. Nếu người dùng yêu cầu lưu toàn bộ, đưa vào cả thay đổi mới, sửa và xóa thuộc dự án. Bỏ qua thông tin đăng nhập, dữ liệu nội bộ, workbook thật, tệp tải về và ảnh chụp công việc. Stage theo đường dẫn đã kiểm tra. Nếu index có thay đổi nằm ngoài phạm vi commit, giữ nguyên chúng và hỏi cách xử lý trước khi tạo commit.
3. Xem lại `git diff --cached --stat` và `git diff --cached --check`. Nếu có thay đổi đã stage, tạo commit với thông điệp mô tả nội dung. Nếu không có thay đổi để commit, kiểm tra xem nhánh còn commit chưa push hay không.
4. Push tới upstream đã xác định. Nếu nhánh mới chưa có upstream và remote đích rõ ràng, dùng `git push -u <remote> <nhánh>`. Khi Git từ chối vì lịch sử phân kỳ, fetch và báo tình trạng để người dùng quyết định cách tích hợp. Tránh force-push và tránh tự sửa lịch sử. Nếu HTTPS cần xác thực, kiểm tra `gh auth status` và dùng GitHub CLI làm credential helper cho lần push nếu phù hợp.
5. Xác minh commit trên nhánh từ xa và kiểm tra lại trạng thái Git. Báo nhánh, mã commit, kết quả push và thay đổi còn lại nếu có. Nếu push thất bại sau khi commit, nêu rõ commit đã lưu cục bộ và nguyên nhân.
