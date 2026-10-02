# Bản 4.3.0 — nâng cấp nhận diện Excel và ảnh

- Nhận tiêu đề viết liền, ký tự Unicode toàn chiều rộng và ký tự vô hình; áp dụng cho Excel và tiêu đề bảng chữ OCR.
- Nhận tiêu đề 1–3 dòng và ô gộp, ưu tiên số lượng thực xuất; cho chỉnh thủ công số dòng tiêu đề đến 3. Hỗ trợ ngày viết bằng dấu chấm.
- Thêm tự làm rõ ảnh thiếu tương phản và chỉnh nghiêng nhẹ trước OCR. Chỉ xoay khi có tín hiệu rõ; giữ ảnh gốc, ghi chú thao tác đã áp dụng và có nút tắt tùy chọn.
- Các trường hợp nghiêng +3°, nghiêng −4° kèm giảm tương phản trên ảnh minh họa đều đọc đúng ba mặt hàng và xuất tổng 340.000 đồng. 142 kiểm thử Python đạt, gồm các bố cục mới và bảo toàn chức năng cũ.
- Sửa tách bảng nhiều khách hàng khi ảnh có thêm lề trắng sau chỉnh nghiêng. Ảnh mẫu nghiêng 3° vẫn giữ đủ 5 khách hàng × 17 dòng, bỏ phần tổng kết; tiêu đề được suy đoán có cảnh báo để đối chiếu.
- Tiếp tục dùng bộ OCR Việt/Anh đi kèm, không cần API key hoặc Internet. Không huấn luyện lại trọng số từ một ảnh; chữ/số mờ vẫn cần đối chiếu trước khi xuất.

Giải nén toàn bộ ZIP vào thư mục mới. Đóng bản cũ và mở `Chay-Windows.bat` trong thư mục mới. Cần Python 3.10 trở lên.
