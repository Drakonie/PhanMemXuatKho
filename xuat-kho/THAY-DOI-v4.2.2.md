# Bản 4.2.2 — nền trắng và xuất file dễ dùng hơn

- Giao diện Normal dùng nền trắng hoàn toàn, giữ lựa chọn khi mở lại trang.
- Chữ hướng dẫn dễ đọc hơn, nút và vùng chọn lớn hơn. Quy trình ba bước rõ ràng: chọn và đọc file, kiểm tra dữ liệu, tạo và tải file.
- Luồng xuất có giới hạn thời gian chờ 90 giây và nút hủy. Yêu cầu không phản hồi sẽ báo cách thử lại và mở lại các nút, giữ dữ liệu đang chỉnh sửa. Trạng thái/lỗi được hiển thị ngay cạnh nút xuất.
- Sau khi tạo file, có đường dẫn tải thủ công để dùng khi trình duyệt không tải tự động. Nhận lại file đã tạo không cần gửi yêu cầu xuất lần nữa.
- Giữ bản gọn với OCR Việt/Anh, template và thư viện Excel đi kèm; không cần pip hay Internet để chạy local.

Kiểm tra: 135 kiểm thử Python đạt. Kiểm thử trình duyệt xác nhận Excel/ZIP, OCR thật và ảnh nhiều khách hàng; đồng thời giả lập treo trước/sau khi nhận đầu phản hồi, hủy, hết thời gian chờ, giữ chỉnh sửa và tải thủ công khi tải tự động bị chặn. Hai giao diện không tràn ngang tại 320–1440 px; nút chính cao tối thiểu 52 px.

Giải nén toàn bộ ZIP vào thư mục mới. Đóng cửa sổ chạy bản cũ rồi mở `Chay-Windows.bat` trong thư mục mới. Cần Python 3.10 trở lên.
