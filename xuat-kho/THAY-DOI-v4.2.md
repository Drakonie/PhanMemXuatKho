# Bản 4.2 — vùng đọc ảnh, khách hàng và giao diện

- Tự bỏ lề trắng quanh dữ liệu, phóng ảnh nhỏ để đọc rõ hơn. Có nút chọn vùng đọc bằng kéo khung hoặc nhập tỷ lệ; khôi phục toàn ảnh hoặc quay lại chế độ tự động.
- Bổ sung đọc từng ô và nhận diện nhóm khách hàng cho bảng Excel có lưới, cột NT/MG/Tổng/Thành tiền. Tạo từng phiếu theo khách với tên hàng/đơn giá chung; không cộng trùng NT/MG với Tổng. Tiêu đề đơn giá suy đoán được ghi rõ để đối chiếu.
- Tick chọn xuất từng phiếu khách hàng. Giữ ảnh gốc, các dòng nguồn và bản chữ riêng từng khách; các số chưa rõ/lệch tiền vẫn chặn xuất. Sửa OCR của một khách giữ chỉnh sửa và bản nháp các khách còn lại.
- Chuyển giữa Normal và Linh linh yêu mà không tải lại trang hoặc mất dữ liệu. Nhớ lựa chọn giao diện; cả hai bố cục được kiểm tra 320–1440px.
- Giữ bộ khởi chạy có thư viện Excel đi kèm, không cần pip hay Internet để mở ứng dụng.

## Xác minh

130 kiểm thử Python; các bộ kiểm tra Chromium về Excel, OCR tiếng Việt thật, lỗi hồi quy và các chức năng mới. Ảnh thật kèm theo tách 5 nhóm khách hàng, giữ 17 dòng nguồn mỗi nhóm, ngày 02/10/2026; các ô sai/chưa rõ yêu cầu sửa và chặn xuất.

Ảnh này có chữ nhỏ nên nhận diện tên/số chưa chính xác toàn bộ. Không huấn luyện trọng số mô hình hoặc xác nhận số tiền thật từ kết quả OCR. Kiểm thử trên Linux; chưa chạy trực tiếp trên Windows.

Giải nén toàn bộ ZIP vào thư mục mới và mở Chay-Windows.bat. Python 3.10 trở lên.
