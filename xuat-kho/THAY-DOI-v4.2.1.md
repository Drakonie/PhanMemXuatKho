# Bản 4.2.1 — giảm dung lượng gói chạy

- Chỉ giữ hai bộ OCR LSTM mà ứng dụng sử dụng: SIMD và bộ dự phòng cho trình duyệt không hỗ trợ SIMD. Bỏ các bộ OCR cũ và file WASM đóng gói trùng; hai bộ giữ lại đã chứa đầy đủ WASM.
- Giữ nguyên mô hình tiếng Việt/Anh, bộ nhận diện, template, sticker/nền/banner, chức năng chọn vùng đọc và xuất theo khách hàng.
- Không đưa ảnh chụp giao diện/kết quả kiểm thử, cache Python hoặc log vào ZIP. Vẫn giữ năm file mẫu đầu vào và mã kiểm thử.
- Giữ toàn bộ thư viện Excel cần thiết để mở ứng dụng không cần pip hoặc Internet, kể cả khi dùng Python 3.13t.
- Thêm `donggoi.py` để đóng gói lại ở mức nén 9, kiểm tra OCR/thư viện theo SHA-256 và kiểm tra ZIP trước khi lưu.
- Bổ sung kiểm tra OCR thực tế với cả hai bộ SIMD và không SIMD, chặn mọi yêu cầu đến dịch vụ bên ngoài trong bài kiểm tra.

Giải nén toàn bộ ZIP vào thư mục mới, rồi mở `Chay-Windows.bat`. Cần Python 3.10 trở lên. Không cần triển khai lên Vercel.
