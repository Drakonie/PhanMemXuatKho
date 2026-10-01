# Công cụ xuất kho từ Excel nhiều sheet

Ứng dụng Python chạy trên máy của bạn, đọc file Excel có bố cục thay đổi và tạo file `.xlsx` theo `templates/Template.xlsx` chuẩn đã cung cấp.

## Chạy nhanh

Cần Python 3.10 trở lên. Trên Windows, giải nén toàn bộ dự án rồi nhấp đúp `Chay-Windows.bat`. Bộ khởi chạy kiểm tra Python, dùng cùng một Python để cài thư viện/chạy ứng dụng và chỉ mở trình duyệt sau khi máy chủ sẵn sàng. Giữ cửa sổ lệnh mở khi sử dụng. Nếu cổng 8080 đang bận, ứng dụng thử 8081–8089 và mở đúng địa chỉ. Nếu lỗi, gửi nội dung cửa sổ hoặc file `Loi-khoi-dong.log`. Hoặc mở terminal trong thư mục dự án:

```sh
python -m pip install -r requirements.txt
python app.py
```

Mở **http://127.0.0.1:8080** trong trình duyệt. Dừng ứng dụng bằng Ctrl+C. Nếu máy dùng lệnh `python3`, thay `python` bằng `python3`.

## Sử dụng

1. Chọn file Excel dữ liệu `.xlsx`. Template chuẩn được dùng sẵn; bạn có thể tải lên bản khác có cùng cấu trúc ô.
2. Bấm **Nhận diện dữ liệu**. Ứng dụng quét từng sheet, tìm dòng tiêu đề trong 100 dòng đầu và nhận diện cột theo tên, không cố định số thứ tự cột.
3. Chọn sheet cần xuất. Nhập hoặc kiểm tra **Khách hàng**, **Người nhận**, **Ngày xuất**, **Địa chỉ khách hàng** và **Số phiếu xuất**. Dữ liệu khách hàng trên đầu sheet được gợi ý; cần kiểm tra khi file có nhiều tên đơn vị.
4. Kiểm tra bảng hàng và các cảnh báo. Nếu nhận diện chưa đúng, mở **Cột nhận diện**, chọn dòng tiêu đề/các cột và bấm **Đọc lại theo cột đã chọn**. Lỗi dữ liệu phải sửa trong file nguồn rồi tải lại.
5. Bấm **Tải file xuất kho**. Các sheet được chọn trở thành các phiếu riêng trong một file Excel.

## Quy tắc

- Mỗi sheet nguồn có một bảng hàng liên tục, một khách hàng; dừng ở dòng tổng cộng. Nhiều bảng hoặc nhiều khách hàng trong cùng sheet cần tách sheet trước.
- Các trường chính: tên hàng, số lượng, đơn giá, thành tiền. ĐVT, mã hàng và ghi chú là tùy chọn. Nhận tên có dấu/không dấu và một số biến thể phổ biến như `Tên sản phẩm`, `Mặt hàng`, `SL`, `ĐVT`, `Thành tiền công nợ`.
- Nếu không có thành tiền, tính số lượng × đơn giá và làm tròn đến đồng. Nếu có thành tiền mà không có đơn giá, suy ra đơn giá khi số lượng khác 0. Sai lệch thành tiền lớn hơn 1 đồng, số âm hoặc số không đọc được sẽ chặn xuất sheet đó.
- Đọc giá trị công thức đã lưu trong Excel. Với công thức chưa có kết quả lưu, hỗ trợ phép toán cơ bản, tham chiếu ô/các sheet và SUM vùng trong cùng sheet. Công thức khác cần mở trong Excel, tính lại và lưu trước khi tải lên.
- Số Excel được giữ theo kiểu số. Chuỗi số hỗ trợ `4,3`, `120.000`, `1.234,50`, `1,234.50`; chuỗi có một dấu phân cách và đúng 3 chữ số sau dấu được hiểu là nhóm nghìn (trừ phần nguyên là 0). Với dữ liệu dễ nhầm, nên lưu ô dưới kiểu số trong Excel.
- Sheet có tên tổng hợp/summary/total và sheet ẩn không được chọn mặc định. Đây là quy tắc gợi ý theo tên, không phải kiểm tra trùng nội dung; hãy chọn đúng sheet trước khi xuất.
- Số lượng 0 được giữ mặc định; có tùy chọn bỏ khi xuất. Không quy đổi hoặc gộp đơn vị tính giữa các sheet.
- Xuất cả số lượng yêu cầu và thực xuất theo số lượng nguồn. Thành tiền và tổng tiền xuất dạng số để hiển thị được ngay cả khi trình đọc không tính lại công thức. Tự viết tổng tiền bằng chữ tiếng Việt.
- Giữ tiêu đề doanh nghiệp, phần nội dung mặc định, định dạng và chữ ký của template. Tên khách hàng, người nhận, ngày và địa chỉ của ví dụ cũ được thay bằng dữ liệu đã kiểm tra. Vượt 15 mặt hàng sẽ thêm dòng và đẩy phần tổng/chữ ký xuống dưới; khi in, bảng có thể dùng nhiều trang.
- Chỉ hỗ trợ `.xlsx`; `.xls` cần lưu lại thành `.xlsx`. Template khác phải giữ bố cục chuẩn (tiêu đề A5, bảng dòng 11, tổng cộng B27/H27). Không chạy macro.

## Dữ liệu mẫu đã kiểm tra

`samples/Phieu-xuat-kho-mau.xlsx` chứa ba phiếu ngày 01/10/2026:

| Sheet | Tổng tiền |
|---|---:|
| Trung tâm | 593.000 đồng |
| Phân hiệu 1 (Nghĩa Hưng) | 432.000 đồng |
| Phân hiệu 2 (An Hà) | 544.000 đồng |
| Tổng ba phiếu | 1.569.000 đồng |

Sheet Tổng hợp chung được bỏ chọn để tránh xuất trùng. Tên khách hàng dùng tên đơn vị ghi trên đầu mỗi sheet. Số phiếu và địa chỉ để trống vì nguồn chưa cung cấp. Phiếu Trung tâm giữ dòng Rau muống có số lượng 0 và ghi chú của nguồn.

## Kiểm thử

```sh
python -m unittest discover -s tests -v
```

Có 10 bài kiểm thử (8 bài xử lý Excel và 2 bài kiểm tra khởi chạy): file mẫu, cột đảo thứ tự, công thức chưa có cache, công thức liên sheet, nhiều hơn 15 dòng, dữ liệu thiếu/sai, số Việt Nam, số tiền bằng chữ và tùy chọn bỏ dòng 0. Giao diện đã kiểm tra tải file, nhận diện 4 sheet/chọn 3 sheet, tổng tiền, tải kết quả và bố cục điện thoại bằng Chromium.

Ứng dụng dùng thư viện chuẩn Python và openpyxl. File tải lên được xử lý trên máy chạy ứng dụng; phiên được giữ trong bộ nhớ tối đa 1 giờ (dọn khi có yêu cầu tiếp theo), giới hạn 20 phiên và tổng tải lên 25 MB. Khởi động lại sẽ xóa các phiên. Mặc định chỉ lắng nghe localhost. Đây là công cụ chạy cục bộ, chưa triển khai lên website công khai.

Bản sửa khởi chạy: dùng `py -3` sau khi kiểm tra chạy được, không gọi file .py trực tiếp qua liên kết Windows và không mở trang trước khi máy chủ sẵn sàng. Kiểm thử Python/máy chủ được thực hiện trên Linux; file .bat cần được chạy thử trên máy Windows của bạn.
