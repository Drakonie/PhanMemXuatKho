# Xuất kho Huyền An 68 — phiên bản 4.3.0 gọn

Công cụ chạy trên máy của bạn, nhận diện dữ liệu Excel nhiều sheet hoặc bảng hàng trong ảnh và xuất phiếu theo template chuẩn đã cung cấp. Không cần tài khoản hoặc dịch vụ trực tuyến để xử lý dữ liệu.

Bản gọn giữ đầy đủ Excel, OCR tiếng Việt/Anh, chọn vùng đọc, xuất theo khách hàng và hai giao diện. Gói chỉ mang các bộ OCR ứng dụng sử dụng, không chứa bản WASM trùng hoặc ảnh chụp màn hình kiểm thử. Thư viện Excel và mô hình OCR vẫn có sẵn để chạy không cần pip hay Internet.

Bản 4.2.2 đổi Normal sang nền trắng, tăng cỡ chữ và kích thước nút. Quy trình ba bước chỉ rõ chọn/đọc file, kiểm tra dữ liệu và tải kết quả. Khi xuất, có nút **Hủy tạo file** và giới hạn thời gian chờ 90 giây; sau khi tạo xong có nút **Tải lại file đã tạo** để dùng nếu trình duyệt không tự tải. Trạng thái/lỗi xuất hiện ngay cạnh nút xuất; hủy hoặc lỗi vẫn giữ dữ liệu đã chỉnh sửa.

Giao diện “Linh linh yêu” dùng bộ ảnh bạn cung cấp: banner “Xuất kho cùng Linh”, nền pastel và các sticker mèo. Bảng dữ liệu giữ nền sáng, dễ đối chiếu; bố cục thích ứng với máy tính và điện thoại. Nút **Normal / Linh linh yêu** ở đầu trang đổi giao diện ngay, giữ file tải lên và mọi chỉnh sửa; lựa chọn giao diện được nhớ trong trình duyệt.

## Chạy trên Windows

1. Giải nén **toàn bộ** file ZIP vào một thư mục mới.
2. Nhấp đúp `Chay-Windows.bat`.
3. Trình duyệt mở khi ứng dụng sẵn sàng. Giữ cửa sổ chạy ứng dụng mở trong lúc sử dụng.

Khi nâng cấp, đóng cửa sổ chạy bản cũ trước khi mở `Chay-Windows.bat` trong thư mục mới.

Cần Python 3.10 trở lên. Thư viện `openpyxl 3.1.5` và `et_xmlfile 2.0.0` đã có sẵn trong thư mục `vendor`: không cần pip hoặc Internet để mở ứng dụng. Giữ toàn bộ thư mục `vendor` cạnh `khoidong.py`. Khi Python đang dùng là bản free-threaded như `python3.13t.exe`, bộ khởi chạy ưu tiên Python tiêu chuẩn nếu đã cài; nếu không, tiếp tục dùng Python hiện tại với thư viện đi kèm. Ứng dụng sử dụng lại phiên đang chạy và thử cổng 8080–8089 nếu cổng bị chiếm. Nếu lỗi, gửi nội dung cửa sổ hoặc `Loi-khoi-dong.log`.

Chạy bằng terminal trên Windows/macOS/Linux:

```sh
python khoidong.py
```

Hoặc `python app.py --port 8080` và mở `http://127.0.0.1:8080`. Dừng bằng Ctrl+C. Máy dùng lệnh `python3` thì thay `python` bằng `python3`.

## Quy trình sử dụng

1. Chọn **Excel nhiều sheet** để tải `.xlsx`, hoặc **Ảnh chụp bảng hàng** để chọn JPG/PNG/WebP (tối đa 10 ảnh, 20 MB mỗi ảnh), kéo thả hay chụp bằng điện thoại. Template chuẩn được dùng sẵn. Có thể thay bằng bản template cùng cấu trúc ô.
2. Bấm **Nhận diện dữ liệu** và kiểm tra danh sách sheet cùng điểm nhận diện.
3. Chọn sheet cần xuất. Sheet tổng hợp, sheet ẩn và bảng suy đoán không có tiêu đề được bỏ chọn mặc định.
4. Kiểm tra khách hàng, người nhận, ngày xuất, địa chỉ và số phiếu. Nếu nhiều tên đơn vị xuất hiện trên đầu sheet, nhập tên khách hàng đúng thay vì dựa vào đề xuất.
5. Kiểm tra từng dòng hàng. Có thể sửa tên hàng, đơn vị, số lượng, đơn giá và ghi chú trên giao diện; số lượng/đơn giá thay đổi sẽ cập nhật thành tiền. Mã hàng từ nguồn được giữ khi xuất. Bỏ chọn dòng không cần xuất hoặc khôi phục các thay đổi.
6. Nếu cột đọc sai, mở phần dữ liệu nguồn và cấu hình nhận diện: chọn dòng/số dòng tiêu đề và cột tương ứng rồi đọc lại. Dòng hàng chưa đọc được cần chỉnh nhận diện hoặc sửa Excel nguồn; ứng dụng không tự bỏ những dòng lỗi này.
7. Xuất **một workbook** với mỗi sheet được chọn thành một phiếu, hoặc **ZIP** chứa mỗi phiếu một file Excel riêng.

Tìm kiếm và phân trang giúp xử lý bảng dài. Chức năng áp dụng thông tin chung giúp nhập cùng ngày/khách hàng/địa chỉ cho các sheet đã chọn. Kiểm tra phạm vi áp dụng trước khi dùng.

## Nhận diện nâng cấp ở bản 4.3

- Excel: nhận tiêu đề viết liền như `TenHangHoa`, `SoLuong`, `DonGia`, `ThanhTien`; xử lý chữ Unicode toàn chiều rộng và ký tự vô hình trong tiêu đề.
- Tìm tiêu đề tối đa ba dòng, kể cả ô gộp và nhóm số lượng yêu cầu/thực xuất. Phần chỉnh nhận diện cho phép chọn 1–3 dòng. Ngày dạng `02.10.2026` được nhận diện.
- Bảng chữ OCR cũng nhận các tiêu đề viết liền khi cột bị đảo vị trí.
- Tùy chọn **Tự làm rõ / chỉnh nghiêng nhẹ** mặc định bật: điều chỉnh tương phản ảnh tối/mờ tương phản và tìm góc nghiêng trong khoảng ±7°. Chỉ xoay khi tín hiệu hàng chữ đủ rõ. Không sửa file ảnh gốc; các thao tác đã áp dụng xuất hiện dưới ảnh đối chiếu.
- Ảnh thử nghiêng +3° và nghiêng −4° kèm giảm tương phản đều đọc đúng ba mặt hàng, số lượng, đơn giá, tổng 340.000 đồng và xuất được Excel. Đây là kiểm tra trên ảnh minh họa, không bảo đảm mọi ảnh đều chính xác. Chữ nhòe, ảnh méo phối cảnh hoặc nghiêng mạnh vẫn cần chụp lại/chọn vùng và kiểm tra số liệu.

## Nhập dữ liệu từ ảnh

1. Chụp bảng thẳng, đủ sáng, rõ tên hàng và các cột số. Có thể xoay ảnh từng bước 90° trước khi đọc; ảnh chữ in rõ cho kết quả tốt hơn chữ viết tay.
2. Bấm **Đọc ảnh và nhận diện**. Bộ OCR tiếng Việt + tiếng Anh chạy ngay trong trình duyệt, có tiến trình và nút hủy. Mỗi ảnh thành một sheet để kiểm tra riêng.
3. Đối chiếu ảnh gốc cạnh bảng kết quả. Phóng to ảnh, sửa mặt hàng/số lượng/đơn giá và điền khách hàng, ngày xuất nếu chưa nhận diện được.
4. Nếu OCR làm mất cột hoặc dòng, mở phần nguồn ảnh, sửa văn bản với các cột ngăn bằng `|`, tab hoặc ít nhất hai dấu cách rồi **Đọc lại văn bản đã sửa**. Ví dụ:

```text
Khách hàng: Trường Minh An
Ngày xuất: 01/10/2026
Tên hàng | ĐVT | Số lượng | Đơn giá | Thành tiền
Thịt heo | kg | 2 | 100.000 | 200.000
```

5. Các phiếu từ ảnh **chưa chọn xuất mặc định**. Đánh dấu đã kiểm tra từng phiếu; đọc lại văn bản hoặc cột của ảnh yêu cầu kiểm tra lại. Dòng thiếu số, số không đọc được và lệch thành tiền vẫn bị chặn khi xuất.

OCR dùng Tesseract.js 6 và mô hình tiếng Việt/Anh được đóng gói cùng dự án. Không cần khóa API, không gọi dịch vụ OCR ngoài. Ảnh nằm trong bộ nhớ trình duyệt; văn bản và vị trí chữ nhận diện được gửi đến máy chủ ứng dụng để dựng bảng Excel. Phiên không được lưu vào cơ sở dữ liệu. Khi chạy local, tài nguyên OCR có sẵn nên xử lý ảnh không cần Internet khi dùng bản ZIP đầy đủ. Lần tải OCR đầu trên web có thể lâu hơn vì cần tải khoảng 7 MB tài nguyên; trình duyệt lưu bộ ngôn ngữ để dùng lại.

Điểm OCR và điểm nhận diện cấu trúc chỉ giúp tìm phần cần kiểm tra, không bảo đảm chữ hoặc số đã đúng. Nhận diện hình ảnh không thể tự bảo đảm mọi dòng trong ảnh được đọc đủ; luôn đối chiếu ảnh, nhất là số tiền và các dòng mờ.


## Ảnh bảng chung nhiều khách hàng

Bản 4.2 bổ sung đọc từng ô cho ảnh Excel có lưới như ảnh mẫu: tên hàng/ĐVT/đơn giá dùng chung, mỗi khách có các cột NT, MG, Tổng và Thành tiền. Ứng dụng lấy Tổng của từng khách để tạo phiếu riêng, giữ ảnh gốc để đối chiếu. Nếu tiêu đề đơn giá chưa đọc rõ, cột đề xuất sẽ có lưu ý để kiểm tra.

- **Tự bỏ lề trắng** là mặc định. Mỗi ảnh có nút **Chọn vùng đọc**: kéo khung hoặc nhập tỷ lệ Trái/Trên/Rộng/Cao. Nút **Toàn bộ ảnh** bỏ vùng đã chọn. Xoay ảnh sẽ đặt lại vùng đọc.
- Giữ tên hàng, đơn giá chung, tiêu đề khách hàng và các cột cần xuất trong khung. Ngày/khách hàng nằm ngoài vùng đọc phải nhập bổ sung.
- Sau nhận diện, dùng mục **Chọn xuất theo khách hàng**. Tick đồng nghĩa đã đối chiếu dữ liệu; chỉ những phiếu được tick mới xuất. Các phiếu lỗi vẫn bị chặn khi xuất.
- Sửa văn bản OCR của một khách và đọc lại chỉ thay phiếu của khách đó, giữ chỉnh sửa/bản nháp của khách khác. Bảng chữ đã tách theo từng khách để dễ sửa.
- Ảnh mẫu `samples/anh-nhieu-khach-hang.jpg` do bạn cung cấp: thử OCR thật tách được **5 nhóm**, giữ **17 dòng nguồn mỗi nhóm** và ngày 02/10/2026. Ảnh nhỏ khiến nhiều tên/số đọc sai; đây là kiểm tra đầy đủ cấu trúc, không phải xác nhận 85 dòng đã chính xác. Các dòng chưa đọc được nằm trong bảng chữ để sửa, không tự bị coi là số 0.

Chức năng này cải thiện quy tắc nhận diện và xử lý ảnh; không huấn luyện lại trọng số Tesseract từ một ảnh. Với bố cục chưa nhận ra, ứng dụng dùng luồng OCR thông thường và phần sửa văn bản.

## Nhận diện mới

- Không cố định số thứ tự cột. Hỗ trợ nhiều nhãn tiếng Việt có/không dấu và tiếng Anh: tên hàng/sản phẩm/vật tư/nguyên liệu, số lượng/SL/Qty, đơn giá/Unit price, thành tiền/Amount, ĐVT/Unit, mã hàng/SKU và ghi chú.
- Đọc tiêu đề từ một đến ba dòng và các ô gộp. Ưu tiên số lượng **thực xuất** nếu bảng có cả yêu cầu và thực xuất. Nhiều cột vẫn cùng phù hợp sẽ yêu cầu chọn thủ công.
- Hỗ trợ các bảng hoặc tiêu đề lặp trong một sheet khi cùng thông tin phiếu. Bảng có khách hàng/ngày/số phiếu/người nhận/địa chỉ khác nhau phải tách sheet để tránh gộp phiếu sai.
- Đề xuất cột cho bảng không có tiêu đề khi nhiều dòng chứng minh quan hệ số lượng × đơn giá = thành tiền. Đề xuất này cần xác nhận và có thể không được tạo nếu dữ liệu chưa đủ rõ.
- Tìm khách hàng, ngày, địa chỉ, người nhận, số phiếu và nội dung trong phần đầu bảng. Không tự dùng tên một đơn vị khi có nhiều ứng viên.
- **Điểm nhận diện là điểm quy tắc**, có lý do đi kèm, không phải xác suất chính xác hoặc bảo đảm file đã đúng. Vẫn cần kiểm tra kết quả.

## Kiểm tra số liệu và công thức

- Bắt buộc tên hàng, số lượng và đơn giá hoặc thành tiền. ĐVT, mã hàng và ghi chú là tùy chọn.
- Thiếu thành tiền: tính số lượng × đơn giá, làm tròn đến đồng. Thiếu đơn giá: suy ra từ thành tiền/số lượng khi số lượng khác 0. Các phép suy ra đều có cảnh báo.
- Chặn số âm, số không đọc được, công thức lỗi và lệch thành tiền quá 1 đồng. Lỗi thành tiền ở dòng đã đọc có thể sửa trên giao diện; dòng không đọc được phải sửa nguồn/cột nhận diện. Dữ liệu được kiểm tra lại ở phía máy chủ trước khi xuất.
- Đọc kết quả công thức đã lưu trong Excel. Nếu chưa có kết quả lưu, hỗ trợ phép toán cơ bản, tham chiếu ô/vùng/liên sheet và một số hàm: SUM, ROUND, ROUNDUP, ROUNDDOWN, IF, IFERROR, MIN, MAX, PRODUCT, AVERAGE, ABS, INT. Các công thức khác và biểu thức lũy thừa phức tạp chưa có kết quả lưu cần mở bằng Excel, tính lại và lưu trước khi tải lên. Đây không phải bộ tính toán đầy đủ của Excel.
- Hỗ trợ ô kiểu số và chuỗi số Việt Nam/quốc tế như `4,3`, `120.000`, `1.234,50`, `1,234.50`. Chuỗi số lượng `1.234` có thể mang nghĩa khác nhau; ứng dụng nêu cảnh báo và yêu cầu kiểm tra. Nên lưu số dưới kiểu số trong Excel.
- Giữ số lượng 0 mặc định; có tùy chọn bỏ khi xuất. Không quy đổi hoặc gộp đơn vị tính giữa các sheet.

## Template và file xuất

Template chuẩn ở `templates/Template.xlsx`: tiêu đề A5, bảng hàng dòng 11, 15 dòng mẫu từ 12–26, tổng cộng B27/H27. Template thay thế phải giữ cấu trúc này.

Giữ thông tin doanh nghiệp và chữ ký của mẫu. Thay dữ liệu khách hàng/người nhận/ngày/địa chỉ cũ, xóa các mặt hàng ví dụ, thêm dòng khi vượt 15 hàng, cập nhật tổng tiền và tổng bằng chữ tiếng Việt. Số lượng yêu cầu và thực xuất cùng lấy số lượng đã kiểm tra. Thành tiền/tổng tiền được ghi dạng số để hiển thị ngay trong các trình đọc không tính lại công thức.

Không hỗ trợ trực tiếp `.xls`, file có mật khẩu hoặc nhiều khách hàng nằm chung một bảng không có cấu trúc phân chia rõ. `.xls` cần lưu lại thành `.xlsx`. Phân loại sheet tổng hợp dựa vào tên sheet là gợi ý; hãy kiểm tra để tránh chọn cả chi tiết và tổng hợp.

## Ví dụ kèm theo

- `samples/data-mau.xlsx`: file gốc bạn cung cấp.
- `samples/Phieu-xuat-kho-mau.xlsx`: 3 phiếu ngày 01/10/2026, tổng **1.569.000 đồng** (593.000 + 432.000 + 544.000); sheet Tổng hợp chung không xuất.
- `samples/data-da-dang.xlsx`: dữ liệu **minh họa tạo thêm** để thử cột đảo, tiêu đề gộp, bảng chỉ có thành tiền và bảng không tiêu đề. Không phải dữ liệu kinh doanh thật.
- `samples/anh-minh-hoa.png`: ảnh bảng hàng **giả lập** để thử OCR, không phải dữ liệu kinh doanh thật.

Ảnh chụp giao diện và kết quả kiểm thử không nằm trong gói gọn. Các bài kiểm tra trình duyệt có thể tạo lại chúng khi chạy.

## Kiểm thử và vận hành

```sh
python -m unittest discover -s tests -v
```

**142 kiểm thử Python tự động đều đạt**, gồm API local/Vercel, bảo toàn dòng OCR, xuất theo template và kiểm tra gói ZIP. Kiểm thử gồm nhận diện nhiều bố cục, công thức, số tiền, lỗi dữ liệu, chỉnh/bỏ dòng, workbook/ZIP, giữ template/thêm dòng và khởi chạy/tái sử dụng phiên. Giao diện được kiểm tra bằng Chromium trên máy tính và các bố cục điện thoại từ 320 px. Các bộ kiểm tra trình duyệt kiểm tra luồng Excel, OCR thực tế và các lỗi hồi quy: giữ bản nháp, hủy/thử lại OCR, phóng to ảnh, số thập phân và xuất file. Các tệp JavaScript cũng vượt qua kiểm tra cú pháp. Luồng nhận diện ảnh thực tế đã được kiểm tra với Tesseract tiếng Việt: 3 dòng hàng tổng 340.000 đồng, sửa văn bản thành 4 dòng tổng 370.000 đồng, chặn dòng thiếu số lượng và xuất ZIP từ 2 ảnh. Cả hai bộ OCR SIMD/không SIMD đọc đúng ba dòng mẫu khi tải mới mô hình Việt/Anh từ ứng dụng; không cần WASM riêng hoặc dịch vụ bên ngoài.

File tải lên chỉ được xử lý trên máy chạy ứng dụng. Phiên nằm trong bộ nhớ, hết hạn sau 1 giờ không hoạt động; dọn khi có yêu cầu xử lý tiếp theo. Khởi động lại xóa phiên. Tối đa 20 phiên, 25 MB tổng tải lên, 100 MB sau giải nén. Mặc định chỉ lắng nghe localhost. Các thư viện Excel openpyxl và et_xmlfile được đóng gói sẵn; giao diện và OCR không dùng CDN.

Bản 4.1.1 bổ sung kiểm tra chọn runtime Windows, giữ thông báo lỗi thư viện và khởi chạy không có pip. Python 3.13.5 được kiểm tra trong môi trường không có site-packages, đọc/xuất đúng ba phiếu mẫu tổng 1.569.000 đồng.

Kiểm thử tự động và trình duyệt thực hiện trên Linux. Bộ khởi chạy `.bat` cần chạy trên máy Windows thực tế để xác nhận tương thích với Python và cấu hình máy của bạn.

## Đóng gói lại bản gọn

```sh
python donggoi.py
```

Công cụ tạo ZIP cạnh thư mục dự án, nén ở mức 9 và kiểm tra tính toàn vẹn trước khi hoàn tất. Có thể chỉ định nơi lưu bằng `python donggoi.py --output Du-an-xuat-kho-gon.zip`. Gói gồm mã nguồn, template, giao diện, OCR, thư viện đi kèm, kiểm thử và năm file dữ liệu mẫu cần thiết. Công cụ đối chiếu mã SHA-256 của OCR/thư viện và không đóng gói cache Python, log khởi chạy, ảnh xem trước hay thông tin đăng nhập.

Kiểm tra hai bộ OCR (SIMD và bộ dùng khi trình duyệt không hỗ trợ SIMD) bằng `python tests/ocr_assets_browser.py --base http://127.0.0.1:8080`. Bài kiểm tra cần Playwright/Chromium; ứng dụng thông thường không cần cài những công cụ kiểm thử này.

Kiểm tra xử lý xuất bị treo, hủy, hết thời gian chờ, giữ chỉnh sửa và tải thủ công bằng `python tests/export_browser_regressions.py --base http://127.0.0.1:8080`. Kiểm tra này cũng dùng Playwright/Chromium.

## Cấu hình Vercel tùy chọn (không cần để chạy local)

Vercel dùng `api/index.py` theo cơ chế không lưu phiên: nguồn Excel đã nhận diện và cấu hình cột được giữ trong trình duyệt, gửi lại khi sửa nhận diện/xuất. Với ảnh, máy chủ nhận văn bản OCR và trả về workbook trung gian; bản thân ảnh không được tải lên máy chủ.

Giới hạn yêu cầu và kết quả trên bản Vercel là 4 MB. Ảnh tối đa 20 MB vẫn có thể đọc trên trình duyệt, nhưng văn bản OCR, workbook trung gian và template gửi đến máy chủ phải nằm trong giới hạn này. Nếu bảng quá lớn, chia thành các lần nhập nhỏ.

Triển khai bằng `npx vercel --prod` khi đã đăng nhập. Đóng gói chỉ `api/`, `static/` (bao gồm `static/ocr/`), `templates/Template.xlsx`, `app.py`, `engine.py`, `image_import.py`, `image_tables.py`, `dependencies.py`, `requirements.txt`, `vercel.json` và cấu hình Python; không đưa dữ liệu mẫu, ảnh khách hàng, kiểm thử hay thông tin đăng nhập vào bản triển khai. Bản anonymous của Vercel là tạm thời và cần nhận dự án vào tài khoản Vercel trước thời hạn mà CLI thông báo.

Kiểm tra ảnh nghiêng/thiếu tương phản bằng OCR thật: `python tests/smart_image_browser.py --base http://127.0.0.1:8080` (cần Playwright/Chromium).
