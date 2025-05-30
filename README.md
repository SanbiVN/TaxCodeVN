# TaxCode
 Tra cứu mã số thuế cá nhân và doanh nghiệp cho Excel với nguồn dữ liệu Tổng cục thế và masothue.vn \
 Sử dụng Micosoft Edge WebView2 dựa trên trình IDE twinBasic để viết mã và đóng gói

[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/TaxCodeVN/total.svg)](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v3.26.rar) 

[Nhấn tải TaxCodeVN Rar](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v3.26.rar) \
[Nhấn tải TaxCodeVN Add-in cho Excel](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v3.26.xlam) \
[Nhấn tải TaxCode dành cho Excel và Window cũ hơn](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCode_IE.xlsm)


Với tải tệp rar sẽ gồm tệp Exe mở trực tiếp và cả tệp Add-in. \
Với tải tệp Add-in cài đặt sẽ có nút nhấn để hiện cửa sổ tải thông tin. Nút nhấn nằm tại Tab Ribbon Data (Dữ liệu), mục TaxCode.
Hoặc gõ hàm TaxCodeWindow() vào ô bất kì để mở. \
(Lưu ý: Trước khi cài đặt add-in cần bỏ ```unblock``` tệp xlam trước khi cài đặt vào Excel) \

(Add-in được khóa mật khẩu VBA, vì lí do bảo mật khi ứng dụng hoạt động ở máy tính Công ty)

![image](https://github.com/user-attachments/assets/e53a88c7-fcdd-45fe-a501-b2f2d31ba531)

#### HƯỚNG DẪN 
1. Chọn nguồn dữ liệu (Nguồn tct có nhập captcha bằng tay)
2. Thiết lập tiêu chí tìm kiếm
3. Chọn ô là tiêu đề của cột dữ liệu CCCD/MST trong bảng tính Excel
4. Thiết lập cột dữ liệu trả về cần thiết, nhập số thứ tự vào hộp nhập, để xác định cột ghi kết quả, để trống thì bỏ qua. Thứ tự bắt đầu từ cột MST/CCCD + 1.
5. Nhấn "BẮT ĐẦU" để khởi tạo hoặc tìm kiếm ngay
6. Nguồn là Tổng cục thuế nhập Captcha, nhập đủ 5 ký tự, tự động tải. Nguồn masothue.vn đang tải tự nhiên lỗi, cần nhấn ```Làm mới``` để duyệt ReCaptcha.

Tùy chọn ```Lấy thông tin mã số thuế thứ 2 nếu có``` vì có CCCD có hơn 1 mã số thuế \
Tùy chọn ```Chèn dòng``` nếu có mã số thuế thứ 2


![402415720-60ddbf4e-b8e6-48d6-8840-6a0e9ecae4ee](https://github.com/user-attachments/assets/0797ecb0-e292-41b5-a7c8-f3d5ee82868f)

Khi tải danh sách nhiều mã từ nguồn masothue.vn, do quá trình tải thông tin và ghi thì cần ứng dụng Excel ở chế độ "rảnh tay", nếu không sẽ ảnh hưởng đến quá trình làm việc của bạn. \
Vì vậy hãy mở tệp của bạn ở Excel tiến trình khác, có thể nhấn icon Excel gốc phải phía trên của ứng dụng để mở. Lúc này bạn có thể làm việc tại Excel chính.


### Ví dụ

Bạn cần có danh sách như bảng tính Excel bên dưới \
Mở ứng dụng tải thông tin, tại mục ```Thiết lập cột tìm kiếm```, hãy chọn trang tính tương ứng và nhập ```A1``` vào mục ```Nhập ô tiêu đề``` vì ô ```A1``` là tiêu đề cột MST/CCCD.  \
Thứ tự cột tính từ cột A là 0, cột B sẽ là 1, từ đó thiết lập cột ghi dữ liệu​.

![image](https://github.com/user-attachments/assets/44047f32-45db-49b8-96a8-f0793dc57833)



## Liên hệ của tôi:

#### Zalo 
Zalo: 0384170514
<p align="left">
<img title="@FasterOfficeVBA" src="https://github.com/user-attachments/assets/970644a2-f125-440f-9bd9-2f8888187a22" width="200">
</p>

#### Facebook Messenger
[@Sanbi](https://m.me/he.sanbi)

#### Facebook Page
[FasterOfficeVBA](https://facebook.com/FasterOfficeVBA)

#### Youtube
[@FasterOfficeVBA](https://www.youtube.com/@FasterOfficeVBA)

#### TikTok
[@FasterOfficeVBA](https://www.tiktok.com/@fasterofficevba)
