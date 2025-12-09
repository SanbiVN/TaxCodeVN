# TaxCode - Tra cứu mã số thuế
Tra cứu mã số thuế cá nhân và doanh nghiệp cho Excel từ nguồn dữ liệu Tổng cục thuế ```https://tracuunnt.gdt.gov.vn/```\
Add-in sử dụng công nghệ Web nhân Chromium và WebView2 Runtime để điều khiển tải dữ liệu.

(Hình ảnh hiển thị trên thanh ribbon Excel)​
<img width="1585" height="136" alt="image" src="https://github.com/user-attachments/assets/4ea43ad1-7c2d-4e4d-8214-0c1063e96b08" />


------------------------------------------------
## TẢI XUỐNG
<!-- items that need to be updated release to release -->
[ptUserAddin]: https://github.com/SanbiVN/TaxCodeVN/releases/download/v4.10/TaxCode_v4.10.zip
[ptUserWebview2]: https://go.microsoft.com/fwlink/p/?LinkId=2124703

|  Thông tin   | Tải xuống | Lượt tải |
|----------------|----------|----------|
| TaxCodeVN Add-in | [TaxCode_v4.10.zip][ptUserAddin] | [![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/TaxCodeVN/total.svg)](https://github.com/SanbiVN/TaxCodeVN/releases/download/v4.10/TaxCode_v4.10.zip)  |


<!-- 
[Nhấn tải TaxCodeVN Add-in](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v4.2.rar) 
[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/TaxCodeVN/total.svg)](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v4.2.rar) 
 -->
 Ứng dụng chỉ hoạt động trên HĐH Windows 7 trở lên (Không bao gồm máy sử dụng chip ARM). \
(Add-in được khóa mật khẩu VBA, vì lí do bảo mật khi ứng dụng hoạt động ở máy tính Công ty)

--------------------------------------------------------------

## CÀI ĐẶT
 Tệp Add-in xlam để cài đặt vào Excel 
> Sau khi giải nén, vào thông tin tệp ngoài thư mục bỏ unblock tệp trước khi cài đặt nếu có \
	Cài đặt: Thực hiện cài đặt như một Add-in bình thường \
	Trong thẻ Deverloper chọn Excel Add-ins, sau đó chọn nút Browse... vào thư mục chứa tệp Add-in, đánh dấu Add-in vừa thêm và chọn nút OK \
	(Mở thẻ Deverloper: chuột phải vào thanh Ribbon, chọn Customize the Ribbon) \
> Ứng dụng chỉ cần cài đặt một lần duy nhất, còn lại tự động kiểm tra và tải cài đặt phiên bản mới
 
--------------------------------------------------------------
## HƯỚNG DẪN 

### Tra cứu liên tục tự động
1. Tạo một tên thiết lập mới dành cho trang tính hiện tại chứa danh sách cần tìm kiếm.
2. Chọn mục tìm kiếm là Doanh nghiệp/Cá Nhân.
3. Nhập ô bắt đầu, Thiết lập cột ghi dữ liệu tính từ ô bắt đầu, nhập số thứ tự vào hộp nhập, để xác định cột ghi kết quả, để 0 thì bỏ qua. 
4. Lưu thiết lập, nhấn "BẮT ĐẦU" để khởi tạo hoặc tìm kiếm ngay
5. Nguồn là Tổng cục thuế nhập Captcha, nhập đủ 5 ký tự, tự động tải.
> ***Lưu ý: Ô bắt đầu là tiêu đề cột dữ liệu, tiêu đề không được để trống​

Ảnh bên dưới, chọn ô bắt đầu là B1, nếu chọn cột C ghi dữ liệu thì ô C1 cũng không để trống tiêu đề

> ![taxcode_list](https://github.com/user-attachments/assets/54b711b2-a21a-4891-bd46-9859770dd133)


### Tra cứu nhanh​
> Mục tra cứu nhanh giúp tra cứu nhanh chóng khi cần tra cứu một dòng dữ liệu, hoặc một thông tin.​


--------------------------------------------------------------
Sau khi nhấn bắt đầu một giao diện cửa sổ như sau sẽ xuất hiện, một cơ chế sửa captcha nhanh chóng, với các nút 1, 2, 3, 4, 5 là thứ tự chọn vị trí, sau khi chọn, có thể nhấn phím vật lý hoặc dùng chuột nhấn vào bàn phím ảo có sẵn trên màn hình.

>  <img width="490" height="179" alt="1764294767315" src="https://github.com/user-attachments/assets/65b27735-cfd2-4015-b046-d4738055cdd8" />





--------------------------------------------------------------
## LƯU Ý CHUNG
Lưu ý về trang tra cứu TCT, Khi tra cứu trang sẽ mắc lỗi với api tracking, cookie và captcha không đồng nhất, gây lỗi nhập captcha mặc dù đúng nhiều lần nhưng báo sai, không trả về dữ liệu. Lỗi này do trang nguồn, không phải lỗi tại ứng dụng tra cứu. Trang nguồn dùng công nghệ F5 (BIG-IP, NGINX, WAF) để chóng lại các cuộc tấn công, và các hành vi truy vấn nhanh bởi máy. Điều này gây khó khăn hơn trong việc tra cứu.

(chỉ với một trang tra cứu thông tin của quốc gia, nhưng vấn đề gây ra cho người tra cứu quá căng, tôi đã cố hết sức)

--------------------------------------------------------------
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
