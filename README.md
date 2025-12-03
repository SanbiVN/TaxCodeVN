# TaxCode - Tra cứu mã số thuế
 Tra cứu mã số thuế cá nhân và doanh nghiệp cho Excel từ nguồn dữ liệu Tổng cục thuế \
 Sử dụng Micosoft Edge WebView2 dựa trên trình IDE twinBasic để viết mã và đóng gói

 <img width="1548" height="137" alt="image" src="https://github.com/user-attachments/assets/ed80cbb3-92fc-4884-93b4-49a954328d21" />

--------------------------------------------------------------
## TẢI XUỐNG
<!-- items that need to be updated release to release -->
[ptUserAddin]: https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v4.2.rar
[ptUserWebview2]: https://go.microsoft.com/fwlink/p/?LinkId=2124703

|  Thông tin   | Tải xuống | Lượt tải |
|----------------|----------|----------|
| TaxCodeVN Add-in | [TaxCodeVN_v4.2.rar][ptUserAddin] | [![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/TaxCodeVN/total.svg)]([ptUserAddin])  |


<!-- 
[Nhấn tải TaxCodeVN Add-in](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v4.2.rar) 
[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/TaxCodeVN/total.svg)](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v4.2.rar) 
 -->

(Lưu ý: Sau khi giải nén tệp tải về, trước khi cài đặt add-in cần vào Properies của tệp trong thư mục bỏ ```Unblock``` tệp xlam trước khi cài đặt vào Excel nếu có) 

(Add-in được khóa mật khẩu VBA, vì lí do bảo mật khi ứng dụng hoạt động ở máy tính Công ty)



--------------------------------------------------------------
## HƯỚNG DẪN 

### Tra cứu liên tục tự động
1. Tạo một tên thiết lập mới dành cho trang tính hiện tại chứa danh sách cần tìm kiếm.
2. Chọn mục tìm kiếm là Doanh nghiệp/Cá Nhân.
3. Nhập ô bắt đầu, Thiết lập cột ghi dữ liệu tính từ ô bắt đầu, nhập số thứ tự vào hộp nhập, để xác định cột ghi kết quả, để 0 thì bỏ qua. 
4. Lưu thiết lập, nhấn "BẮT ĐẦU" để khởi tạo hoặc tìm kiếm ngay
5. Nguồn là Tổng cục thuế nhập Captcha, nhập đủ 5 ký tự, tự động tải.

### Tra cứu nhanh​
Mục tra cứu nhanh giúp tra cứu nhanh chóng khi cần tra cứu một dòng dữ liệu, hoặc một thông tin.​


--------------------------------------------------------------
Sau khi nhấn bắt đầu một giao diện cửa sổ như sau sẽ xuất hiện, một cơ chế sửa captcha nhanh chóng, với các nút 1, 2, 3, 4, 5 là thứ tự chọn vị trí, sau khi chọn, có thể nhấn phím vật lý hoặc dùng chuột nhấn vào bàn phím ảo có sẵn trên màn hình.

>  <img width="490" height="179" alt="1764294767315" src="https://github.com/user-attachments/assets/65b27735-cfd2-4015-b046-d4738055cdd8" />


--------------------------------------------------------------
## LƯU Ý CHUNG
***Ứng dụng hiện tại chỉ hoạt động trên HĐH Windows (10,11+) có thể cài đặt WebView2 Runtime của microsoft, dự án sẽ hỗ trợ Windows 7 trong tương lai.

***Lưu ý về trang tra cứu TCT, Khi tra cứu trang sẽ mắc lỗi với api tracking, cookie và captcha không đồng nhất, gây lỗi nhập captcha mặc dù đúng nhiều lần nhưng báo sai, không trả về dữ liệu. Lỗi này do trang nguồn, không phải lỗi tại ứng dụng tra cứu. Trang nguồn dùng công nghệ F5 (BIG-IP, NGINX, WAF) để chóng lại các cuộc tấn công, và các hành vi truy vấn nhanh bởi máy. Điều này gây khó khăn hơn trong việc tra cứu.

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
