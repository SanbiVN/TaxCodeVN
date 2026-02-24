# TaxCode - Tra cứu mã số thuế
Tra cứu mã số thuế cá nhân và doanh nghiệp cho Excel từ nguồn dữ liệu Tổng cục thuế ```https://tracuunnt.gdt.gov.vn/```\
Và từ nguồn bên ngoài như masothue và thuvienphapluat\
Add-in sử dụng công nghệ Web nhân Chromium và WebView2 Runtime để điều khiển tải dữ liệu.

(Hình ảnh hiển thị trên thanh ribbon Excel)​
<img width="1578" height="175" alt="image" src="https://github.com/user-attachments/assets/d9382a0b-9a60-47ba-8cfe-1f7cc5bdfa6d" />

### DANH MỤC
- [TẢI XUỐNG](#tải-xuống)
- [HƯỚNG DẪN CÀI ĐẶT](#hướng-dẫn-cài-đặt)
- [HƯỚNG DẪN SỬ DỤNG](#hướng-dẫn-sử-dụng)
  - [Tra cứu liên tục tự động](#tra-cứu-liên-tục-tự-động)
  - [Tra cứu nhanh](#tra-cứu-nhanh)
- [CÁC NÚT CHỨC NĂNG KHÁC](#các-nút-chức-năng-khác)
- [LƯU Ý CHUNG](#lưu-ý-chung)
- [Liên hệ](#liên-hệ)

  

------------------------------------------------
## TẢI XUỐNG
<!-- items that need to be updated release to release -->
[ptUserAddin]: https://github.com/SanbiVN/TaxCodeVN/releases/download/v4.36/TaxCode_v4.36.zip
[ptUserWebview2]: https://go.microsoft.com/fwlink/p/?LinkId=2124703

|  Thông tin   | Tải xuống | Lượt tải |
|--------------|-----------|----------|
| TaxCode Add-in Excel | [TaxCode_v4.36.zip][ptUserAddin] | [![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/TaxCodeVN/total.svg)](https://github.com/SanbiVN/TaxCodeVN/releases/download/v4.36/TaxCode_v4.36.zip)  |


<!-- 
[Nhấn tải TaxCodeVN Add-in](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v4.2.rar) 
[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/TaxCodeVN/total.svg)](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v4.2.rar) 
 -->
 Ứng dụng chỉ hoạt động trên HĐH Windows 7 trở lên (Không bao gồm máy sử dụng chip ARM). \
(Add-in được khóa mật khẩu VBA, vì lí do bảo mật khi ứng dụng hoạt động ở máy tính Công ty)

--------------------------------------------------------------

## HƯỚNG DẪN CÀI ĐẶT
Tệp Add-in xlam để cài đặt vào Excel, sau khi cài đặt thì giao diện sử dụng hiển thị trên thanh Ribbon với tên ```TaxTCT```. Ứng dụng chỉ cần cài đặt một lần duy nhất, còn lại tự động kiểm tra và tải cài đặt phiên bản mới. \
Giải nén vào một thư mục được đặt tên phù hợp, sau khi giải nén, vào thông tin tệp ngoài thư mục bỏ unblock tệp trước khi cài đặt nếu có.

> <img width="363" height="478" alt="image" src="https://github.com/user-attachments/assets/359bee94-f4b7-4fa2-bc48-23ab7723fd7b" />

**Cách 1:** 
- Mở trực tiếp Add-in hoặc nhấn chuột vào tệp để mở, trong Excel cần **```Enabled Macro```** để chương trình hoạt động. 
- Nếu chương trình chưa cài đặt khởi động cùng Excel, khi nhấn **BẮT ĐẦU** chương trình sẽ hỏi có cài đặt khởi động vào Excel không?

**Cách 2:** Thực hiện cài đặt Add-in bằng tay: 
  - Nếu chưa có tab Deverloper hiển thị trên thanh Ribbon (Thanh công cụ): nhấn chuột phải vào thanh Ribbon, chọn **```Customize the Ribbon```**.
  - Trong thẻ Deverloper chọn **```Excel Add-ins```**, sau đó chọn nút **```Browse...```** vào thư mục chứa tệp Add-in, đánh dấu Add-in vừa thêm và chọn nút OK 
  - Nếu đã cài đặt vào Excel, nhưng mỗi khi mở ứng dụng không thấy trên thanh Ribbon, thì vào **```Task Manager```** cần End Task ứng dụng Excel chạy ngầm.

 Nếu ứng dụng bị chặn không cho chạy macro thì hãy vào Cài đặt Excel, vào Trust Center, vào tạo đường dẫn thư mục an toàn cho thư mục chứa add-in tải về.
 
--------------------------------------------------------------
## HƯỚNG DẪN SỬ DỤNG
(Trên các nút Ribbon đều có típ hướng dẫn, chỉ cần rê chuột vào nút để xem)
### Tra cứu liên tục tự động
- **Bước 1**:
  + Tạo một tên thiết lập mới dành cho trang tính hiện tại chứa danh sách cần tra cứu tải về. Thiết lập này dùng để sử dụng lại về sau này khi mở lại trang tính đã thiết lập.
  + Chọn mục tìm kiếm là Doanh nghiệp/Cá Nhân. 
  + Nhập ô bắt đầu của danh sách tra cứu. 
  + Tích chọn nguồn tải dữ liệu. 
  + Tích chọn vào hộp **```Tải liên tục```** nếu là nguồn Tổng cục thuế. 
  + Chọn cách gợi ý captcha nếu cần thiết nếu là nguồn Tổng cục thuế.
> <img width="373" height="111" alt="image" src="https://github.com/user-attachments/assets/b785e77a-f6e0-44a8-bc8d-c09c599dc77f" />
- **Bước 2**: Thiết lập cột ghi dữ liệu, có thể nhập cột ký tự A-Z, hoặc nhập số thứ tự tính từ ô bắt đầu, để xác định cột ghi kết quả, để trống hoặc 0 thì bỏ qua. 
> <img width="296" height="109" alt="image" src="https://github.com/user-attachments/assets/96736e04-f075-44ec-8c77-917db0bb4b04" />

> Ảnh bên dưới minh họa, chọn ô bắt đầu là B1, nếu chọn cột C ghi dữ liệu thì ô C1 cũng không để trống tiêu đề

> ![taxcode_list](https://github.com/user-attachments/assets/54b711b2-a21a-4891-bd46-9859770dd133)

- **Bước 3**: Nếu muốn tải chi nhánh, tích chọn vào **```Thông tin chi nhánh```**, chọn thời gian ```độ trễ``` mỗi lượt tải.
> Mặc định chương trình sử dụng WebView2 để tải trang, có thể tùy chọn dùng trình duyệt đã cài trên máy tính nếu WebView không hoạt động.
> Tùy chọn trình duyệt chuẩn khi sử dụng ở máy tính Công Ty, vì các cách khác có thể bị chặn.
>> <img width="175" height="111" alt="image" src="https://github.com/user-attachments/assets/78dffa25-f271-41fc-b82e-dc56aeace631" />

- **Bước 4**: Nhấn lưu thiết lập và nhấn **```BẮT ĐẦU```** để tự động tra cứu và tải về ghi vào trang tính.
  Luôn luôn cần nhấn lưu lại mỗi khi đổi thông số trước khi nhấn **BẮT ĐẦU**:
> <img width="376" height="129" alt="image" src="https://github.com/user-attachments/assets/556e7020-1d34-42e4-891a-fee52fa78146" />

  
##### Lưu ý: 
+ Nếu ô bắt đầu là tiêu đề của cột dữ liệu, thì các tiêu đề của các cột ghi dữ liệu không trống​.
+ Nếu muốn tải lại thông tin thì xóa dữ liệu có sẵn và bất đầu lại.
+ Khi ứng dụng hoạt động mở Excel ở tiến trình khác (Cửa sổ mới khác tiến khác) để chương trình tải và ghi, không làm ảnh hưởng đến quá trình làm việc với Excel.

### Tra cứu nhanh​
> <img width="210" height="108" alt="image" src="https://github.com/user-attachments/assets/e8a5b318-3d6e-4ae3-9aa9-66a74b272a01" />
- Mục tra cứu nhanh giúp tra cứu nhanh chóng khi cần tra cứu một dòng dữ liệu, hoặc một thông tin.​
- Nhập một thông tin vào hộp để tra cứu nhanh. Không ghi dữ liệu.
- Hoặc **```Tìm tại ô chọn```**, Nếu muốn ghi kết quả, tích chọn hộp kiểm **```Ghi kết quả```**

--------------------------------------------------------------
Sau khi nhấn bắt đầu một giao diện cửa sổ như sau sẽ xuất hiện, một cơ chế sửa captcha nhanh chóng, với các nút 1, 2, 3, 4, 5 là thứ tự chọn vị trí, sau khi chọn, có thể nhấn phím vật lý hoặc dùng chuột nhấn vào bàn phím ảo có sẵn trên màn hình.

>  <img width="490" height="179" alt="1764294767315" src="https://github.com/user-attachments/assets/65b27735-cfd2-4015-b046-d4738055cdd8" />


## CÁC NÚT CHỨC NĂNG KHÁC
> <img width="188" height="105" alt="image" src="https://github.com/user-attachments/assets/159ba15b-6d2e-4168-966d-d6477f5c31a5" />
1. **Cập nhật add-in**: nhấn nút để thực hiện kiểm tra hoặc cập nhật Add-in nhanh chóng.
2. **Mở tiến trình**: nhấn để xem hướng dẫn mở Excel ở tiến trình khác.
3. **Đặt lại**: Nhấn để xóa tất cả thiết lập.
4. **Đóng**: Nhấn để đóng cửa sổ tra cứu, trình cập nhật chạy ngầm hoặc thoát add-in nếu không cần sử dụng nữa.

--------------------------------------------------------------
## LƯU Ý CHUNG
Lưu ý về trang tra cứu TCT, Khi tra cứu trang sẽ mắc lỗi với api tracking, cookie và captcha không đồng nhất, gây lỗi nhập captcha mặc dù đúng nhiều lần nhưng báo sai, không trả về dữ liệu. Lỗi này do trang nguồn, không phải lỗi tại ứng dụng tra cứu. Trang nguồn dùng công nghệ F5 (BIG-IP, NGINX, WAF) để chóng lại các cuộc tấn công, và các hành vi truy vấn nhanh bởi máy. Điều này gây khó khăn hơn trong việc tra cứu.

(chỉ với một trang tra cứu thông tin của quốc gia, nhưng vấn đề gây ra cho người tra cứu quá căng, tôi đã cố hết sức)

--------------------------------------------------------------
## Liên hệ

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
