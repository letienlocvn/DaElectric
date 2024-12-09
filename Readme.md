Tôi và bạn sẽ một application java.  
File này là một bảng tính excel ghi chép hóa đơn tiền điện. Nó không phải là hóa đơn.

Bạn vui lòng nắm hoàn cảnh của tôi như sau:
- Tôi có một kế toán đã hoàn thành file excel này và kế toán viên đó gửi lại cho tôi.
- Nhiệm vụ của tôi là phải thay thế nội dung tương ứng của file excel này vào một file hóa đơn excel mới (file này đã được thiết kế để in ra).

Mục đích chính của ứng dụng:
- Tiết kiệm thời gian tôi phải nhập thủ công từ file excel.
- Ứng dụng này sẽ lưu dữ liệu đã đọc được vào một file excel mới (file hóa đơn tiền điện).
  Cấu trúc file Excel:
  - Có những cột STT	Họ Tên	 Chỉ số cũ 	 Chỉ số mới 	Số trong tháng	Đơn giá	Thành tiền	Công ghi điện	Tổng thanh toán
  - File này chỉ có một sheet thôi.

Chức năng mong muốn
- Dựa trên thông tin của file hóa đơn dùng để in ra. Thay thế các số tương ứng từ file quản lí tiền điện sang file hóa đơn.
  Ví dụ: Chỉ số cũ, chỉ số mới, đơn giá.
- Không cần giao diện người dùng (phần này có thể phát triển sau)

Công nghệ và thư viện
Bất kì công nghệ gì tốt để giải quyết tốt vấn đề và hiệu quả. Dùng trên laptop

Đầu vào và đầu ra
- Đầu vào: File Excel này sẽ được nhập thủ công bởi kế toán.
- Đầu ra: xuất ra file Excel khác hoặc lưu vào database để mapping sang một file excel hóa khác.


1. Cấu trúc file hóa đơn
   Cấu trúc của một file hóa đơn được thể hiện trong file HoaDon2023.xlsx
   Phải chi tiết được các nội dung được mapping như:
- Chỉ số cũ:
- Chỉ số mới:
- Đơn giá:
- Tổng tiền thanh toán

Các trường cố định là công ghi điện. hoặc nội dung description của liên 2. (* Vui lòng thanh toán tiền trong vòng 2 ngày kể từ khi nhận được hóa đơn này. Ngày 5 hàng tháng là hạn chót nộp tiền
*  Quy định của hợp đồng của điện lực áp dụng cho trường hợp không nộp đúng hạn, sẽ tính phạt 8% / tháng cho số tiền nộp trễ.)

2. Mapping dữ liệu
   Tất cả các dòng sẽ được gộp vào một file hóa đơn duy nhất
   Có cần hỗ trợ tính toán dữ liệu không? Không, tại vì trong file excel đó đã có công thức tính rồi. Nhiệm vụ của mình là thay thế chỉ số cũ và chỉ số mới, đơn giá. Tự khắc các thành phần khác sẽ tự động cập nhật.

3. Lưu đầu ra
   Sẽ do người dùng chỉ định để lưu.
- Định dạng file hóa đơn đầu ra sẽ giữ nguyên (Excel), hay có cần chuyển đổi sang định dạng khác (PDF chẳng hạn): Nên lựa chọn như thế nào, vì tôi cũng chưa biết cách tốt nhất? Nhưng cuối cùng mục đích là dùng để in ra được file hóa đơn này.

4. Xử lý lỗi
   Nếu dữ liệu từ file đầu vào bị thiếu hoặc không khớp, Tùy từng trường hợp. Ở phần này sẽ phát triển dựa trên các yêu cầu. Nhưng trước mắt hãy cố gắng xử lý hợp lý logic nhất có thể.
   Có cần nhật ký (log) để báo cáo các dòng lỗi hoặc không thể xử lý không: Có

