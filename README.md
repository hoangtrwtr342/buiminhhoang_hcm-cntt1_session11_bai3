# Bài 3 : Chuẩn hóa quy trình tra cứu dữ liệu kho hàng


## Phần 1: Phân tích 4 tham số của VLOOKUP

  - 1. lookup_value (A2): Ô Mã linh kiện cần đem đi dò. Để tham chiếu tương đối (không có dấu $) để khi kéo chuột xuống, Excel tự đổi thành A3, A4....

  - 2. table_array (Danh_Muc!$A$2:$C$100): Vùng bảng dữ liệu gốc để tra cứu. Bắt buộc bấm F4 để khóa $A$2:$C$100 — nếu quên dấu $, khi kéo công thức vùng này sẽ bị tụt dòng dần (A3:C101, A4:C102), làm mất dấu các mã phía trên và sinh lỗi #N/A hàng loạt (Bẫy 1).

  - 3. col_index_num (2 hoặc 3): Vị trí cột muốn lấy dữ liệu ra, tính từ trái qua phải của bảng dò (cột 2 là Tên linh kiện, cột 3 là Đơn giá).

  - 4. range_lookup (FALSE hoặc 0): Chế độ tìm kiếm chính xác tuyệt đối. Vì VLOOKUP so khớp từng ký tự, nếu bảng tra có mã bị dính khoảng trắng như "LK01 " (Bẫy 2), hàm sẽ không nhận diện được và báo lỗi. Do đó, cần dùng hàm TRIM làm sạch cột mã trước khi tra.


##Phần 2: Công thức triển khai hoàn chỉnh

1. Hàm IFERROR bọc bên ngoài chỉ đóng vai trò dự phòng: nếu không tìm thấy mã, nó sẽ trả về kết quả dễ nhìn thay vì báo lỗi đỏ #N/A.

2. Lấy Tên linh kiện (nhập tại B2 của Sheet Đơn hàng):
=IFERROR(VLOOKUP(A2, Danh_Muc!$A$2:$C$100, 2, FALSE), "Không tìm thấy")

3. Lấy Đơn giá (nhập tại C2 của Sheet Đơn hàng):
=IFERROR(VLOOKUP(A2, Danh_Muc!$A$2:$C$100, 3, FALSE), 0)
