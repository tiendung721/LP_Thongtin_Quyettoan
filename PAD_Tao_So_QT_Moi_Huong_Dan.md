# Tạo nhiều số quyết toán mới từ PM

1. Mở PAD flow **Test - tao dong quyet toan moi**, xóa các action cũ trong subflow Main, rồi dán toàn bộ nội dung [PAD_Tao_So_QT_Moi.txt](PAD_Tao_So_QT_Moi.txt). Flow mới tự tạo biến `CreateCount` bằng action **Set variable**; **không cần tạo Input variable** trong bảng Variables.
2. Lưu flow. Trong PM, bấm **Tạo số quyết toán mới** và nhập số lượng từ 1 đến 100. PM ghi yêu cầu vào `runtime/create_new_request.json` trước khi mở PAD; PAD đọc yêu cầu, xóa file sau khi đọc và lặp đúng số lượt.
3. Nếu chuyển thư mục dự án, sửa `ProjectRoot` ở action đầu tiên trong PAD cho đúng đường dẫn mới.
4. Chạy thử với 1, rồi 2 và kiểm tra số dòng thực tế trên website trước khi chạy số lượng lớn. Flow dùng tọa độ màn hình theo bản gốc và chưa kiểm tra phản hồi lưu của website; cần xác nhận độ phân giải, trạng thái trang và thao tác lưu trên máy chạy PAD.

Flow chỉ đọc yêu cầu do PM tạo. Bấm **Run** trực tiếp trong PAD khi không có file yêu cầu sẽ hiện thông báo thiếu yêu cầu và dừng.
