1. Trang tính CaiDatDeThi (Dùng để quản lý đề, cấu hình thời gian, ngày hạn)

Tạo trang tính đặt tên chính xác là CaiDatDeThi và tạo các cột tại hàng đầu tiên (Dòng 1):
• Cột A: MonHoc (Mã/Tên môn thi, ví dụ: Toan_Chuong1, TiengAnh_GiuaKy)
• Cột B: ThoiGianLamBai (Nhập số phút đếm ngược, nếu nhập 0 hoặc bỏ trống = Không giới hạn thời gian)
• Cột C: NgayBatDau (Định dạng Ngày/Tháng/Năm Giờ:Phút, ví dụ: 01/12/2026 08:00)
• Cột D: NgayKetThuc (Định dạng tương tự ngày bắt đầu, học sinh truy cập ngoài khoảng này sẽ bị khóa đề)
• Cột E: SoCauHoiBoc (Số lượng câu hỏi sẽ bốc ngẫu nhiên từ ngân hàng để tạo đề)

2. Trang tính NganHangDe (Hỗ trợ Trắc nghiệm, Điền khuyết, Hình ảnh)

Tạo trang tính đặt tên là NganHangDe và tạo các cột tại hàng đầu tiên (Dòng 1):
• Cột A: Id (Mã câu hỏi: 1, 2, 3...)
• Cột B: MonHoc (Phải trùng khớp với cột MonHoc bên trang CaiDatDeThi)
• Cột C: LoaiCauHoi (Nhập chữ: MULTIPLE_CHOICE đối với trắc nghiệm, hoặc FILL_BLANK đối với điền khuyết)
• Cột D: NoiDungCauHoi (Nội dung câu hỏi bằng văn bản)
• Cột E: HinhAnh (Dán đường link URL hình ảnh minh họa công thức/đồ thị, nếu không có thì bỏ trống)
• Cột F: DapAnA (Nếu là điền khuyết thì bỏ trống cột F, G, H, I)
• Cột G: DapAnB
• Cột H: DapAnC
• Cột I: DapAnD
• Cột J: DapAnDung (Đối với trắc nghiệm nhập: A, B, C, D. Đối với điền khuyết, nhập đáp án chính xác hoặc các đáp án chấp nhận được cách nhau bằng dấu sổ đứng, ví dụ: điền khuyết|dien khuyet|fill blank)

3. Trang tính KetQua

Tạo trang tính đặt tên là KetQua để lưu điểm. Các cột hàng 1: ThoiGian, TenHocSinh, MonHoc, SoCauDung, TongSoCau, DiemSo

Sử dụng tính năng nhận diện bảng của Google Sheets:
1. Thiết kế file Excel/Word mẫu offline: Bạn tạo một file Excel thiết lập sẵn các cột theo đúng thứ tự: Id, MonHoc, LoaiCauHoi, NoiDungCauHoi, HinhAnh, DapAnA, DapAnB, DapAnC, DapAnD, DapAnDung.
2. Nhập liệu tốc độ cao: Bạn soạn đề thi bằng Excel hoặc Word (dưới dạng bảng tính).
	• Đối với hình ảnh: Bạn hãy upload hình ảnh lên một hosting ảnh miễn phí (như ImgBB hoặc chính Google Drive để chế độ công khai ai cũng xem được) rồi dán link ảnh .jpg/.png vào cột HinhAnh.
3. Import: Bạn chỉ cần mở file Excel lên, bấm Ctrl + A để copy toàn bộ dữ liệu câu hỏi, sau đó mở trang tính NganHangDe trên Google Sheets ra và dán đè (Ctrl + V) nối tiếp vào các hàng bên dưới. Hệ thống GAS sẽ tự động nhận diện và cập nhật ngay lập tức mà không cần xử lý nạp file phức tạp.
