# Hướng Dẫn Sử Dụng - Quản Lý Kho Gỗ

## 1. Mở File và Bật Macros

1. Mở file `QuanLyKho.xlsm`
2. Khi Excel hỏi về Macros, chọn **Enable Content** hoặc **Enable Macros**
3. Nếu không thấy thông báo, vào **File > Options > Trust Center > Trust Center Settings > Macro Settings** và chọn **Enable all macros**

---

## 2. Giao Diện Chính (Sheet SO DO KHO)

### Sơ đồ kho
- 104 ô kho được đánh số K1 đến K104
- Xếp thành 4 hàng, mỗi hàng 26 ô
- Bên trái có nhãn "VĂN PHÒNG" để định hướng

### Màu sắc ô kho
| Màu | Ý nghĩa |
|-----|---------|
| Xanh lá nhạt | Ô đang mở, có hàng tồn |
| Trắng | Ô đang mở, trống |
| Xám | Ô đang đóng |

### Vùng thông tin (bên phải)
- **Mã Vị Trí**: Hiển thị ô đang chọn
- **Trạng Thái**: Mở hoặc Đóng
- **Danh Sách Tồn**: Các sản phẩm đang có trong ô

---

## 3. Thao Tác Cơ Bản

### 3.1 Xem thông tin ô kho
1. Click vào ô kho bất kỳ (K1-K104) trên sơ đồ
2. Vùng thông tin bên phải sẽ tự động cập nhật

### 3.2 Nhập hàng vào kho
1. Click chọn ô kho muốn nhập
2. Nhấn nút **[Nhập]**
3. Trong form hiện ra:
   - Chọn **Sản phẩm** từ danh sách
   - Nhập **Số tấm**
   - Ghi chú (nếu cần)
4. Nhấn **[Xác nhận]**

### 3.3 Xuất hàng khỏi kho
1. Click chọn ô kho muốn xuất
2. Nhấn nút **[Xuất]**
3. Trong form hiện ra:
   - Chọn **Sản phẩm** (chỉ hiện SP đang có trong ô)
   - Nhập **Số tấm** (không được lớn hơn số tồn)
   - Ghi chú (nếu cần)
4. Nhấn **[Xác nhận]**

### 3.4 Đóng/Mở ô kho
- Nhấn **[Đóng]** để đánh dấu ô không sử dụng
- Nhấn **[Mở]** để mở lại ô đã đóng
- Ô đang đóng không thể nhập/xuất hàng

---

## 4. Quản Lý Dữ Liệu

### Sheet VI TRI
- Danh sách 104 vị trí K1-K104
- Cột **TrangThai**: "Mo" hoặc "Dong" (không dấu)
- Có thể sửa trực tiếp hoặc dùng nút trên SO DO KHO

### Sheet SAN PHAM
- Thêm sản phẩm mới: Điền vào dòng trống
- **MaSP**: Mã duy nhất (VD: SoiTrangAA-17)
- **MaGo**: Tên loại gỗ
- **DoDay**: Độ dày (mm)
- **TrangThai**: "Dung" hoặc "Ngung" (không dấu)

### Sheet PHAT SINH
- Lịch sử tự động ghi khi nhập/xuất
- Không nên sửa trực tiếp
- Cột **SoTamQuyDoi**: Dương khi nhập, âm khi xuất

### Sheet TON KHO
- Tự động tính từ PHAT SINH
- Không nên sửa trực tiếp

---

## 5. Báo Cáo (Sheet BAO CAO)

### Tạo báo cáo
1. Chuyển sang sheet BAO CAO
2. Nhập **Từ ngày** và **Đến ngày**
3. Nhấn nút **[Báo cáo tổng hợp]** hoặc **[Báo cáo chi tiết]**

### Loại báo cáo
- **Tổng hợp**: Tổng nhập/xuất theo từng sản phẩm
- **Chi tiết**: Liệt kê từng phát sinh trong kỳ

---

## 6. Lưu Ý Quan Trọng

1. **Luôn lưu file** sau khi thao tác quan trọng
2. **Backup định kỳ** để tránh mất dữ liệu
3. **Không xóa** các dòng header trong các sheet
4. **Không đổi tên** các sheet (SO DO KHO, VI TRI, SAN PHAM, PHAT SINH, TON KHO, BAO CAO)
5. File phải lưu dạng **.xlsm** để giữ VBA code

---

## 7. Xử Lý Lỗi Thường Gặp

| Lỗi | Nguyên nhân | Cách xử lý |
|-----|-------------|------------|
| Nút không hoạt động | Macros bị tắt | Enable Macros và mở lại file |
| "Ô kho đang đóng" | Ô có Trạng Thái = Đóng | Nhấn nút [Mở] |
| "Không có hàng để xuất" | Ô trống | Kiểm tra lại vị trí |
| Số xuất > số tồn | Nhập sai số lượng | Nhập số <= số tồn hiện tại |
