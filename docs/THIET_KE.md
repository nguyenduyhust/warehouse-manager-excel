# Tài Liệu Thiết Kế: Hệ Thống Quản Lý Kho Gỗ Excel

## 1. Tổng Quan

Hệ thống quản lý kho gỗ sử dụng Excel VBA, bao gồm 6 sheets chính để theo dõi vị trí, sản phẩm, xuất nhập và tồn kho.

---

## 1.1 Quy Ước Đặt Tên

| Loại | Quy ước | Ví dụ |
|------|---------|-------|
| Tên Sheet | VIẾT HOA, có cách | SO DO KHO, VI TRI, SAN PHAM |
| Tên Cột (Header) | PascalCase, không dấu | MaViTri, TrangThai, SoTamQuyDoi |
| Giá trị dữ liệu | Có dấu tiếng Việt | "Mở", "Đóng", "Nhập", "Xuất" |
| Mã vị trí | K + số thứ tự | K1, K2, ... K104 |
| Mã sản phẩm | TênGỗViếtTắt-ĐộDày | SoiTrangAA-17, CotTho-9 |

---

## 2. Cấu Trúc Các Sheet

### 2.1 Sheet SO DO KHO (Giao diện chính)

**Bố cục:**
- Cột A-AA: Sơ đồ kho 104 ô (K1-K104), xếp 4 hàng x 26 cột
- Cột AC-AG: Vùng thông tin ô kho

**Sơ đồ kho (4 hàng):**
```
Hàng 2: K1  - K26  (cột A-Z)
Hàng 3: K27 - K52  (cột A-Z)
Hàng 5: K53 - K78  (cột A-Z)
Hàng 6: K79 - K104 (cột A-Z)
```
- Hàng 4: Mũi tên (→) chỉ hướng đi lại giữa các ô
- Cột bên trái: Nhãn "VĂN PHÒNG"

**Vùng THÔNG TIN Ô KHO (cột AC-AG):**
| Dòng | Nội dung |
|------|----------|
| 2 | Mã Vị Trí: [Giá trị] |
| 3 | Trạng Thái: [Mở/Đóng] |
| 4 | Nút: [Nhập] [Xuất] [Đóng] [Mở] |
| 6 | Tiêu đề: Danh Sách Tồn |
| 7 | Header: MaSP | MaGo | DoDay | Ton |
| 8+ | Dữ liệu tồn kho của ô được chọn |

**Màu sắc ô kho:**
- Xanh lá: Ô đang mở, có hàng
- Trắng/Xám nhạt: Ô đang mở, trống
- Đỏ/Xám đậm: Ô đóng

---

### 2.2 Sheet VI TRI (Danh mục vị trí)

| Cột | Tên | Mô tả |
|-----|-----|-------|
| A | MaViTri | Mã vị trí (K1-K104) |
| B | TrangThai | "Mở" hoặc "Đóng" |
| C | GhiChu | Ghi chú tùy chọn |

**Dữ liệu khởi tạo:** 104 dòng (K1 đến K104), mặc định TrangThai = "Mở"

---

### 2.3 Sheet SAN PHAM (Danh mục sản phẩm)

| Cột | Tên | Mô tả |
|-----|-----|-------|
| A | MaSP | Mã sản phẩm (VD: SoiTrangAA-17) |
| B | MaGo | Tên loại gỗ (VD: Sồi trắng AA) |
| C | DoDay | Độ dày (mm) |
| D | TrangThai | "Dùng" hoặc "Ngừng" |
| E | GhiChu | Ghi chú tùy chọn |

**Quy tắc đặt MaSP:** `[TênGỗViếtTắt]-[ĐộDày]`
- SoiTrangAA-17 = Sồi trắng AA, 17mm
- CotTho-9 = Cốt thô, 9mm

---

### 2.4 Sheet PHAT SINH (Lịch sử xuất nhập)

| Cột | Tên | Mô tả |
|-----|-----|-------|
| A | Ngay | Ngày giao dịch (dd/mm/yyyy) |
| B | Gio | Giờ giao dịch (hh:mm:ss) |
| C | Loai | "Nhập" hoặc "Xuất" |
| D | MaViTri | Mã vị trí (K1-K104) |
| E | MaSP | Mã sản phẩm |
| F | SoTam | Số tấm thực tế |
| G | SoTamQuyDoi | Số tấm quy đổi (tính theo công thức) |
| H | MaGo | Mã gỗ (tự động từ MaSP) |
| I | DoDay | Độ dày (tự động từ MaSP) |
| J | GhiChu | Ghi chú tùy chọn |

**Công thức SoTamQuyDoi:**
```
Nếu Loai = "Nhập" thì SoTamQuyDoi = SoTam (số dương)
Nếu Loai = "Xuất" thì SoTamQuyDoi = -SoTam (số âm)
```
> Giúp tính tồn kho dễ dàng bằng SUM(SoTamQuyDoi)

---

### 2.5 Sheet TON KHO (Tồn kho theo vị trí)

| Cột | Tên | Mô tả |
|-----|-----|-------|
| A | MaViTri | Mã vị trí |
| B | MaSP | Mã sản phẩm |
| C | MaGo | Mã gỗ |
| D | DoDay | Độ dày |
| E | SoTam | Số tấm tồn |

**Cập nhật:** Tự động tính từ PHAT SINH (Tổng Nhập - Tổng Xuất)

**Lưu ý:** Mỗi ô kho có thể chứa nhiều loại sản phẩm khác nhau

---

### 2.6 Sheet BAO CAO (Báo cáo)

**Bộ lọc:**
- Chọn loại báo cáo: Ngày / Tháng / Năm
- Chọn khoảng thời gian: Từ ngày - Đến ngày

**Nội dung báo cáo:**
1. Tổng nhập theo sản phẩm
2. Tổng xuất theo sản phẩm
3. Tồn kho cuối kỳ
4. Chi tiết phát sinh trong kỳ

---

## 3. Chức Năng VBA

### 3.1 Sự kiện Click ô kho (Worksheet_SelectionChange)

Khi click vào ô K1-K104 trên SO DO KHO:
1. Cập nhật vùng thông tin với MaViTri được chọn
2. Hiển thị TrangThai từ sheet VI TRI
3. Load danh sách tồn từ sheet TON KHO
4. Enable/Disable nút phù hợp (Đóng/Mở tùy trạng thái)

### 3.2 Nút NHẬP

1. Mở form nhập liệu:
   - MaViTri: Hiển thị (không sửa)
   - MaSP: Dropdown từ SAN PHAM (chỉ TrangThai = "Dùng")
   - SoTam: Nhập số
   - GhiChu: Tùy chọn
2. Validate: Ô phải đang "Mở"
3. Ghi vào PHAT SINH với Loai = "Nhập"
4. Cập nhật TON KHO
5. Refresh sơ đồ và thông tin

### 3.3 Nút XUẤT

1. Mở form xuất liệu:
   - MaViTri: Hiển thị
   - MaSP: Dropdown chỉ hiện SP đang tồn tại vị trí này
   - SoTam: Nhập số (validate <= SoTam tồn)
   - GhiChu: Tùy chọn
2. Validate: Số xuất không được > số tồn
3. Ghi vào PHAT SINH với Loai = "Xuất"
4. Cập nhật TON KHO
5. Refresh sơ đồ và thông tin

### 3.4 Nút ĐÓNG

1. Kiểm tra ô đang "Mở"
2. Cập nhật TrangThai = "Đóng" trong VI TRI
3. Đổi màu ô trên sơ đồ
4. Refresh thông tin

### 3.5 Nút MỞ

1. Kiểm tra ô đang "Đóng"
2. Cập nhật TrangThai = "Mở" trong VI TRI
3. Đổi màu ô trên sơ đồ
4. Refresh thông tin

### 3.6 Cập nhật màu sơ đồ

```
Sub UpdateWarehouseColors()
    For each ô K1-K104:
        If TrangThai = "Đóng" Then
            Màu = Xám
        ElseIf SoTam > 0 Then
            Màu = Xanh lá
        Else
            Màu = Trắng
        End If
    Next
End Sub
```

---

## 4. Quyết Định Thiết Kế (Đã Xác Nhận)

| Câu hỏi | Quyết định |
|---------|------------|
| Công thức SoTamQuyDoi | Nhập = +SoTam, Xuất = -SoTam |
| Một ô kho chứa nhiều SP? | CÓ |
| Loại nút | Shape với Macro |
| Form nhập/xuất | UserForm VBA |
| Loại báo cáo | Tổng hợp + Chi tiết theo SP |
| Màu sắc | Xanh lá (có hàng), Trắng (trống), Xám (đóng) |

---

## 5. Kế Hoạch Triển Khai

| Bước | Nội dung |
|------|----------|
| 1 | Tạo cấu trúc các sheet với header và dữ liệu mẫu |
| 2 | Viết VBA cho sự kiện click ô kho |
| 3 | Tạo UserForm nhập/xuất |
| 4 | Viết logic xử lý nhập/xuất/đóng/mở |
| 5 | Tạo báo cáo |
| 6 | Test và hoàn thiện |

---

## 6. Ghi Chú Kỹ Thuật

- File sẽ lưu dạng `.xlsm` (Excel Macro-Enabled)
- VBA code sẽ được tổ chức trong các Module riêng biệt
- Sử dụng Named Ranges cho các vùng dữ liệu quan trọng
