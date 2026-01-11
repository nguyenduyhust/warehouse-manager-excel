# Hướng Dẫn Import VBA Code

## Danh Sách File

| File | Loại | Import vào |
|------|------|------------|
| `modMain.bas` | Module | Insert > Module |
| `modBaoCao.bas` | Module | Insert > Module |
| `Sheet_SODOKHO.cls` | Sheet Code | Sheet1 (SO DO KHO) |
| `ThisWorkbook.cls` | Workbook Code | ThisWorkbook |
| `frmNhapXuat.frm` | UserForm | Insert > UserForm |

---

## Cách Import

### Bước 1: Chuẩn bị file Excel

1. Mở `QuanLyKho.xlsx`
2. **File → Save As** → Chọn `Excel Macro-Enabled Workbook (*.xlsm)`
3. Lưu thành `QuanLyKho.xlsm`

### Bước 2: Mở VBA Editor

Nhấn **Alt + F11**

### Bước 3: Import Modules

**Cách 1 - Import trực tiếp (nếu Excel hỗ trợ):**
1. **File → Import File**
2. Chọn file `modMain.bas`
3. Lặp lại với `modBaoCao.bas`

**Cách 2 - Copy/Paste:**
1. **Insert → Module**
2. Đổi tên thành `modMain` trong Properties
3. Mở file `modMain.bas` bằng text editor
4. Copy toàn bộ nội dung (bỏ dòng `Attribute VB_Name`)
5. Paste vào module
6. Lặp lại cho `modBaoCao.bas`

### Bước 4: Thêm Sheet Code

1. Trong Project Explorer, **double-click** `Sheet1 (SO DO KHO)`
2. Mở file `Sheet_SODOKHO.cls` bằng text editor
3. Copy code (bỏ phần comment đầu)
4. Paste vào cửa sổ code

### Bước 5: Thêm ThisWorkbook Code

1. **Double-click** `ThisWorkbook`
2. Mở file `ThisWorkbook.cls`
3. Copy code
4. Paste vào

### Bước 6: Tạo UserForm

1. **Insert → UserForm**
2. Đổi `(Name)` thành `frmNhapXuat`
3. Tạo các control theo bảng trong file `frmNhapXuat.frm`
4. Double-click UserForm để mở code
5. Copy code từ file `frmNhapXuat.frm`
6. Paste vào

### Bước 7: Tạo nút bấm

1. Quay lại Excel
2. Vào sheet **SO DO KHO**
3. **Insert → Shapes → Rectangle**
4. Vẽ 4 nút: Nhập, Xuất, Đóng, Mở
5. Right-click từng nút → **Assign Macro**:
   - Nhập → `NhapHang`
   - Xuất → `XuatHang`
   - Đóng → `DongOKho`
   - Mở → `MoOKho`

### Bước 8: Lưu và Test

1. **Ctrl + S** lưu file
2. Đóng và mở lại
3. Enable Macros
4. Test bằng cách click vào ô kho

---

## Cấu Trúc Project

```
VBAProject (QuanLyKho.xlsm)
├── Microsoft Excel Objects
│   ├── Sheet1 (SO DO KHO)  ← Sheet_SODOKHO.cls
│   ├── Sheet2 (VI TRI)
│   ├── Sheet3 (SAN PHAM)
│   ├── Sheet4 (PHAT SINH)
│   ├── Sheet5 (TON KHO)
│   ├── Sheet6 (BAO CAO)
│   └── ThisWorkbook        ← ThisWorkbook.cls
├── Forms
│   └── frmNhapXuat         ← frmNhapXuat.frm
└── Modules
    ├── modMain             ← modMain.bas
    └── modBaoCao           ← modBaoCao.bas
```

---

## Lưu Ý

- Giá trị trong code dùng **không dấu** để tránh lỗi encoding:
  - `"Mo"` thay vì `"Mở"`
  - `"Dong"` thay vì `"Đóng"`
  - `"Nhap"` / `"Xuat"` thay vì `"Nhập"` / `"Xuất"`
  - `"Dung"` thay vì `"Dùng"`

- Cần cập nhật dữ liệu trong Excel cho khớp:
  - Sheet VI TRI: `TrangThai` = "Mo" hoặc "Dong"
  - Sheet SAN PHAM: `TrangThai` = "Dung" hoặc "Ngung"
