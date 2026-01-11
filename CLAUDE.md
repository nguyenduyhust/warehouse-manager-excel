# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Tổng Quan Dự Án

Hệ thống quản lý kho gỗ sử dụng Excel VBA. Quản lý 104 ô kho (K1-K104) chứa các tấm gỗ với các loại sản phẩm khác nhau.

## Cấu Trúc Thư Mục

```
warehouse-manager-excel/
├── QuanLyKho.xlsm          # File Excel chính (có VBA)
├── README.md               # Giới thiệu dự án
├── CLAUDE.md               # File này
├── docs/
│   ├── THIET_KE.md         # Thiết kế chi tiết
│   └── HUONG_DAN_SU_DUNG.md # Hướng dẫn người dùng
└── vba/                    # Code VBA tách riêng từng module
    ├── README.md           # Hướng dẫn import
    ├── modMain.bas         # Module chính
    ├── modBaoCao.bas       # Module báo cáo
    ├── Sheet_SODOKHO.cls   # Code cho sheet SO DO KHO
    ├── ThisWorkbook.cls    # Code cho ThisWorkbook
    └── frmNhapXuat.frm     # UserForm nhập/xuất
```

## Cấu Trúc Excel (6 Sheets)

| Sheet | Mục đích |
|-------|----------|
| SO DO KHO | Giao diện chính - sơ đồ kho 104 ô + vùng thông tin |
| VI TRI | Danh mục vị trí K1-K104, trạng thái Mở/Đóng |
| SAN PHAM | Danh mục sản phẩm gỗ (MaSP, MaGo, DoDay) |
| PHAT SINH | Lịch sử xuất/nhập hàng |
| TON KHO | Tồn kho theo vị trí và sản phẩm |
| BAO CAO | Báo cáo tổng hợp và chi tiết |

## Logic Nghiệp Vụ

- **SoTamQuyDoi**: Nhập = +SoTam, Xuất = -SoTam (dùng SUM để tính tồn)
- **Một ô kho**: Có thể chứa nhiều loại sản phẩm
- **Trạng thái ô**: Mở (có thể xuất/nhập) hoặc Đóng (không thể thao tác)

## Quy Ước Đặt Tên

- **Tên Sheet**: VIẾT HOA, có cách (SO DO KHO, VI TRI)
- **Tên Cột (Header)**: PascalCase, không dấu (MaViTri, TrangThai, SoTam)
- **Giá trị dữ liệu**: Không dấu tiếng Việt để tương thích VBA
  - Trạng thái vị trí: `"Mo"`, `"Dong"`
  - Loại phát sinh: `"Nhap"`, `"Xuat"`
  - Trạng thái sản phẩm: `"Dung"`, `"Ngung"`

## VBA Modules

- `modMain` - Hàm chính: xử lý click, CRUD, cập nhật màu
- `modBaoCao` - Tạo báo cáo tổng hợp và chi tiết
- `frmNhapXuat` - UserForm cho nhập/xuất hàng
- Sheet SO DO KHO - Event Worksheet_SelectionChange

## Làm Việc Với File Excel

Sử dụng Python + openpyxl để đọc/chỉnh sửa file Excel:
```python
import openpyxl
wb = openpyxl.load_workbook('QuanLyKho.xlsx')
```

Lưu ý: Không thể thêm VBA code bằng openpyxl, cần copy thủ công từ `docs/VBA_CODE.md`.
