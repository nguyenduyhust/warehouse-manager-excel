# Quản Lý Kho Gỗ - Excel VBA

Hệ thống quản lý kho gỗ sử dụng Excel VBA, theo dõi 104 ô kho chứa các tấm gỗ với nhiều loại sản phẩm khác nhau.

## Tính Năng

- **Sơ đồ kho trực quan**: 104 ô kho (K1-K104) với màu sắc thể hiện trạng thái
- **Quản lý xuất/nhập**: Ghi nhận đầy đủ lịch sử với ngày giờ
- **Tồn kho tự động**: Tính toán tồn theo vị trí và sản phẩm
- **Báo cáo**: Tổng hợp và chi tiết theo khoảng thời gian

## Cấu Trúc File Excel

| Sheet | Mục đích |
|-------|----------|
| SO DO KHO | Giao diện chính - sơ đồ kho + thông tin ô |
| VI TRI | Danh mục 104 vị trí, trạng thái Mở/Đóng |
| SAN PHAM | Danh mục sản phẩm gỗ |
| PHAT SINH | Lịch sử xuất/nhập hàng |
| TON KHO | Tồn kho theo vị trí và sản phẩm |
| BAO CAO | Báo cáo tổng hợp và chi tiết |

## Hướng Dẫn Cài Đặt

1. Mở file `QuanLyKho.xlsm`
2. Enable Macros khi được hỏi
3. Nhấn `Alt + F11` để mở VBA Editor
4. Import code từ `docs/VBA_CODE.md`

Chi tiết xem tại: [Hướng dẫn sử dụng](docs/HUONG_DAN_SU_DUNG.md)

## Tài Liệu

- [Thiết kế chi tiết](docs/THIET_KE.md) - Cấu trúc dữ liệu, logic nghiệp vụ
- [Code VBA](docs/VBA_CODE.md) - Toàn bộ code với hướng dẫn
- [Hướng dẫn sử dụng](docs/HUONG_DAN_SU_DUNG.md) - Dành cho người dùng cuối

## Yêu Cầu

- Microsoft Excel 2016 trở lên (hỗ trợ .xlsm)
- Bật Macros để sử dụng đầy đủ tính năng
