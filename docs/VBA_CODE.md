# Code VBA - Hệ Thống Quản Lý Kho Gỗ

## Hướng Dẫn Cài Đặt

1. Mở file Excel (.xlsm)
2. Nhấn `Alt + F11` để mở VBA Editor
3. Tạo các Module và UserForm theo hướng dẫn bên dưới
4. Copy code vào từng phần tương ứng

---

## 1. Module: modMain (Insert > Module)

```vba
Option Explicit

' ===== CONSTANTS - TÊN SHEET =====
Public Const SHEET_SODOKHO As String = "SO DO KHO"
Public Const SHEET_VITRI As String = "VI TRI"
Public Const SHEET_SANPHAM As String = "SAN PHAM"
Public Const SHEET_PHATSINH As String = "PHAT SINH"
Public Const SHEET_TONKHO As String = "TON KHO"
Public Const SHEET_BAOCAO As String = "BAO CAO"

' ===== CONSTANTS - TÊN CỘT (Header PascalCase) =====
' Sheet VI TRI: MaViTri | TrangThai | GhiChu
' Sheet SAN PHAM: MaSP | MaGo | DoDay | TrangThai | GhiChu
' Sheet PHAT SINH: Ngay | Gio | Loai | MaViTri | MaSP | SoTam | SoTamQuyDoi | MaGo | DoDay | GhiChu
' Sheet TON KHO: MaViTri | MaSP | MaGo | DoDay | SoTam

' Vùng thông tin ô kho trên SO DO KHO
Public Const INFO_COL As String = "AD"
Public Const INFO_MAVITRI_ROW As Integer = 2
Public Const INFO_TRANGTHAI_ROW As Integer = 3
Public Const INFO_TONKHO_START_ROW As Integer = 8

' Biến lưu vị trí đang chọn
Public CurrentMaViTri As String

' ===== MAPPING Ô KHO =====
' Trả về địa chỉ cell của MaViTri trên sheet SODOKHO
Public Function GetCellAddress(ByVal MaViTri As String) As String
    Dim num As Integer
    num = CInt(Mid(MaViTri, 2))

    Dim rowNum As Integer
    Dim colNum As Integer

    Select Case num
        Case 1 To 26
            rowNum = 2
            colNum = num
        Case 27 To 52
            rowNum = 3
            colNum = num - 26
        Case 53 To 78
            rowNum = 5
            colNum = num - 52
        Case 79 To 104
            rowNum = 6
            colNum = num - 78
    End Select

    GetCellAddress = Cells(rowNum, colNum).Address(False, False)
End Function

' Trả về MaViTri từ địa chỉ cell
Public Function GetMaViTriFromCell(ByVal cellRow As Integer, ByVal cellCol As Integer) As String
    Dim num As Integer

    Select Case cellRow
        Case 2
            If cellCol >= 1 And cellCol <= 26 Then num = cellCol
        Case 3
            If cellCol >= 1 And cellCol <= 26 Then num = cellCol + 26
        Case 5
            If cellCol >= 1 And cellCol <= 26 Then num = cellCol + 52
        Case 6
            If cellCol >= 1 And cellCol <= 26 Then num = cellCol + 78
        Case Else
            GetMaViTriFromCell = ""
            Exit Function
    End Select

    If num >= 1 And num <= 104 Then
        GetMaViTriFromCell = "K" & num
    Else
        GetMaViTriFromCell = ""
    End If
End Function

' ===== LẤY TRẠNG THÁI Ô KHO =====
Public Function GetTrangThai(ByVal MaViTri As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_VITRI)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = MaViTri Then
            GetTrangThai = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i

    GetTrangThai = "Mở" ' Mặc định
End Function

' ===== CẬP NHẬT TRẠNG THÁI =====
Public Sub SetTrangThai(ByVal MaViTri As String, ByVal TrangThai As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_VITRI)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = MaViTri Then
            ws.Cells(i, 2).Value = TrangThai
            Exit Sub
        End If
    Next i
End Sub

' ===== LẤY DANH SÁCH TỒN KHO THEO VỊ TRÍ =====
Public Function GetTonKho(ByVal MaViTri As String) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TONKHO)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim result() As Variant
    Dim count As Integer
    count = 0

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = MaViTri And ws.Cells(i, 5).Value > 0 Then
            count = count + 1
            ReDim Preserve result(1 To 4, 1 To count)
            result(1, count) = ws.Cells(i, 2).Value ' MaSP
            result(2, count) = ws.Cells(i, 3).Value ' MaGo
            result(3, count) = ws.Cells(i, 4).Value ' DoDay
            result(4, count) = ws.Cells(i, 5).Value ' SoTam
        End If
    Next i

    If count = 0 Then
        GetTonKho = Empty
    Else
        GetTonKho = result
    End If
End Function

' ===== HIỂN THỊ THÔNG TIN Ô KHO =====
Public Sub ShowOKhoInfo(ByVal MaViTri As String)
    If MaViTri = "" Then Exit Sub

    CurrentMaViTri = MaViTri

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SODOKHO)

    ' Hiển thị MaViTri
    ws.Range(INFO_COL & INFO_MAVITRI_ROW).Value = MaViTri

    ' Hiển thị TrangThai
    ws.Range(INFO_COL & INFO_TRANGTHAI_ROW).Value = GetTrangThai(MaViTri)

    ' Xóa danh sách tồn cũ
    Dim clearRow As Integer
    For clearRow = INFO_TONKHO_START_ROW To INFO_TONKHO_START_ROW + 50
        ws.Range("AC" & clearRow & ":AF" & clearRow).ClearContents
    Next clearRow

    ' Hiển thị danh sách tồn mới
    Dim tonkho As Variant
    tonkho = GetTonKho(MaViTri)

    If Not IsEmpty(tonkho) Then
        Dim j As Integer
        For j = 1 To UBound(tonkho, 2)
            ws.Cells(INFO_TONKHO_START_ROW + j - 1, 29).Value = tonkho(1, j) ' MaSP - cot AC
            ws.Cells(INFO_TONKHO_START_ROW + j - 1, 30).Value = tonkho(2, j) ' MaGo - cot AD
            ws.Cells(INFO_TONKHO_START_ROW + j - 1, 31).Value = tonkho(3, j) ' DoDay - cot AE
            ws.Cells(INFO_TONKHO_START_ROW + j - 1, 32).Value = tonkho(4, j) ' SoTam - cot AF
        Next j
    End If
End Sub

' ===== CẬP NHẬT MÀU SƠ ĐỒ KHO =====
Public Sub UpdateWarehouseColors()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SODOKHO)

    Dim i As Integer
    Dim MaViTri As String
    Dim cellAddr As String
    Dim trangthai As String
    Dim hasTon As Boolean

    Application.ScreenUpdating = False

    For i = 1 To 104
        MaViTri = "K" & i
        cellAddr = GetCellAddress(MaViTri)
        trangthai = GetTrangThai(MaViTri)
        hasTon = HasTonKho(MaViTri)

        With ws.Range(cellAddr)
            If trangthai = "Đóng" Then
                .Interior.Color = RGB(192, 192, 192) ' Xám
                .Font.Color = RGB(128, 128, 128)
            ElseIf hasTon Then
                .Interior.Color = RGB(198, 239, 206) ' Xanh lá nhạt
                .Font.Color = RGB(0, 97, 0)
            Else
                .Interior.Color = RGB(255, 255, 255) ' Trắng
                .Font.Color = RGB(0, 0, 0)
            End If
        End With
    Next i

    Application.ScreenUpdating = True
End Sub

' Kiểm tra có tồn kho không
Public Function HasTonKho(ByVal MaViTri As String) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TONKHO)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = MaViTri And ws.Cells(i, 5).Value > 0 Then
            HasTonKho = True
            Exit Function
        End If
    Next i

    HasTonKho = False
End Function

' ===== GHI PHÁT SINH =====
Public Sub GhiPhatSinh(ByVal Loai As String, ByVal MaViTri As String, _
                       ByVal MaSP As String, ByVal SoTam As Double, _
                       ByVal GhiChu As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_PHATSINH)

    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ' Lấy thông tin sản phẩm
    Dim MaGo As String
    Dim DoDay As Double
    GetSanPhamInfo MaSP, MaGo, DoDay

    ' Tính SoTamQuyDoi
    Dim SoTamQuyDoi As Double
    If Loai = "Nhập" Then
        SoTamQuyDoi = SoTam
    Else
        SoTamQuyDoi = -SoTam
    End If

    ' Ghi dữ liệu
    ws.Cells(newRow, 1).Value = Date          ' Ngay
    ws.Cells(newRow, 2).Value = Time          ' Gio
    ws.Cells(newRow, 3).Value = Loai          ' Loai
    ws.Cells(newRow, 4).Value = MaViTri       ' MaViTri
    ws.Cells(newRow, 5).Value = MaSP          ' MaSP
    ws.Cells(newRow, 6).Value = SoTam         ' SoTam
    ws.Cells(newRow, 7).Value = SoTamQuyDoi   ' SoTamQuyDoi
    ws.Cells(newRow, 8).Value = MaGo          ' MaGo
    ws.Cells(newRow, 9).Value = DoDay         ' DoDay
    ws.Cells(newRow, 10).Value = GhiChu       ' GhiChu

    ' Format ngày giờ
    ws.Cells(newRow, 1).NumberFormat = "dd/mm/yyyy"
    ws.Cells(newRow, 2).NumberFormat = "hh:mm:ss"
End Sub

' Lấy thông tin sản phẩm
Public Sub GetSanPhamInfo(ByVal MaSP As String, ByRef MaGo As String, ByRef DoDay As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SANPHAM)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = MaSP Then
            MaGo = ws.Cells(i, 2).Value
            DoDay = ws.Cells(i, 3).Value
            Exit Sub
        End If
    Next i

    MaGo = ""
    DoDay = 0
End Sub

' ===== CẬP NHẬT TỒN KHO =====
Public Sub UpdateTonKho(ByVal MaViTri As String, ByVal MaSP As String, _
                        ByVal SoTamChange As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TONKHO)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Tìm dòng hiện có
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = MaViTri And ws.Cells(i, 2).Value = MaSP Then
            ws.Cells(i, 5).Value = ws.Cells(i, 5).Value + SoTamChange
            Exit Sub
        End If
    Next i

    ' Không tìm thấy - thêm mới
    Dim newRow As Long
    newRow = lastRow + 1

    Dim MaGo As String
    Dim DoDay As Double
    GetSanPhamInfo MaSP, MaGo, DoDay

    ws.Cells(newRow, 1).Value = MaViTri
    ws.Cells(newRow, 2).Value = MaSP
    ws.Cells(newRow, 3).Value = MaGo
    ws.Cells(newRow, 4).Value = DoDay
    ws.Cells(newRow, 5).Value = SoTamChange
End Sub

' ===== LẤY DANH SÁCH SẢN PHẨM ĐANG DÙNG =====
Public Function GetSanPhamList() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SANPHAM)

    Dim result As New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 4).Value = "Dùng" Then
            result.Add ws.Cells(i, 1).Value
        End If
    Next i

    Set GetSanPhamList = result
End Function

' ===== LẤY DANH SÁCH SP TỒN TẠI VỊ TRÍ =====
Public Function GetSanPhamTonTaiViTri(ByVal MaViTri As String) As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TONKHO)

    Dim result As New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = MaViTri And ws.Cells(i, 5).Value > 0 Then
            result.Add ws.Cells(i, 2).Value
        End If
    Next i

    Set GetSanPhamTonTaiViTri = result
End Function

' Lấy số tồn
Public Function GetSoTamTon(ByVal MaViTri As String, ByVal MaSP As String) As Double
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TONKHO)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = MaViTri And ws.Cells(i, 2).Value = MaSP Then
            GetSoTamTon = ws.Cells(i, 5).Value
            Exit Function
        End If
    Next i

    GetSoTamTon = 0
End Function

' ===== XỬ LÝ NÚT ĐÓNG/MỞ =====
Public Sub DongOKho()
    If CurrentMaViTri = "" Then
        MsgBox "Vui lòng chọn một ô kho!", vbExclamation
        Exit Sub
    End If

    If GetTrangThai(CurrentMaViTri) = "Đóng" Then
        MsgBox "Ô kho này đã đóng!", vbExclamation
        Exit Sub
    End If

    SetTrangThai CurrentMaViTri, "Đóng"
    UpdateWarehouseColors
    ShowOKhoInfo CurrentMaViTri
    MsgBox "Đã đóng ô " & CurrentMaViTri, vbInformation
End Sub

Public Sub MoOKho()
    If CurrentMaViTri = "" Then
        MsgBox "Vui lòng chọn một ô kho!", vbExclamation
        Exit Sub
    End If

    If GetTrangThai(CurrentMaViTri) = "Mở" Then
        MsgBox "Ô kho này đã mở!", vbExclamation
        Exit Sub
    End If

    SetTrangThai CurrentMaViTri, "Mở"
    UpdateWarehouseColors
    ShowOKhoInfo CurrentMaViTri
    MsgBox "Đã mở ô " & CurrentMaViTri, vbInformation
End Sub

' ===== XỬ LÝ NÚT NHẬP/XUẤT =====
Public Sub NhapHang()
    If CurrentMaViTri = "" Then
        MsgBox "Vui lòng chọn một ô kho!", vbExclamation
        Exit Sub
    End If

    If GetTrangThai(CurrentMaViTri) = "Đóng" Then
        MsgBox "Ô kho này đang đóng! Không thể nhập hàng.", vbExclamation
        Exit Sub
    End If

    frmNhapXuat.ShowForm "Nhập", CurrentMaViTri
End Sub

Public Sub XuatHang()
    If CurrentMaViTri = "" Then
        MsgBox "Vui lòng chọn một ô kho!", vbExclamation
        Exit Sub
    End If

    If GetTrangThai(CurrentMaViTri) = "Đóng" Then
        MsgBox "Ô kho này đang đóng! Không thể xuất hàng.", vbExclamation
        Exit Sub
    End If

    If Not HasTonKho(CurrentMaViTri) Then
        MsgBox "Ô kho này không có hàng để xuất!", vbExclamation
        Exit Sub
    End If

    frmNhapXuat.ShowForm "Xuất", CurrentMaViTri
End Sub

' ===== KHỞI TẠO DỮ LIỆU BAN ĐẦU =====
Public Sub InitializeData()
    ' Khởi tạo VITRI
    Dim wsVitri As Worksheet
    Set wsVitri = ThisWorkbook.Sheets(SHEET_VITRI)

    If wsVitri.Cells(2, 1).Value = "" Then
        Dim i As Integer
        For i = 1 To 104
            wsVitri.Cells(i + 1, 1).Value = "K" & i
            wsVitri.Cells(i + 1, 2).Value = "Mở"
        Next i
    End If

    ' Cập nhật màu sơ đồ
    UpdateWarehouseColors
End Sub
```

---

## 2. Sheet Code: SODOKHO (Double-click sheet > View Code)

```vba
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Chỉ xử lý khi click vào vùng sơ đồ kho
    If Target.Cells.Count > 1 Then Exit Sub

    Dim MaViTri As String
    MaViTri = GetMaViTriFromCell(Target.Row, Target.Column)

    If MaViTri <> "" Then
        ShowOKhoInfo MaViTri
    End If
End Sub
```

---

## 3. UserForm: frmNhapXuat

### Tạo UserForm:
1. Insert > UserForm
2. Đặt tên: `frmNhapXuat`
3. Thêm các controls sau:

| Control | Name | Caption/Text |
|---------|------|--------------|
| Label | lblTitle | Nhập Hàng |
| Label | lblViTri | Vị trí: |
| Label | lblViTriValue | K1 |
| Label | lblSanPham | Sản phẩm: |
| ComboBox | cboSanPham | |
| Label | lblSoTam | Số tấm: |
| TextBox | txtSoTam | |
| Label | lblTon | Tồn hiện tại: |
| Label | lblTonValue | 0 |
| Label | lblGhiChu | Ghi chú: |
| TextBox | txtGhiChu | |
| CommandButton | btnOK | Xác nhận |
| CommandButton | btnCancel | Hủy |

### Code cho frmNhapXuat:

```vba
Option Explicit

Private mLoai As String
Private mMaViTri As String

Public Sub ShowForm(ByVal Loai As String, ByVal MaViTri As String)
    mLoai = Loai
    mMaViTri = MaViTri

    ' Cập nhật title
    If Loai = "Nhập" Then
        Me.Caption = "Nhập Hàng"
        lblTitle.Caption = "NHẬP HÀNG VÀO KHO"
        lblTon.Visible = False
        lblTonValue.Visible = False
    Else
        Me.Caption = "Xuất Hàng"
        lblTitle.Caption = "XUẤT HÀNG KHỎI KHO"
        lblTon.Visible = True
        lblTonValue.Visible = True
    End If

    ' Hiển thị vị trí
    lblViTriValue.Caption = MaViTri

    ' Load danh sách sản phẩm
    LoadSanPham

    ' Reset form
    txtSoTam.Value = ""
    txtGhiChu.Value = ""
    lblTonValue.Caption = "0"

    Me.Show
End Sub

Private Sub LoadSanPham()
    cboSanPham.Clear

    Dim spList As Collection

    If mLoai = "Nhập" Then
        Set spList = GetSanPhamList()
    Else
        Set spList = GetSanPhamTonTaiViTri(mMaViTri)
    End If

    Dim sp As Variant
    For Each sp In spList
        cboSanPham.AddItem sp
    Next sp

    If cboSanPham.ListCount > 0 Then
        cboSanPham.ListIndex = 0
    End If
End Sub

Private Sub cboSanPham_Change()
    If mLoai = "Xuất" And cboSanPham.Value <> "" Then
        lblTonValue.Caption = GetSoTamTon(mMaViTri, cboSanPham.Value)
    End If
End Sub

Private Sub btnOK_Click()
    ' Validate
    If cboSanPham.Value = "" Then
        MsgBox "Vui lòng chọn sản phẩm!", vbExclamation
        Exit Sub
    End If

    If txtSoTam.Value = "" Or Not IsNumeric(txtSoTam.Value) Then
        MsgBox "Vui lòng nhập số tấm hợp lệ!", vbExclamation
        Exit Sub
    End If

    Dim SoTam As Double
    SoTam = CDbl(txtSoTam.Value)

    If SoTam <= 0 Then
        MsgBox "Số tấm phải lớn hơn 0!", vbExclamation
        Exit Sub
    End If

    ' Kiểm tra tồn kho khi xuất
    If mLoai = "Xuất" Then
        Dim tonHienTai As Double
        tonHienTai = GetSoTamTon(mMaViTri, cboSanPham.Value)
        If SoTam > tonHienTai Then
            MsgBox "Số xuất (" & SoTam & ") không được lớn hơn số tồn (" & tonHienTai & ")!", vbExclamation
            Exit Sub
        End If
    End If

    ' Ghi phát sinh
    GhiPhatSinh mLoai, mMaViTri, cboSanPham.Value, SoTam, txtGhiChu.Value

    ' Cập nhật tồn kho
    If mLoai = "Nhập" Then
        UpdateTonKho mMaViTri, cboSanPham.Value, SoTam
    Else
        UpdateTonKho mMaViTri, cboSanPham.Value, -SoTam
    End If

    ' Refresh giao diện
    UpdateWarehouseColors
    ShowOKhoInfo mMaViTri

    MsgBox mLoai & " thành công!", vbInformation

    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Unload Me
    End If
End Sub
```

---

## 4. Module: modBaoCao

```vba
Option Explicit

' Tạo báo cáo tổng hợp
Public Sub TaoBaoCaoTongHop(ByVal TuNgay As Date, ByVal DenNgay As Date)
    Dim wsPhatSinh As Worksheet
    Dim wsBaoCao As Worksheet
    Set wsPhatSinh = ThisWorkbook.Sheets(SHEET_PHATSINH)
    Set wsBaoCao = ThisWorkbook.Sheets(SHEET_BAOCAO)

    ' Xóa dữ liệu cũ (giữ header)
    wsBaoCao.Range("A10:Z1000").ClearContents

    ' Tiêu đề báo cáo
    wsBaoCao.Range("A1").Value = "BÁO CÁO TỔNG HỢP XUẤT NHẬP KHO"
    wsBaoCao.Range("A2").Value = "Từ ngày: " & Format(TuNgay, "dd/mm/yyyy") & " - Đến ngày: " & Format(DenNgay, "dd/mm/yyyy")

    ' Header
    wsBaoCao.Range("A5").Value = "MaSP"
    wsBaoCao.Range("B5").Value = "MaGo"
    wsBaoCao.Range("C5").Value = "DoDay"
    wsBaoCao.Range("D5").Value = "Tổng Nhập"
    wsBaoCao.Range("E5").Value = "Tổng Xuất"
    wsBaoCao.Range("F5").Value = "Chênh lệch"

    ' Tính toán
    Dim lastRowPS As Long
    lastRowPS = wsPhatSinh.Cells(wsPhatSinh.Rows.Count, "A").End(xlUp).Row

    Dim dictSP As Object
    Set dictSP = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim ngay As Date
    Dim MaSP As String
    Dim Loai As String
    Dim SoTam As Double

    For i = 2 To lastRowPS
        ngay = wsPhatSinh.Cells(i, 1).Value
        If ngay >= TuNgay And ngay <= DenNgay Then
            MaSP = wsPhatSinh.Cells(i, 5).Value
            Loai = wsPhatSinh.Cells(i, 3).Value
            SoTam = wsPhatSinh.Cells(i, 6).Value

            If Not dictSP.Exists(MaSP) Then
                dictSP.Add MaSP, Array(0, 0, wsPhatSinh.Cells(i, 8).Value, wsPhatSinh.Cells(i, 9).Value)
            End If

            Dim arr As Variant
            arr = dictSP(MaSP)

            If Loai = "Nhập" Then
                arr(0) = arr(0) + SoTam
            Else
                arr(1) = arr(1) + SoTam
            End If

            dictSP(MaSP) = arr
        End If
    Next i

    ' Xuất báo cáo
    Dim rowBC As Long
    rowBC = 6

    Dim key As Variant
    For Each key In dictSP.Keys
        arr = dictSP(key)
        wsBaoCao.Cells(rowBC, 1).Value = key
        wsBaoCao.Cells(rowBC, 2).Value = arr(2)
        wsBaoCao.Cells(rowBC, 3).Value = arr(3)
        wsBaoCao.Cells(rowBC, 4).Value = arr(0)
        wsBaoCao.Cells(rowBC, 5).Value = arr(1)
        wsBaoCao.Cells(rowBC, 6).Value = arr(0) - arr(1)
        rowBC = rowBC + 1
    Next key

    MsgBox "Đã tạo báo cáo thành công!", vbInformation
End Sub

' Tạo báo cáo chi tiết theo sản phẩm
Public Sub TaoBaoCaoChiTietSP(ByVal TuNgay As Date, ByVal DenNgay As Date)
    Dim wsPhatSinh As Worksheet
    Dim wsBaoCao As Worksheet
    Set wsPhatSinh = ThisWorkbook.Sheets(SHEET_PHATSINH)
    Set wsBaoCao = ThisWorkbook.Sheets(SHEET_BAOCAO)

    ' Xóa và tạo header
    wsBaoCao.Range("A10:Z1000").ClearContents

    wsBaoCao.Range("A10").Value = "CHI TIẾT THEO SẢN PHẨM"

    ' Header chi tiết
    wsBaoCao.Range("A12").Value = "Ngày"
    wsBaoCao.Range("B12").Value = "Giờ"
    wsBaoCao.Range("C12").Value = "Loại"
    wsBaoCao.Range("D12").Value = "Vị trí"
    wsBaoCao.Range("E12").Value = "MaSP"
    wsBaoCao.Range("F12").Value = "MaGo"
    wsBaoCao.Range("G12").Value = "DoDay"
    wsBaoCao.Range("H12").Value = "SoTam"
    wsBaoCao.Range("I12").Value = "GhiChu"

    ' Copy dữ liệu
    Dim lastRowPS As Long
    lastRowPS = wsPhatSinh.Cells(wsPhatSinh.Rows.Count, "A").End(xlUp).Row

    Dim rowBC As Long
    rowBC = 13

    Dim i As Long
    Dim ngay As Date

    For i = 2 To lastRowPS
        ngay = wsPhatSinh.Cells(i, 1).Value
        If ngay >= TuNgay And ngay <= DenNgay Then
            wsBaoCao.Cells(rowBC, 1).Value = wsPhatSinh.Cells(i, 1).Value
            wsBaoCao.Cells(rowBC, 2).Value = wsPhatSinh.Cells(i, 2).Value
            wsBaoCao.Cells(rowBC, 3).Value = wsPhatSinh.Cells(i, 3).Value
            wsBaoCao.Cells(rowBC, 4).Value = wsPhatSinh.Cells(i, 4).Value
            wsBaoCao.Cells(rowBC, 5).Value = wsPhatSinh.Cells(i, 5).Value
            wsBaoCao.Cells(rowBC, 6).Value = wsPhatSinh.Cells(i, 8).Value
            wsBaoCao.Cells(rowBC, 7).Value = wsPhatSinh.Cells(i, 9).Value
            wsBaoCao.Cells(rowBC, 8).Value = wsPhatSinh.Cells(i, 6).Value
            wsBaoCao.Cells(rowBC, 9).Value = wsPhatSinh.Cells(i, 10).Value

            wsBaoCao.Cells(rowBC, 1).NumberFormat = "dd/mm/yyyy"
            wsBaoCao.Cells(rowBC, 2).NumberFormat = "hh:mm:ss"

            rowBC = rowBC + 1
        End If
    Next i

    MsgBox "Đã tạo báo cáo chi tiết!", vbInformation
End Sub
```

---

## 5. ThisWorkbook Code

```vba
Option Explicit

Private Sub Workbook_Open()
    ' Khởi tạo dữ liệu khi mở file
    InitializeData
End Sub
```

---

## 6. Tạo Nút Bấm trên Sheet SODOKHO

### Cách 1: Sử dụng Shape với Macro
1. Insert > Shapes > Rectangle
2. Vẽ 4 nút: Nhập, Xuất, Đóng, Mở
3. Right-click > Assign Macro:
   - Nút Nhập: `NhapHang`
   - Nút Xuất: `XuatHang`
   - Nút Đóng: `DongOKho`
   - Nút Mở: `MoOKho`

### Cách 2: Sử dụng Form Controls (Developer > Insert > Button)

---

## 7. Thiết Lập Sheet BAOCAO

Thêm các control để chọn khoảng thời gian:

| Vị trí | Control | Mô tả |
|--------|---------|-------|
| A3 | Label | "Từ ngày:" |
| B3 | TextBox/Cell | Nhập ngày bắt đầu |
| C3 | Label | "Đến ngày:" |
| D3 | TextBox/Cell | Nhập ngày kết thúc |
| E3 | Button | "Báo cáo tổng hợp" - gọi TaoBaoCaoTongHop |
| F3 | Button | "Báo cáo chi tiết" - gọi TaoBaoCaoChiTietSP |

---

## Ghi Chú Quan Trọng

1. **Lưu file dạng .xlsm** để giữ VBA code
2. **Enable Macros** khi mở file
3. **Backup dữ liệu** trước khi test
4. Chạy `InitializeData` lần đầu để khởi tạo danh sách vị trí
