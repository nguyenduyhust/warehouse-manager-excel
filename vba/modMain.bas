Attribute VB_Name = "modMain"
Option Explicit

' ===== CONSTANTS - TEN SHEET =====
Public Const SHEET_SODOKHO As String = "SO DO KHO"
Public Const SHEET_VITRI As String = "VI TRI"
Public Const SHEET_SANPHAM As String = "SAN PHAM"
Public Const SHEET_PHATSINH As String = "PHAT SINH"
Public Const SHEET_TONKHO As String = "TON KHO"
Public Const SHEET_BAOCAO As String = "BAO CAO"

' ===== CONSTANTS - TEN COT (Header PascalCase) =====
' Sheet VI TRI: MaViTri | TrangThai | GhiChu
' Sheet SAN PHAM: MaSP | MaGo | DoDay | TrangThai | GhiChu
' Sheet PHAT SINH: Ngay | Gio | Loai | MaViTri | MaSP | SoTam | SoTamQuyDoi | MaGo | DoDay | GhiChu
' Sheet TON KHO: MaViTri | MaSP | MaGo | DoDay | SoTam

' Vung thong tin o kho tren SO DO KHO
Public Const INFO_COL As String = "AD"
Public Const INFO_MAVITRI_ROW As Integer = 2
Public Const INFO_TRANGTHAI_ROW As Integer = 3
Public Const INFO_TONKHO_START_ROW As Integer = 8

' Bien luu vi tri dang chon
Public CurrentMaViTri As String

' ===== MAPPING O KHO =====
' Tra ve dia chi cell cua MaViTri tren sheet SODOKHO
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

' Tra ve MaViTri tu dia chi cell
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

' ===== LAY TRANG THAI O KHO =====
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

    GetTrangThai = "Mo" ' Mac dinh
End Function

' ===== CAP NHAT TRANG THAI =====
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

' ===== LAY DANH SACH TON KHO THEO VI TRI =====
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

' ===== HIEN THI THONG TIN O KHO =====
Public Sub ShowOKhoInfo(ByVal MaViTri As String)
    If MaViTri = "" Then Exit Sub

    CurrentMaViTri = MaViTri

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SODOKHO)

    ' Hien thi MaViTri
    ws.Range(INFO_COL & INFO_MAVITRI_ROW).Value = MaViTri

    ' Hien thi TrangThai
    ws.Range(INFO_COL & INFO_TRANGTHAI_ROW).Value = GetTrangThai(MaViTri)

    ' Xoa danh sach ton cu
    Dim clearRow As Integer
    For clearRow = INFO_TONKHO_START_ROW To INFO_TONKHO_START_ROW + 50
        ws.Range("AC" & clearRow & ":AF" & clearRow).ClearContents
    Next clearRow

    ' Hien thi danh sach ton moi
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

' ===== CAP NHAT MAU SO DO KHO =====
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
            If trangthai = "Dong" Then
                .Interior.Color = RGB(192, 192, 192) ' Xam
                .Font.Color = RGB(128, 128, 128)
            ElseIf hasTon Then
                .Interior.Color = RGB(198, 239, 206) ' Xanh la nhat
                .Font.Color = RGB(0, 97, 0)
            Else
                .Interior.Color = RGB(255, 255, 255) ' Trang
                .Font.Color = RGB(0, 0, 0)
            End If
        End With
    Next i

    Application.ScreenUpdating = True
End Sub

' Kiem tra co ton kho khong
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

' ===== GHI PHAT SINH =====
Public Sub GhiPhatSinh(ByVal Loai As String, ByVal MaViTri As String, _
                       ByVal MaSP As String, ByVal SoTam As Double, _
                       ByVal GhiChu As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_PHATSINH)

    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ' Lay thong tin san pham
    Dim MaGo As String
    Dim DoDay As Double
    GetSanPhamInfo MaSP, MaGo, DoDay

    ' Tinh SoTamQuyDoi
    Dim SoTamQuyDoi As Double
    If Loai = "Nhap" Then
        SoTamQuyDoi = SoTam
    Else
        SoTamQuyDoi = -SoTam
    End If

    ' Ghi du lieu
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

    ' Format ngay gio
    ws.Cells(newRow, 1).NumberFormat = "dd/mm/yyyy"
    ws.Cells(newRow, 2).NumberFormat = "hh:mm:ss"
End Sub

' Lay thong tin san pham
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

' ===== CAP NHAT TON KHO =====
Public Sub UpdateTonKho(ByVal MaViTri As String, ByVal MaSP As String, _
                        ByVal SoTamChange As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TONKHO)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Tim dong hien co
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = MaViTri And ws.Cells(i, 2).Value = MaSP Then
            ws.Cells(i, 5).Value = ws.Cells(i, 5).Value + SoTamChange
            Exit Sub
        End If
    Next i

    ' Khong tim thay - them moi
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

' ===== LAY DANH SACH SAN PHAM DANG DUNG =====
Public Function GetSanPhamList() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SANPHAM)

    Dim result As New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 4).Value = "Dung" Then
            result.Add ws.Cells(i, 1).Value
        End If
    Next i

    Set GetSanPhamList = result
End Function

' ===== LAY DANH SACH SP TON TAI VI TRI =====
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

' Lay so ton
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

' ===== XU LY NUT DONG/MO =====
Public Sub DongOKho()
    If CurrentMaViTri = "" Then
        MsgBox "Vui long chon mot o kho!", vbExclamation
        Exit Sub
    End If

    If GetTrangThai(CurrentMaViTri) = "Dong" Then
        MsgBox "O kho nay da dong!", vbExclamation
        Exit Sub
    End If

    SetTrangThai CurrentMaViTri, "Dong"
    UpdateWarehouseColors
    ShowOKhoInfo CurrentMaViTri
    MsgBox "Da dong o " & CurrentMaViTri, vbInformation
End Sub

Public Sub MoOKho()
    If CurrentMaViTri = "" Then
        MsgBox "Vui long chon mot o kho!", vbExclamation
        Exit Sub
    End If

    If GetTrangThai(CurrentMaViTri) = "Mo" Then
        MsgBox "O kho nay da mo!", vbExclamation
        Exit Sub
    End If

    SetTrangThai CurrentMaViTri, "Mo"
    UpdateWarehouseColors
    ShowOKhoInfo CurrentMaViTri
    MsgBox "Da mo o " & CurrentMaViTri, vbInformation
End Sub

' ===== XU LY NUT NHAP/XUAT =====
Public Sub NhapHang()
    If CurrentMaViTri = "" Then
        MsgBox "Vui long chon mot o kho!", vbExclamation
        Exit Sub
    End If

    If GetTrangThai(CurrentMaViTri) = "Dong" Then
        MsgBox "O kho nay dang dong! Khong the nhap hang.", vbExclamation
        Exit Sub
    End If

    frmNhapXuat.ShowForm "Nhap", CurrentMaViTri
End Sub

Public Sub XuatHang()
    If CurrentMaViTri = "" Then
        MsgBox "Vui long chon mot o kho!", vbExclamation
        Exit Sub
    End If

    If GetTrangThai(CurrentMaViTri) = "Dong" Then
        MsgBox "O kho nay dang dong! Khong the xuat hang.", vbExclamation
        Exit Sub
    End If

    If Not HasTonKho(CurrentMaViTri) Then
        MsgBox "O kho nay khong co hang de xuat!", vbExclamation
        Exit Sub
    End If

    frmNhapXuat.ShowForm "Xuat", CurrentMaViTri
End Sub

' ===== KHOI TAO DU LIEU BAN DAU =====
Public Sub InitializeData()
    ' Khoi tao VITRI
    Dim wsVitri As Worksheet
    Set wsVitri = ThisWorkbook.Sheets(SHEET_VITRI)

    If wsVitri.Cells(2, 1).Value = "" Then
        Dim i As Integer
        For i = 1 To 104
            wsVitri.Cells(i + 1, 1).Value = "K" & i
            wsVitri.Cells(i + 1, 2).Value = "Mo"
        Next i
    End If

    ' Cap nhat mau so do
    UpdateWarehouseColors
End Sub
