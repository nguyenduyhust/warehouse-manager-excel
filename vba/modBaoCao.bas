Attribute VB_Name = "modBaoCao"
Option Explicit

' Tao bao cao tong hop
Public Sub TaoBaoCaoTongHop(ByVal TuNgay As Date, ByVal DenNgay As Date)
    Dim wsPhatSinh As Worksheet
    Dim wsBaoCao As Worksheet
    Set wsPhatSinh = ThisWorkbook.Sheets(SHEET_PHATSINH)
    Set wsBaoCao = ThisWorkbook.Sheets(SHEET_BAOCAO)

    ' Xoa du lieu cu (giu header)
    wsBaoCao.Range("A10:Z1000").ClearContents

    ' Tieu de bao cao
    wsBaoCao.Range("A1").Value = "BAO CAO TONG HOP XUAT NHAP KHO"
    wsBaoCao.Range("A2").Value = "Tu ngay: " & Format(TuNgay, "dd/mm/yyyy") & " - Den ngay: " & Format(DenNgay, "dd/mm/yyyy")

    ' Header
    wsBaoCao.Range("A5").Value = "MaSP"
    wsBaoCao.Range("B5").Value = "MaGo"
    wsBaoCao.Range("C5").Value = "DoDay"
    wsBaoCao.Range("D5").Value = "Tong Nhap"
    wsBaoCao.Range("E5").Value = "Tong Xuat"
    wsBaoCao.Range("F5").Value = "Chenh lech"

    ' Tinh toan
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

            If Loai = "Nhap" Then
                arr(0) = arr(0) + SoTam
            Else
                arr(1) = arr(1) + SoTam
            End If

            dictSP(MaSP) = arr
        End If
    Next i

    ' Xuat bao cao
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

    MsgBox "Da tao bao cao thanh cong!", vbInformation
End Sub

' Tao bao cao chi tiet theo san pham
Public Sub TaoBaoCaoChiTietSP(ByVal TuNgay As Date, ByVal DenNgay As Date)
    Dim wsPhatSinh As Worksheet
    Dim wsBaoCao As Worksheet
    Set wsPhatSinh = ThisWorkbook.Sheets(SHEET_PHATSINH)
    Set wsBaoCao = ThisWorkbook.Sheets(SHEET_BAOCAO)

    ' Xoa va tao header
    wsBaoCao.Range("A10:Z1000").ClearContents

    wsBaoCao.Range("A10").Value = "CHI TIET THEO SAN PHAM"

    ' Header chi tiet
    wsBaoCao.Range("A12").Value = "Ngay"
    wsBaoCao.Range("B12").Value = "Gio"
    wsBaoCao.Range("C12").Value = "Loai"
    wsBaoCao.Range("D12").Value = "Vi tri"
    wsBaoCao.Range("E12").Value = "MaSP"
    wsBaoCao.Range("F12").Value = "MaGo"
    wsBaoCao.Range("G12").Value = "DoDay"
    wsBaoCao.Range("H12").Value = "SoTam"
    wsBaoCao.Range("I12").Value = "GhiChu"

    ' Copy du lieu
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

    MsgBox "Da tao bao cao chi tiet!", vbInformation
End Sub
