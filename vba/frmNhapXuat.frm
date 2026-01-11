' UserForm: frmNhapXuat
' Tao UserForm trong VBA Editor: Insert > UserForm
' Dat ten (Name): frmNhapXuat
'
' === CAC CONTROL CAN TAO ===
' | Control       | Name           | Caption/Text        |
' |---------------|----------------|---------------------|
' | Label         | lblTitle       | NHAP HANG VAO KHO   |
' | Label         | lblViTri       | Vi tri:             |
' | Label         | lblViTriValue  | K1                  |
' | Label         | lblSanPham     | San pham:           |
' | ComboBox      | cboSanPham     |                     |
' | Label         | lblSoTam       | So tam:             |
' | TextBox       | txtSoTam       |                     |
' | Label         | lblTon         | Ton hien tai:       |
' | Label         | lblTonValue    | 0                   |
' | Label         | lblGhiChu      | Ghi chu:            |
' | TextBox       | txtGhiChu      |                     |
' | CommandButton | btnOK          | Xac nhan            |
' | CommandButton | btnCancel      | Huy                 |
'
' === CODE CHO USERFORM ===

Option Explicit

Private mLoai As String
Private mMaViTri As String

Public Sub ShowForm(ByVal Loai As String, ByVal MaViTri As String)
    mLoai = Loai
    mMaViTri = MaViTri

    ' Cap nhat title
    If Loai = "Nhap" Then
        Me.Caption = "Nhap Hang"
        lblTitle.Caption = "NHAP HANG VAO KHO"
        lblTon.Visible = False
        lblTonValue.Visible = False
    Else
        Me.Caption = "Xuat Hang"
        lblTitle.Caption = "XUAT HANG KHOI KHO"
        lblTon.Visible = True
        lblTonValue.Visible = True
    End If

    ' Hien thi vi tri
    lblViTriValue.Caption = MaViTri

    ' Load danh sach san pham
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

    If mLoai = "Nhap" Then
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
    If mLoai = "Xuat" And cboSanPham.Value <> "" Then
        lblTonValue.Caption = GetSoTamTon(mMaViTri, cboSanPham.Value)
    End If
End Sub

Private Sub btnOK_Click()
    ' Validate
    If cboSanPham.Value = "" Then
        MsgBox "Vui long chon san pham!", vbExclamation
        Exit Sub
    End If

    If txtSoTam.Value = "" Or Not IsNumeric(txtSoTam.Value) Then
        MsgBox "Vui long nhap so tam hop le!", vbExclamation
        Exit Sub
    End If

    Dim SoTam As Double
    SoTam = CDbl(txtSoTam.Value)

    If SoTam <= 0 Then
        MsgBox "So tam phai lon hon 0!", vbExclamation
        Exit Sub
    End If

    ' Kiem tra ton kho khi xuat
    If mLoai = "Xuat" Then
        Dim tonHienTai As Double
        tonHienTai = GetSoTamTon(mMaViTri, cboSanPham.Value)
        If SoTam > tonHienTai Then
            MsgBox "So xuat (" & SoTam & ") khong duoc lon hon so ton (" & tonHienTai & ")!", vbExclamation
            Exit Sub
        End If
    End If

    ' Ghi phat sinh
    GhiPhatSinh mLoai, mMaViTri, cboSanPham.Value, SoTam, txtGhiChu.Value

    ' Cap nhat ton kho
    If mLoai = "Nhap" Then
        UpdateTonKho mMaViTri, cboSanPham.Value, SoTam
    Else
        UpdateTonKho mMaViTri, cboSanPham.Value, -SoTam
    End If

    ' Refresh giao dien
    UpdateWarehouseColors
    ShowOKhoInfo mMaViTri

    MsgBox mLoai & " thanh cong!", vbInformation

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
