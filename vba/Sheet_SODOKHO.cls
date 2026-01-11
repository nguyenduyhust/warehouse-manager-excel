' Sheet Code cho SO DO KHO
' Copy vao: VBA Editor > Double-click "Sheet1 (SO DO KHO)"

Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Chi xu ly khi click vao vung so do kho
    If Target.Cells.Count > 1 Then Exit Sub

    Dim MaViTri As String
    MaViTri = GetMaViTriFromCell(Target.Row, Target.Column)

    If MaViTri <> "" Then
        ShowOKhoInfo MaViTri
    End If
End Sub
