' ThisWorkbook module - Add this to your Excel workbook's VBA project
' In the VBA Editor, find your workbook under "VBAProject" and double-click "ThisWorkbook"

Private Sub Workbook_Open()
    On Error Resume Next
    StartSystem
    On Error GoTo 0
End Sub
