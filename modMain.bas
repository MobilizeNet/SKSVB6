Attribute VB_Name = "modMain"

Public CurrentUserAdmin As Boolean
Public UserFullname As String
Public UserLevel As String
Public UserId As String

Public ConnectionString As String

Public DetectionType As Integer
Global n As Double, i As Long, s As String, d As Date
Public msg As String
Public ImgName As String, ImgSrc As String





Public Sub Main()
    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Orders.mdb;Persist Security Info=False"

    OpenConnection
    CurrentUserAdmin = True
    UserFullname = "Allan Cantillo"
    UserLevel = "Administrator"
    UserId = "acantillo"
    frmMain.Show
End Sub

Public Sub LogStatus(message As String, Optional frm As Form)
    Dim sb As StatusBar
    Set sb = Nothing
    frmMain.sbStatusBar.Panels(1).text = message
    If Not frm Is Nothing Then
        If frm Is frmAdjustStockManual Then
            Set sb = frmAdjustStockManual.sbStatusBar
        ElseIf frm Is frmActionOrderReception Then
            Set sb = frmActionOrderReception.sbStatusBar
        ElseIf frm Is frmActionOrderRequest Then
            Set sb = frmActionOrderRequest.sbStatusBar
        ElseIf frm Is frmAddStockManual Then
            Set sb = frmAddStockManual.sbStatusBar
        ElseIf frm Is frmReceptionApproval Then
            Set sb = frmReceptionApproval.sbStatusBar
        ElseIf frm Is frmOrderReception Then
            Set sb = frmOrderReception.sbStatusBar
        ElseIf frm Is frmOrderRequest Then
            Set sb = frmOrderRequest.sbStatusBar
        ElseIf frm Is frmRequestApproval Then
            Set sb = frmRequestApproval.sbStatusBar
        End If
        If Not sb Is Nothing Then
            If Not sb.Panels(1) Is Nothing Then
                sb.Panels(1).text = message
            End If
        End If
    End If
End Sub

Public Sub ClearLogStatus(Optional frm As Form)
    LogStatus "", frm
End Sub



