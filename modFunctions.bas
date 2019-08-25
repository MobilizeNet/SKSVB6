Attribute VB_Name = "modFunctions"
Option Explicit

Public Sub AppendAND(ByRef filter As String)
    If filter <> Empty Then
        filter = filter & " AND "
    End If
End Sub

Public Function AddToCollection(col As Collection, Item As String) As Boolean
AddToCollection = False
If Not Exists(col, Item) Then
    col.Add Item, Item
    AddToCollection = True
End If
End Function

Public Function Exists(col As Collection, Index As String) As Boolean
Dim o As Variant
On Error GoTo Error
    o = col(Index)
Error:
   Exists = o <> Empty
End Function


Public Function DoubleValue(strValue As String)
If Len(strValue) <> 0 Then
    DoubleValue = CDbl(strValue)
Else
    DoubleValue = 0
End If
End Function

Public Function ValidateTextBoxDouble(txBox As textbox, parentForm As Form)
On Error GoTo err:
   DoubleValue txBox.text
   ValidateTextBoxDouble = True
   Exit Function
err:
   modMain.LogStatus "The value inserted is not valid", parentForm
   txBox.text = ""
   txBox.SetFocus
   ValidateTextBoxDouble = False
End Function

Public Function ValidateTextDouble(text As String, parentForm As Form)
On Error GoTo err:
   DoubleValue text
   ValidateTextDouble = True
   Exit Function
err:
   modMain.LogStatus "The value inserted is not valid", parentForm
   ValidateTextDouble = False
End Function

Public Sub SelectAll(ByRef txtBox As textbox)
txtBox.SelStart = 0
txtBox.SelLength = Len(txtBox)
End Sub

Public Function UpCase(ByRef KeyAscii As Integer)
UpCase = Asc(UCase(Chr(KeyAscii)))
End Function


''''''''''''''''''''''''''''''''''
''' Combobox related functions '''
''''''''''''''''''''''''''''''''''

Public Sub LoadCombo(Table As String, combo As ComboBox, _
                    field As String, Optional valueField As String)
ExecuteSql "Select * From " & Table
combo.Clear
If (valueField <> Empty) Then
    While Not rs.EOF
        combo.AddItem (rs.Fields(field))
        combo.ItemData(combo.NewIndex) = rs.Fields(valueField)
        rs.MoveNext
    Wend
Else
    While Not rs.EOF
        combo.AddItem (rs.Fields(field))
        rs.MoveNext
    Wend
End If
'If strDefault <> Empty Then
   ' combo = strDefault
'End If
End Sub


Public Function ComboEmpty(ByRef combo As ComboBox, _
                Optional strip As Variant, _
                Optional Index As Integer) _
                As Boolean
If combo.ListIndex = -1 Then
    ComboEmpty = True
    MsgBox "Please select an option from the list", vbExclamation
    If Index <> Empty Then
        'strip.SelectedItem = strip.Tabs(Index)
    End If
    combo.SetFocus
Else
    ComboEmpty = False
End If
End Function

Public Function NoRecords(lstView As ListView, Optional Prompt As String) As Boolean
If lstView.ListItems.Count = 0 Or lstView.SelectedItem Is Nothing Then
    If Prompt <> Empty Then
        MsgBox Prompt, vbExclamation
    End If
    NoRecords = True
Else
    NoRecords = False
End If
End Function

Public Function RcrdId(Table As String, Optional Identifier As String, Optional FldNo As String) As String
Dim RcrdNo As Integer
ExecuteSql "Select * from " & Table & " order by " & FldNo & " ASC"
If rs.EOF = False Then
    rs.MoveLast
    RcrdNo = rs.Fields(FldNo) + 1
Else
    RcrdNo = 1
End If
If Identifier <> Empty Then
    RcrdId = Identifier & RcrdNo & Format(Date, "mm")
Else
    RcrdId = RcrdNo
End If
End Function



'''''''''''''''''''''''''''''''''''''''''
Public Sub SearchShow(Table As String, fieldToSearch As String, itemToSearch As String)
With frmSearch
    .Search Table, fieldToSearch, itemToSearch
    .Show vbModal
End With
End Sub

Public Function ValBox(Prompt As String, Icon As Image, Optional Title As String, _
                        Optional Default As Double, _
                        Optional Header As String = "Value Box") As Double
'With frmValue
'    If Title <> Empty Then
 '       .Caption = Title
'    Else
'        .Caption = App.Title
'    End If
'    .lblHeader.Caption = StrConv(Header, vbUpperCase)
'    .imgIcon.Picture = Icon.Picture
'    .lblPrompt.Caption = Prompt
'    .Default Val(Default)
'    .Show vbModal
'    ValBox = Val(.txtValue.Text)
'    Unload frmValue
'End With
End Function


Public Function TextBoxEmpty(ByRef stext As textbox, Optional TabObject As Variant, Optional TabIndex As Integer) As Boolean
If Trim(stext) = Empty Or stext.text = "  /  /    " Then
    TextBoxEmpty = True
    MsgBox "You need to fill in all required fields", vbExclamation
    If TabIndex <> Empty Then
        'TabObject.SelectedItem = TabObject.Tabs(TabIndex)
    End If
    stext.SetFocus
Else
    TextBoxEmpty = False
End If
End Function

Public Function TextBoxNumberEmpty(ByRef textbox As textbox) As Boolean
'if the input is not a numeric then true
If IsNumeric(textbox.text) = False Then
    TextBoxNumberEmpty = True
    MsgBox "The field requires a numeric value.", vbExclamation
    textbox.SetFocus
    SelectAll textbox
Else
    TextBoxNumberEmpty = False
End If
End Function



Private Sub SaveDetection(Reference As String, Title As String, Description As String, Table As String)
ExecuteSql2 "Select * from " & Table
rs2.AddNew
rs2.Fields!record_no = Val(RcrdId(Table, , "record_no"))
rs2.Fields!Reference = Reference
rs2.Fields!war_type = Title
rs2.Fields!Description = Description
rs2.Update
End Sub



Public Function ExecErr(Prompt As String, _
                        Optional PromptFld As String, _
                        Optional Table As String, _
                        Optional RcrdFld As String, _
                        Optional RcrdStr As String) As String
Dim Rcrds As String
If Table <> Empty Then
    ExecuteSql "Select * from " & Table & " where " & RcrdFld & " = '" & RcrdStr & "'"
    While Not rs.EOF
        Rcrds = Rcrds & rs.Fields(PromptFld) & "; "
        rs.MoveNext
    Wend
    ExecErr = "Error: " & Prompt & vbNewLine & vbNewLine & _
            "Related Records: " & Rcrds
Else
    ExecErr = Prompt
End If
End Function

