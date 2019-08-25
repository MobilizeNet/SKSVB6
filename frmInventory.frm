VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OfrmInventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   Icon            =   "frmInventory.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   7620
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   1164
      ButtonWidth     =   1111
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":06DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":0A2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":0D80
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2355
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox cbYear 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmInventory.frx":10D2
      Left            =   720
      List            =   "frmInventory.frx":10D4
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cbMonth 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmInventory.frx":10D6
      Left            =   3240
      List            =   "frmInventory.frx":1101
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwInventory 
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2355
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Month:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   435
   End
   Begin VB.Label lblInventory 
      AutoSize        =   -1  'True
      Caption         =   "Inventory - List of registered items on inventory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   4110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Items - List of items on Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2820
   End
End
Attribute VB_Name = "OfrmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbMonth_Click()
ViewItems "p_code", "%"
End Sub

Private Sub cbYear_Click()
ViewItems "ProductId", "%"
End Sub

Private Sub Form_Load()
LogStatus "Click Search for more search options"
For i = Format(Date, "yyyy") To 2000 Step -1
    cbYear.AddItem i
Next i
cbYear.AddItem "(View All)"
cbYear.Text = Val(Format(Date, "yyyy"))
cbMonth.ListIndex = Val(Format(Date, "mm"))
ViewItems "ProductID", "%"
ViewInven "ProductID", "%"
End Sub

Public Sub ViewItems(RcrdFld As String, RcrdStr As String)
If cbYear.Text = "(View All)" Or cbMonth.Text = "(none)" Then
    ExecuteSql "Select * from Products"
Else
    ExecuteSql "Select * from Products where " & RcrdFld & " LIKE '" & RcrdStr & "%' and format(reg_date, 'm') = " & _
            cbMonth.ListIndex & " and format(reg_date,'yyyy') = " & cbYear.Text & " Order by ProductID ASC"
End If
With rs
    lvwItems.ListItems.Clear
    While Not .EOF
        Set x = lvwItems.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
If NoRecords(lvwItems) Then Exit Sub
lvwItems.ListItems(1).Selected = True
End Sub
Public Sub ViewInven(RcrdFld As String, RcrdStr As String)
ExecuteSql "Select * from tblInventory where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
With rs
    lvwInventory.ListItems.Clear
    While Not .EOF
        Set x = lvwInventory.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
If NoRecords(lvwInventory) Then Exit Sub
lvwInventory.ListItems(1).Selected = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
lvwItems.Width = ScaleWidth - (lvwItems.Left + 100)
lvwItems.Height = (ScaleHeight - lvwItems.Top) / 2.5
lblInventory.Top = lvwItems.Height + lvwItems.Top + lvwItems.Left
lvwInventory.Top = lblInventory.Top + (lblInventory.Height * 2)
lvwInventory.Width = ScaleWidth - (lvwItems.Left + 100)
lvwInventory.Height = ScaleHeight - (lvwInventory.Top + 100)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'frmMain.tbrMenu.Buttons(2).Value = tbrUnpressed
ClearLogStatus
End Sub

Private Sub lvwInventory_DblClick()
If NoRecords(lvwInventory, "No items on your inventory list. Please search for an item.") Then Exit Sub
frmStocks.ExecSrch "ProductID", lvwInventory.SelectedItem
frmStocks.Show vbModal
End Sub

Private Sub lvwInventory_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu frmMain.mnuRegister
End If
End Sub

Private Sub lvwItems_DblClick()
If NoRecords(lvwItems, "No items on your item list. Please search for an item.") Then Exit Sub
frmAddItem.ExecSrch "ProductID", lvwItems.SelectedItem.SubItems(1)
frmAddItem.Show vbModal
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu frmMain.mnuInventory
End If
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As Button)
ExecButtons Button.Index
End Sub
Public Sub ExecButtons(Index As Integer)
Screen.MousePointer = 11
Select Case Index
    Case 1
        frmAddItem.Show vbModal
    Case 2
        lvwItems_DblClick
    Case 3
        If NoRecords(lvwItems, "No items on your item list. Please search for an item.") Then Exit Sub
        ExecuteSql "Select * from tblInventory where ProductID = '" & lvwItems.SelectedItem.SubItems(1) & "'"
        With rs
            Screen.MousePointer = 0
            If .EOF = False Then
                x = MsgBox("This item, " & lvwItems.SelectedItem.SubItems(1) & _
                            ", is registered on your inventory. Removing this item will affect your inventory records. " & _
                            vbNewLine & vbNewLine & "Do you want to continue?", vbExclamation + vbYesNo)
                If x = vbYes Then
                    ExecuteSql2 "select * from tblItems where ProductId = '" & lvwItems.SelectedItem.SubItems(1) & "'"
                    If rs2.Fields!image_name <> Empty Then
                        Kill App.Path & "\Images\Products\" & rs2.Fields!image_name
                    End If
                    rs2.Delete
                    MsgBox "Item " & lvwItems.SelectedItem.SubItems(1) & " has been deleted", vbInformation
                End If
            Else
                x = MsgBox("Your about to delete item " & lvwItems.SelectedItem.SubItems(1) & ", are you sure?", vbExclamation + vbYesNo)
                If x = vbYes Then
                    ExecuteSql2 "Select * from tblItems where ProductId = '" & lvwItems.SelectedItem.SubItems(1) & "'"
                    If rs2.Fields!image_name <> Empty Then
                        Kill App.Path & "\Images\Products\" & rs2.Fields!image_name
                    End If
                    rs2.Delete
                End If
            End If
                
        End With
        ViewItems "ProductId", "%"
        ViewInven "ProductId", "%"
    Case 4
        frmItemSettings.Show vbModal
    Case 5
        frmSrchOpt.Show vbModal
    Case 6
        If NoRecords(lvwItems, "No available record from the list. Please search for a record.") Then Exit Sub
        ExecuteSql "Select * from tblItems where ProductId = '" & lvwItems.SelectedItem.SubItems(1) & "' and on_inventory = 0"
        With rs
            If .EOF = False Then
                frmRegister.ViewUnreg "ProductId", lvwItems.SelectedItem.SubItems(1)
                frmRegister.lvwItems_DblClick
            End If
        End With
        frmRegister.Show vbModal
    Case 7
        frmStatus.ExecSrch "ProductId", lvwInventory.SelectedItem
        frmStatus.cmdLoad_Click
        frmStatus.Show vbModal
    Case 8
        Unload Me
End Select
Screen.MousePointer = 0
End Sub

