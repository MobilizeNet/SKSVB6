VERSION 5.00
Begin VB.Form OfrmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSrchStrInven 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Php""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   2
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   2040
      TabIndex        =   16
      Text            =   "Search"
      Top             =   6000
      Width           =   2895
   End
   Begin VB.ComboBox cboFilterInven 
      Height          =   315
      ItemData        =   "frmRegister.frx":038A
      Left            =   120
      List            =   "frmRegister.frx":038C
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5295
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmRegister.frx":038E
         Left            =   240
         List            =   "frmRegister.frx":0390
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtSrchStr 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   300
         Left            =   2160
         TabIndex        =   5
         Text            =   "Search"
         Top             =   240
         Width           =   2655
      End
      Begin VB.PictureBox ctrlLiner2 
         Height          =   30
         Left            =   1920
         ScaleHeight     =   30
         ScaleWidth      =   135
         TabIndex        =   4
         Top             =   360
         Width           =   135
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4800
         Picture         =   "frmRegister.frx":0392
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   5295
      Begin VB.ComboBox cboCondition 
         Height          =   315
         ItemData        =   "frmRegister.frx":0C5C
         Left            =   1560
         List            =   "frmRegister.frx":0C5E
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtDescription 
         Height          =   375
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1320
         Width           =   3615
      End
      Begin VB.PictureBox cmdViewCondition 
         Height          =   375
         Left            =   3240
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   25
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "* Condition:"
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
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "|"
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
         Left            =   2760
         TabIndex        =   14
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   4680
         TabIndex        =   13
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "On Hand:"
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
         Left            =   3720
         TabIndex        =   12
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
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
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label lblPcode 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "P-Code:"
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
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   840
      Width           =   7575
   End
   Begin VB.PictureBox ctrlLiner3 
      Height          =   30
      Left            =   1800
      ScaleHeight     =   30
      ScaleWidth      =   135
      TabIndex        =   17
      Top             =   6120
      Width           =   135
   End
   Begin VB.PictureBox cmdRemove 
      Height          =   375
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   20
      Top             =   8280
      Width           =   1215
   End
   Begin VB.PictureBox cmdView 
      Height          =   375
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   21
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox cmdClear 
      Height          =   375
      Left            =   1560
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   22
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox cmdLoad 
      Height          =   375
      Left            =   2880
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   23
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox cmdReg 
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   24
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   4920
      Picture         =   "frmRegister.frx":0C60
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Remove to unregister an item from list"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   8400
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmRegister.frx":152A
      Top             =   8280
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List of unregistered Items on Database."
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
      TabIndex        =   7
      Top             =   3840
      Width           =   3360
   End
End
Attribute VB_Name = "OfrmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClrFlds()
lblPcode.Caption = "---"
lblPrice.Caption = "---"
txtDescription.Text = Empty
ViewUnreg "ProductId", "%"
ViewInventory "ProductId", "%"
End Sub

Private Sub cboFilter_Click()
txtSrchStr_Change
End Sub

Private Sub cboFilterInven_Click()
txtSrchStrInven_Change
End Sub

Private Sub cmdClear_Click()
ClrFlds
End Sub

Private Sub cmdLoad_Click()
lvwItems_DblClick
End Sub

Private Sub cmdReg_Click()
If lblPcode.Caption = "---" Then MsgBox "Please load an item from the list.", vbExclamation: Exit Sub
If ComboEmpty(cboCondition) Then Exit Sub
ExecuteSql "Select * from tblInventory"
With rs
    ExecuteSql2 "Select * from tblItems where ProductId = '" & lblPcode.Caption & "'"
    rs2.Fields!on_inventory = 1
    rs2.Update
    .AddNew
    .Fields!productId = rs2.Fields!productId
    .Fields!Description = rs2.Fields!Description
    .Fields!brand_name = rs2.Fields!brand_name
    .Fields!unit_price = rs2.Fields!unit_price
    .Fields!net_price = rs2.Fields!net_price
    .Fields!bar_code = rs2.Fields!bar_code
    .Fields!Supplier = rs2.Fields!Supplier
    .Fields!sold = 0
    .Fields!Condition = cboCondition.Text
    .Fields!location = rs2.Fields!location
    .Fields!date_added = Format(Date, "mm/dd/yyyy")
    ExecuteSql2 "Select * from tblUnregStocks where ProductId = '" & lblPcode.Caption & "'"
    If rs2.EOF = False Then
        .Fields!quantity = rs2.Fields!stored_qty
        rs2.Delete
    Else
        .Fields!quantity = 0
    End If
    .Update
End With
frmInventory.ViewInven "ProductId", "%"
frmInventory.ViewItems "ProductId", "%"
MsgBox "Item " & lblPcode.Caption & " has been successfully registered to Inventory.", vbInformation
ClrFlds
End Sub

Public Sub ViewInventory(RcrdFld As String, RcrdStr As String)
ExecuteSql "Select ProductId, description, net_price, quantity, location from tblInventory where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
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
End Sub

Private Sub cmdRemove_Click()
ExecRemove lvwInventory.SelectedItem
End Sub

Public Sub ExecRemove(Pcode As String)
If NoRecords(lvwInventory, "No record available on the list. Please Search for an item.") Then Exit Sub
If MsgBox("Your about to remove item " & Pcode & " from your inventory list." & _
            " Removing this item may affect system report." & vbNewLine & vbNewLine & _
            "Do you want to continue?", vbExclamation + vbOKCancel) = vbOK Then
    s = MsgBox("Would you like to save the inventory quantity of this item?", vbExclamation + vbYesNoCancel)
    If s <> vbCancel Then
        If s = vbYes Then
            ExecuteSql "Select * from tblUnregStocks"
            With rs
                .AddNew
                .Fields!productId = Pcode
                .Fields!stored_qty = Val(lvwInventory.SelectedItem.SubItems(3))
                .Update
            End With
        End If
        ExecuteSql "Delete * from tblInventory where ProductId = '" & Pcode & "'"
        ExecuteSql "Select on_inventory from tblItems where ProductId = '" & Pcode & "'"
        rs.Fields!on_inventory = 0
        rs.Update
        MsgBox "Item " & Pcode & " has been removed from inventory.", vbInformation
    End If
End If
ClrFlds
frmInventory.ViewInven "ProductId", "%"
frmInventory.ViewItems "ProductId", "%"
End Sub

Private Sub cmdView_Click()
If cmdView.Caption = "&Inventory >>" Then
    Me.Height = 9255
    cmdView.Caption = "&Inventory <<"
Else
    Me.Height = 6375
    cmdView.Caption = "&Inventory >>"
End If
End Sub

Private Sub cmdViewCondition_Click()
If ComboEmpty(cboCondition) Then Exit Sub
frmItemSettings.loadData "tblStatus", "description", cboCondition.Text
frmItemSettings.Show vbModal
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
ExecuteSql "Select ProductId, ProductId, description, net_price from tblItems"
With rs
    cboFilter.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).name)
    Next i
End With
cboFilter.Text = "description"

ExecuteSql "Select ProductId, description, net_price, quantity, location from tblInventory"
With rs
    cboFilterInven.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilterInven.AddItem (.Fields(i).name)
    Next i
End With
cboFilterInven.Text = "description"

'SetListView lvwItems, True, True
'SetListView lvwInventory, True, True
ViewUnreg "ProductId", "%"
ViewInventory "ProductId", "%"
LoadCombo "tblStatus", cboCondition, "description"
End Sub

Public Sub ViewUnreg(RcrdFld As String, RcrdStr As String)
ExecuteSql "Select ProductId, ProductId, description, net_price from tblItems where " & RcrdFld & " LIKE '" & RcrdStr & "%' and on_inventory = 0 Order By ProductId ASC"
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 11
'frmMain.cmdWarnings.Caption = Warnings & " Warnings"
Screen.MousePointer = 0
End Sub

Public Sub lvwItems_DblClick()
If NoRecords(lvwItems, "No record available on the list. Please Search for an item.") Then Exit Sub
ExecuteSql "Select ProductId, description, net_price from tblItems where ProductId = '" & lvwItems.SelectedItem.SubItems(1) & "'"
With rs
    lblPcode.Caption = .Fields!productId
    txtDescription.Text = .Fields!Description
    lblPrice.Caption = "P " & Format(.Fields!net_price, "#0.00")
End With
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ViewUnreg cboFilter.Text, txtSrchStr.Text
    End If
Else
    ClrFlds
End If
End Sub

Private Sub txtSrchStr_GotFocus()
If txtSrchStr = "Search" Then
    txtSrchStr.Text = Empty
    txtSrchStr.ForeColor = &H80000008
End If
End Sub

Private Sub txtSrchStr_LostFocus()
If Trim(txtSrchStr) = Empty Then
    txtSrchStr.Text = "Search"
    txtSrchStr.ForeColor = &H80000011
End If
End Sub

Private Sub txtSrchStrInven_Change()
If Right(txtSrchStrInven.Text, 1) = "'" Then
    txtSrchStrInven.Text = Empty
End If
If Trim(txtSrchStrInven.Text) <> Empty Then
    If txtSrchStrInven.Text <> "Search" Then
        ViewInventory cboFilterInven.Text, txtSrchStrInven.Text
    End If
Else
    ClrFlds
End If
End Sub

Private Sub txtSrchStrInven_GotFocus()
If txtSrchStrInven = "Search" Then
    txtSrchStrInven.Text = Empty
    txtSrchStrInven.ForeColor = &H80000008
End If
End Sub

Private Sub txtSrchStrInven_LostFocus()
If Trim(txtSrchStrInven) = Empty Then
    txtSrchStrInven.Text = "Search"
    txtSrchStrInven.ForeColor = &H80000011
End If
End Sub
