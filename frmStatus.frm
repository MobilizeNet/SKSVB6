VERSION 5.00
Begin VB.Form OfrmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Date Added"
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
      Left            =   4560
      TabIndex        =   14
      Top             =   960
      Width           =   2415
      Begin VB.Label lblDateAdded 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "---"
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
         Left            =   1200
         TabIndex        =   15
         Top             =   360
         Width           =   225
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
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   6855
      Begin VB.TextBox txtRemarks 
         Height          =   615
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1800
         Width           =   4935
      End
      Begin VB.ComboBox cboCondition 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         ItemData        =   "frmStatus.frx":038A
         Left            =   4800
         List            =   "frmStatus.frx":038C
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtPcode 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.PictureBox cmdViewSup 
         Height          =   375
         Left            =   6360
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   31
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox cmdViewItem 
         Height          =   375
         Left            =   3240
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   28
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblLocation 
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
         Left            =   4800
         TabIndex        =   22
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
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
         Index           =   4
         Left            =   3840
         TabIndex        =   21
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks:"
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
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label lblBrand 
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
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Brand Name:"
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
         TabIndex        =   12
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblQty 
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
         Left            =   4800
         TabIndex        =   11
         Top             =   840
         Width           =   180
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   3840
         TabIndex        =   10
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplier:"
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
         Index           =   3
         Left            =   3840
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Item Code:"
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
         TabIndex        =   6
         Top             =   360
         Width           =   930
      End
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
      TabIndex        =   1
      Top             =   960
      Width           =   4335
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmStatus.frx":038E
         Left            =   240
         List            =   "frmStatus.frx":0390
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1215
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
         Left            =   1800
         TabIndex        =   3
         Text            =   "Search"
         Top             =   240
         Width           =   1935
      End
      Begin VB.PictureBox ctrlLiner2 
         Height          =   30
         Left            =   1560
         ScaleHeight     =   30
         ScaleWidth      =   135
         TabIndex        =   2
         Top             =   360
         Width           =   135
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3720
         Picture         =   "frmStatus.frx":0392
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   8295
      TabIndex        =   0
      Top             =   840
      Width           =   8295
   End
   Begin VB.PictureBox cmdView 
      Height          =   375
      Left            =   5760
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   25
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox cmdClear 
      Height          =   375
      Left            =   4440
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   26
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox cmdUpdate 
      Height          =   375
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   27
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox cmdClose 
      Height          =   375
      Left            =   5760
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   29
      Top             =   7440
      Width           =   1215
   End
   Begin VB.PictureBox cmdLoad 
      Height          =   375
      Left            =   4440
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   30
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double click an item from the list to load details."
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
      TabIndex        =   24
      Top             =   7560
      Width           =   3405
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   0
      Picture         =   "frmStatus.frx":0C5C
      Top             =   7440
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List of registered items on Inventory"
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
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   5280
      Width           =   3120
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmStatus.frx":1526
      Top             =   4680
      Width           =   480
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Indecates required field"
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
      TabIndex        =   20
      Top             =   4815
      Width           =   1845
   End
End
Attribute VB_Name = "OfrmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
ClrFlds
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdLoad_Click()
If NoRecords(lstInventory, "No record available on your inventory list. Please search for an item.") Then Exit Sub
ExecuteSql "Select ProductId, quantity, brand_name, supplier, condition, location, date_added  from tblInventory where ProductId = '" & lstInventory.SelectedItem & "'"
With rs
    txtPcode.Text = .Fields!productId
    lblBrand.Caption = .Fields!brand_name
    cboCondition.Text = .Fields!Condition
    cboSupplier.Text = .Fields!Supplier
    lblQty.Caption = .Fields!quantity
    lblLocation.Caption = .Fields!location
    lblDateAdded.Caption = Format(.Fields!date_added, "mm/dd/yyyy")
    ExecuteSql2 "Select * from tblItemCondition where ProductId = '" & lstInventory.SelectedItem & "'"
    With rs2
        If .EOF = False Then
            txtRemarks.Text = .Fields!remarks
        Else
            txtRemarks.Text = Empty
        End If
    End With
End With
End Sub

Private Sub cmdUpdate_Click()
If txtPcode.Text = Empty Then
    MsgBox "No espicified item. Click Load to load an item from your inventory list", vbExclamation
    Exit Sub
End If
ExecuteSql "Select condition from tblInventory where ProductId = '" & txtPcode.Text & "'"
With rs
    .Fields!Condition = cboCondition.Text
    .Update
End With
ExecuteSql2 "Select * from tblItemCondition where ProductId = '" & txtPcode.Text & "'"
With rs2
    If .EOF Then
        If Trim(txtRemarks.Text) <> Empty Then
            .AddNew
            .Fields!record_no = RcrdId("tblItemCondition", , "record_no")
            .Fields!productId = txtPcode.Text
            .Fields!Condition = cboCondition.Text
            .Fields!remarks = txtRemarks.Text
            .Fields!date_updated = Format(Date, "mm/dd/yyyy")
            .Update
        End If
    Else
        .Fields!remarks = txtRemarks.Text
        .Fields!date_updated = Format(Date, "mm/dd/yyyy")
        .Update
    End If
End With
MsgBox "Item " & txtPcode.Text & "'s condition has been updated.", vbInformation
ExecSrch "ProductId", "%"
frmInventory.ViewInven "ProductId", "%"
End Sub

Private Sub cmdView_Click()
If cmdView.Caption = "&View <<" Then
    Me.Height = 5595
    cmdView.Caption = "&View >>"
Else
    Me.Height = 8355
    cmdView.Caption = "&View <<"
End If
End Sub

Private Sub cmdViewItem_Click()
If txtPcode.Text = Empty Then Exit Sub
Screen.MousePointer = 11
frmAddItem.cmdSave.Enabled = False
frmAddItem.cmdNew.Enabled = False
frmAddItem.ExecSrch "ProductId", txtPcode.Text
frmAddItem.Show vbModal
End Sub

Private Sub cmdViewSup_Click()
If cboSupplier.Text = Empty Then Exit Sub
Screen.MousePointer = 11
With frmSuppliers
    .ExecSrch "company", cboSupplier.Text
    .cmdSave.Enabled = False
    .cmdNew.Enabled = False
    .cmdDelete.Enabled = False
    .cmdEdit.Enabled = False
    .Show vbModal
End With
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Public Sub ExecSrch(RcrdFld As String, RcrdStr As String)
ExecuteSql "Select ProductId, quantity, brand_name, supplier, condition, location, date_added  from tblInventory where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
With rs
    lstInventory.ListItems.Clear
    While Not .EOF
        Set x = lstInventory.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Private Sub Form_Load()
'SetListView lstInventory, True, True
ExecuteSql "Select ProductId, quantity, brand_name, supplier, condition, location  from tblInventory"
With rs
    cboFilter.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).name)
    Next i
End With
cboFilter.Text = "ProductId"

LoadCombo "tblStatus", cboCondition, "description"
LoadCombo "tblSuppliers", cboSupplier, "company"

ExecSrch "ProductId", "%"
End Sub

Private Sub lstInventory_DblClick()
cmdLoad_Click
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ExecSrch cboFilter.Text, txtSrchStr.Text
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

Private Sub ClrFlds()
txtPcode.Text = Empty
lblBrand.Caption = "---"
cboCondition.ListIndex = 0
txtRemarks.Text = Empty
cboSupplier.ListIndex = 0
lblQty.Caption = "---"
lblLocation.Caption = "---"
lblDateAdded.Caption = "---"
ExecSrch "ProductId", "%"
End Sub
