VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   Icon            =   "frmAddItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2520
      TabIndex        =   52
      Top             =   7080
      Width           =   735
   End
   Begin VB.CheckBox chkAutoBrand 
      Caption         =   "Brand name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      MaskColor       =   &H8000000F&
      TabIndex        =   46
      Top             =   6240
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkAutoLocation 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   45
      Top             =   6240
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkAutoType 
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   44
      Top             =   6240
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CheckBox chkVat 
      Height          =   255
      Left            =   1920
      TabIndex        =   43
      Top             =   5760
      Value           =   1  'Checked
      Width           =   240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Product Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   7695
      Begin VB.CommandButton cmdTypeSet 
         Height          =   375
         Left            =   3480
         TabIndex        =   47
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtVat 
         BackColor       =   &H8000000F&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   36
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtPid 
         Height          =   285
         Left            =   5400
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtValue 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmAddItem.frx":038A
         Left            =   1800
         List            =   "frmAddItem.frx":038C
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         ItemData        =   "frmAddItem.frx":038E
         Left            =   1800
         List            =   "frmAddItem.frx":0395
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3720
         Width           =   1575
      End
      Begin VB.ComboBox cboBrand 
         Height          =   315
         ItemData        =   "frmAddItem.frx":03A1
         Left            =   1800
         List            =   "frmAddItem.frx":03A3
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtUsage 
         Height          =   615
         Left            =   4320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtSupPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox cboLocation 
         Height          =   315
         ItemData        =   "frmAddItem.frx":03A5
         Left            =   1800
         List            =   "frmAddItem.frx":03A7
         TabIndex        =   5
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtBar 
         Height          =   285
         Left            =   5400
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtDescription 
         Height          =   435
         Left            =   4320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox txtUnitPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmdValueSet 
         Height          =   375
         Left            =   2640
         TabIndex        =   48
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdLocationSet 
         Height          =   375
         Left            =   3480
         TabIndex        =   49
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmdVatSet 
         Height          =   375
         Left            =   2640
         TabIndex        =   50
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdViewSup 
         Height          =   375
         Left            =   3480
         TabIndex        =   51
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Auto Save new record:"
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
         Left            =   4080
         TabIndex        =   35
         Top             =   3360
         Width           =   1905
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "* Value:"
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
         TabIndex        =   33
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "* Category:"
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
         TabIndex        =   32
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "* Supplier:"
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
         TabIndex        =   31
         Top             =   3720
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Product ID:"
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
         Left            =   4080
         TabIndex        =   30
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "* Brand Name:"
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
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "* Supplier Price:"
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
         TabIndex        =   28
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Usage:"
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
         Left            =   4080
         TabIndex        =   27
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "* Location:"
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
         TabIndex        =   26
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Bar Code: "
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
         Left            =   4080
         TabIndex        =   25
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label16 
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
         Left            =   4080
         TabIndex        =   24
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "* Unit Cost:"
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
         TabIndex        =   23
         Top             =   2280
         Width           =   960
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "* VAT:"
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
         TabIndex        =   22
         Top             =   3240
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Product Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   3015
      Begin VB.Label lblPcode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0000"
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
         TabIndex        =   18
         Top             =   240
         Width           =   420
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
      TabIndex        =   14
      Top             =   960
      Width           =   4575
      Begin VB.PictureBox ctrlLiner2 
         Height          =   30
         Left            =   1920
         ScaleHeight     =   30
         ScaleWidth      =   135
         TabIndex        =   34
         Top             =   360
         Width           =   135
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
         TabIndex        =   16
         Text            =   "Search"
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmAddItem.frx":03A9
         Left            =   240
         List            =   "frmAddItem.frx":03AB
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4080
         Picture         =   "frmAddItem.frx":03AD
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Total Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   600
         TabIndex        =   13
         Top             =   240
         Width           =   105
      End
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   9015
      TabIndex        =   11
      Top             =   840
      Width           =   9015
   End
   Begin MSComDlg.CommonDialog dlgPic 
      Left            =   7200
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Product Image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4800
      TabIndex        =   19
      Top             =   960
      Width           =   3015
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   1215
      End
      Begin VB.Image imgProfile 
         Height          =   1095
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000011&
         FillColor       =   &H80000004&
         Height          =   1095
         Left            =   1560
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "NO IMAGE"
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
         Left            =   1800
         TabIndex        =   20
         Top             =   720
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3600
      TabIndex        =   38
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   5040
      TabIndex        =   39
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   40
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmAddItem.frx":0C77
      Top             =   7080
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
      TabIndex        =   37
      Top             =   7200
      Width           =   1845
   End
End
Attribute VB_Name = "frmAddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Punit As String

Private Sub cboBrand_Change()
Description cboBrand.Text, cboType.Text, txtValue.Text
End Sub

Private Sub cboBrand_Click()
cboBrand_Change
End Sub

Private Sub cboFilter_Click()
txtSrchStr_Change
End Sub

Private Sub cboLocation_Change()
If cboLocation = Empty Then
    lblPcode.Caption = RcrdId("tblItems", "NEW-", "ProductId")
    Exit Sub
End If
lblPcode.Caption = RcrdId("tblItems", StrConv(Left(cboLocation.Text, 3), vbUpperCase) & "-", "ProductId")
End Sub

Private Sub cboLocation_Click()
cboLocation_Change
End Sub

Private Sub cboType_Change()
ExecuteSql "Select * from tblType where description = '" & cboType.Text & "'"
With rs
    If .EOF = False Then
        Punit = .Fields!unit
    Else
        Punit = "NEW"
    End If
End With
Description cboBrand.Text, cboType.Text, txtValue.Text
End Sub

Private Sub cboType_Click()
ExecuteSql "Select * from tblType where description = '" & cboType.Text & "'"
With rs
    If .EOF = False Then
        Punit = .Fields!unit
    Else
        Punit = "NEW"
    End If
End With
txtValue.Text = Empty
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo InvldPic
dlgPic.DialogTitle = "Load Profile Image"
dlgPic.InitDir = "My Documents"
dlgPic.filter = "Jepeg Image (*.jpg;*.jpeg)|*.jpg;*.jpeg|Bitmap Image (*.bmp)|*.bmp|All Files (*.*)|*.*"
dlgPic.ShowOpen
If dlgPic.FileName = "" Then
    Exit Sub
End If
imgProfile.Picture = LoadPicture(dlgPic.FileName)
ImgName = dlgPic.FileTitle
ImgSrc = dlgPic.FileName
Exit Sub
InvldPic:
    MsgBox "It is not a valid picture", vbExclamation
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdLocationSet_Click()
ExecuteSql "Select * from tblLocation where description = '" & cboLocation.Text & "'"
With rs
    If .EOF Then
        MsgBox "This Location is not yet save on database", vbExclamation
        Exit Sub
    Else
        With frmItemSettings
            .tabTables.Tabs(4).Selected = True
            .loadData "tblLocation", "description", cboLocation.Text
            .Show vbModal
        End With
    End If
End With
LoadCombo "tblLocation", cboLocation, "description"
End Sub

Private Sub cmdLoacationSet_Click()

End Sub

Private Sub cmdLocSet_Click()

End Sub

Private Sub cmdNew_Click()
ClrFlds
End Sub

Private Sub cmdRemove_Click()
imgProfile.Picture = LoadPicture()
ImgName = Empty
End Sub

Private Sub cmdSave_Click()

If TextBoxEmpty(cboBrand) Then Exit Sub
If TextBoxEmpty(cboType) Then Exit Sub
If TextBoxEmpty(txtValue) Then Exit Sub
If TextBoxEmpty(txtSupPrice) Then Exit Sub
If TextBoxNumberEmpty(txtSupPrice) Then Exit Sub
If TextBoxEmpty(txtUnitPrice) Then Exit Sub
If TextBoxNumberEmpty(txtUnitPrice) Then Exit Sub
If TextBoxEmpty(cboLocation) Then Exit Sub
If ComboEmpty(cboSupplier) Then Exit Sub

If chkAutoBrand.value = 1 Then
    ExecuteSql "select * from tblBrands where description = '" & Trim(cboBrand.Text) & "'"
    With rs
        If .EOF Then
            .AddNew
            .Fields!record_no = Val(RcrdId("tblBrands", , "record_no"))
            .Fields!Description = Trim(cboBrand.Text)
            .Fields!remarks = ""
            .Update
        End If
    End With
End If

If chkAutoType.value = 1 Then
    ExecuteSql "Select * from tblType where description = '" & Trim(cboType.Text) & "'"
    With rs
        If .EOF Then
            MsgBox "You've selected a new product type, please change it's default unit.", vbExclamation
            Punit = StrBox("Input the unit for this product type", imgIcon, , Punit, "type unit", 1, False)
            .AddNew
            .Fields!record_no = Val(RcrdId("tblType", , "record_no"))
            .Fields!Description = Trim(cboType.Text)
            .Fields!unit = Punit
            .Fields!remarks = ""
            .Update
        End If
    End With
End If

If chkAutoLocation.value = 1 Then
    ExecuteSql "Select * from tblLocation where description = '" & Trim(cboLocation.Text) & "'"
    With rs
        If .EOF Then
            .AddNew
            .Fields!record_no = Val(RcrdId("tblLocation", , "record_no"))
            .Fields!Description = Trim(cboLocation.Text)
            .Fields!remarks = ""
            .Update
        End If
    End With
End If

ExecuteSql "Select * from tblItems where ProductId = '" & lblPcode.Caption & "'"
With rs
    If cmdSave.Caption = "&Save" Then
        .AddNew
        .Fields!productId = Val(RcrdId("tblItems", , "ProductId"))
        .Fields!on_inventory = 0
        msg = "Added new item on database"
        If ImgName <> Empty Then
            FileCopy ImgSrc, App.Path & "\Images\Products\" & ImgName
        End If
    Else
        If ImgName <> Empty And .Fields!image_name = Empty Then
            FileCopy ImgSrc, App.Path & "\Images\Products\" & ImgName
        End If
        If .Fields!image_name <> Empty And .Fields!image_name <> ImgName Then
            Kill App.Path & "\Images\Products\" & .Fields!image_name
            If ImgName <> Empty Then
                FileCopy ImgSrc, App.Path & "\Images\Products\" & ImgName
            End If
        End If
        msg = "Product " & lblPcode.Caption & " has been updated"
    End If
    .Fields!productId = lblPcode.Caption
    .Fields!Description = txtDescription.Text
    .Fields!brand_name = cboBrand.Text
    .Fields!brand_type = cboType.Text
    .Fields!type_value = txtValue.Text
    .Fields!supplier_price = Val(txtSupPrice.Text)
    .Fields!unit_price = Val(txtUnitPrice.Text)
    If chkVat.value = 1 Then
        n = Val(txtUnitPrice.Text) + (Val(txtUnitPrice.Text) * (Val(txtVat.Text) / 100))
    Else
        n = Val(txtUnitPrice.Text)
    End If
    .Fields!net_price = n
    .Fields!location = cboLocation.Text
    .Fields!vat = chkVat.value
    .Fields!product_id = txtPid.Text
    .Fields!bar_code = txtBar.Text
    .Fields!Supplier = cboSupplier.Text
    .Fields!usage = txtUsage.Text
    .Fields!image_name = ImgName
    .Fields!reg_date = Format(Date, "mm/dd/yyyy")
    .Update
End With
MsgBox msg, vbInformation
frmInventory.ViewItems "ProductId", "%"
ExecuteSql "Select * from tblInventory where ProductId = '" & lblPcode.Caption & "'"
With rs
    If .EOF Then
        x = MsgBox("This item is not yet on your inventory list. Would you like to add this to Inventory?", vbQuestion + vbYesNo)
        If x = vbYes Then
            ExecuteSql2 "Select * from tblItems where ProductId = '" & lblPcode.Caption & "'"
            With rs2
                .Fields!on_inventory = 1
                .Update
            End With
            .AddNew
            .Fields!productId = lblPcode.Caption
            .Fields!Description = txtDescription.Text
            .Fields!quantity = 0
            .Fields!brand_name = cboBrand.Text
            .Fields!unit_price = Val(txtUnitPrice.Text)
            .Fields!net_price = n
            .Fields!bar_code = txtBar.Text
            .Fields!Supplier = cboSupplier.Text
            .Fields!sold = 0
            .Fields!Condition = StrBox("Select item status of this item", imgIcon, , "Select", "item status", 3, True, "tblStatus", "description")
            .Fields!location = cboLocation.Text
            .Fields!discount = 0
            .Fields!date_added = Format(Date, "mm/dd/yyyy")
            .Update
            x = MsgBox("Item " & lblPcode.Caption & " has been registered to inventory. Would you like to add the inventory stock of this item now?", vbInformation + vbYesNo)
            If x = vbYes Then
                With frmStocks
                    .ExecSrch "ProductId", lblPcode.Caption
                    .cmdView.Enabled = False
                    .Show vbModal
                End With
            End If
        Else
            MsgBox "You can add this item to your Inventory through item Register.", vbInformation
        End If
    Else
        .Fields!unit_price = Val(txtUnitPrice.Text)
        .Fields!net_price = n
        .Fields!Description = txtDescription.Text
        .Fields!brand_name = cboBrand.Text
        .Fields!bar_code = txtBar.Text
        .Fields!Supplier = cboSupplier.Text
        .Fields!location = cboLocation.Text
        .Update
    End If
End With
ClrFlds
Screen.MousePointer = 11
frmInventory.ViewInven "ProductId", "%"
Screen.MousePointer = 0
End Sub

Private Sub cmdTypeSet_Click()
ExecuteSql "Select * from tblType where description = '" & cboType.Text & "'"
With rs
    If .EOF Then
        MsgBox "This type is not yet save on database", vbExclamation
        Exit Sub
    Else
        With frmItemSettings
            .tabTables.Tabs(2).Selected = True
            .loadData "tblType", "description", cboType.Text
            .Show vbModal
        End With
    End If
End With
LoadCombo "tblType", cboType, "description"
End Sub

Private Sub cmdTypSet_Click()

End Sub

Private Sub cmdValueSet_Click()
ExecuteSql "Select * from tblType where description = '" & cboType.Text & "'"
With rs
    If .EOF Then
        MsgBox "Please select a product type first.", vbExclamation
        Exit Sub
    Else
        With frmItemSettings
            .tabTables.Tabs(2).Selected = True
            .loadData "tblType", "description", cboType.Text
            .Show vbModal
        End With
    End If
End With
LoadCombo "tblType", cboType, "description"
txtValue.Text = Empty
End Sub

Private Sub cmdVatSet_Click()
ExecuteSql "Select * from tblVat where value = " & Val(txtVat.Text) / 100
With rs
    If .EOF Then
        MsgBox "This Location is not yet save on database", vbExclamation
        Exit Sub
    Else
        With frmItemSettings
            .tabTables.Tabs(5).Selected = True
            .loadData "tblVat", "value", Val(txtVat.Text) / 100
            .Show vbModal
        End With
    End If
End With
ExecuteSql "Select * from tblVat"
txtVat.Text = (rs.Fields!value * 100) & " %"
End Sub

Private Sub cmdViewSup_Click()
If cboSupplier.ListIndex = 0 Then Exit Sub
With frmSuppliers
    .ExecSrch "company", cboSupplier.Text
    .cmdSave.Enabled = False
    .cmdNew.Enabled = False
    .cmdDelete.Enabled = False
    .cmdEdit.Enabled = False
    .Show vbModal
End With
ImgName = Empty
imgProfile.Picture = LoadPicture()
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub SetUp()
lblPcode.Caption = RcrdId("Products", "NEW-", "ProductID")
cboSupplier.ListIndex = 0

LoadCombo "tblType", cboType, "description"
LoadCombo "tblBrands", cboBrand, "description"
LoadCombo "tblLocation", cboLocation, "description"
LoadCombo "tblSuppliers", cboSupplier, "company"

ExecuteSql "Select * from tblItems"
lblCount.Caption = rs.RecordCount

ExecuteSql "Select * from tblVat"
txtVat.Text = (rs.Fields!value * 100) & " %"

ImgName = Empty
ImgSrc = Empty
End Sub

Private Sub Description(BrndNme As String, Ptype As String, Pvalue As String)
txtDescription = StrConv(BrndNme & " " & Pvalue & ", " & Ptype, vbUpperCase)
End Sub

Private Sub Form_Load()
ExecuteSql "Select * from Products"
With rs
    cboFilter.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).name)
    Next i
End With
'cboFilter = "description"
SetUp
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 11
'frmMain.cmdWarnings.Caption = Warnings & " Warnings"
Screen.MousePointer = 0
End Sub

Private Sub txtDescription_GotFocus()
SelectAll txtDescription
End Sub

Public Sub ExecSrch(SrchFld As String, SrchStr As String)
ExecuteSql "Select * from tblItems where " & SrchFld & " LIKE '" & SrchStr & "%'"
With rs
    If .EOF = False Then
        cboBrand.Text = .Fields!brand_name
        cboType.Text = .Fields!brand_type
        txtValue.Text = .Fields!type_value
        txtSupPrice.Text = .Fields!supplier_price
        txtUnitPrice.Text = .Fields!unit_price
        cboLocation.Text = .Fields!location
        chkVat.value = .Fields!vat
        txtPid.Text = .Fields!product_id
        txtBar.Text = .Fields!bar_code
        cboSupplier.Text = .Fields!Supplier
        txtUsage.Text = .Fields!usage
        lblPcode.Caption = .Fields!productId
        If .Fields!image_name <> Empty Then
            imgProfile.Picture = LoadPicture(App.Path & "\Images\Products\" & .Fields!image_name)
        Else
            imgProfile.Picture = LoadPicture()
        End If
        ImgName = .Fields!image_name
        txtDescription.Text = .Fields!Description
        cmdSave.Caption = "&Update"
    Else
        ClrFlds
    End If
End With
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ExecSrch cboFilter.Text, txtSrchStr.Text
        frmInventory.ViewItems cboFilter.Text, txtSrchStr.Text
    End If
Else
    ClrFlds
    frmInventory.ViewItems "ProductId", "%"
End If
End Sub

Private Sub txtValue_Change()
Description cboBrand.Text, cboType.Text, txtValue.Text
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
SetUp
txtDescription.Text = Empty
cboBrand.Text = Empty
txtValue.Text = Empty
txtSupPrice.Text = Empty
txtUnitPrice.Text = Empty
cboLocation.Text = Empty
txtPid.Text = Empty
txtBar.Text = Empty
txtUsage.Text = Empty
txtVat.Text = 12 & " %"
cmdSave.Caption = "&Save"
imgProfile.Picture = LoadPicture()
End Sub

Private Sub txtValue_GotFocus()
txtValue.Text = Val(txtValue.Text)
SelectAll txtValue
End Sub

Private Sub txtValue_LostFocus()
txtValue.Text = txtValue.Text & " " & Punit
End Sub
