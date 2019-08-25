VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form OfrmSuppliers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suppliers"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7245
   Icon            =   "frmSuppliers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   7335
      TabIndex        =   26
      Top             =   840
      Width           =   7335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Delivery Scheduler"
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
      TabIndex        =   23
      Top             =   1680
      Width           =   1935
      Begin VB.ComboBox cboSchedType 
         Height          =   315
         ItemData        =   "frmSuppliers.frx":038A
         Left            =   240
         List            =   "frmSuppliers.frx":0397
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   300
         Width           =   1455
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
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   3975
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmSuppliers.frx":03B3
         Left            =   120
         List            =   "frmSuppliers.frx":03B5
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   200
         Width           =   1335
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
         ForeColor       =   &H8000000B&
         Height          =   300
         Left            =   1800
         TabIndex        =   8
         Text            =   "Search"
         Top             =   200
         Width           =   1695
      End
      Begin VB.PictureBox ctrlLiner2 
         Height          =   30
         Left            =   1560
         ScaleHeight     =   30
         ScaleWidth      =   135
         TabIndex        =   27
         Top             =   360
         Width           =   135
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3480
         Picture         =   "frmSuppliers.frx":03B7
         Top             =   105
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Supplier ID"
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
      Left            =   2160
      TabIndex        =   14
      Top             =   1680
      Width           =   1935
      Begin VB.Label lblId 
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
         Left            =   1320
         TabIndex        =   15
         Top             =   300
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Supplier Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   6975
      Begin MSMask.MaskEdBox txtCtel 
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtPno 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtPname 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtPemail 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtDescription 
         Height          =   615
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtCaddress 
         Height          =   735
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtCname 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   3960
         X2              =   3960
         Y1              =   2400
         Y2              =   1200
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   3960
         X2              =   4200
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
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
         Left            =   4320
         TabIndex        =   24
         Top             =   1320
         Width           =   525
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   3960
         X2              =   4200
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CP No.:"
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
         Left            =   4320
         TabIndex        =   22
         Top             =   1800
         Width           =   555
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         X1              =   3960
         X2              =   4200
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "* Tel No.:"
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
         TabIndex        =   21
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Personel"
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
         Left            =   3840
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
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
         Left            =   4320
         TabIndex        =   19
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "* Address:"
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
         TabIndex        =   13
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "* Company:"
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
         TabIndex        =   12
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "* Description:"
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
         TabIndex        =   11
         Top             =   1920
         Width           =   1155
      End
   End
   Begin MSComDlg.CommonDialog dlgPic 
      Left            =   6480
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Logo"
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
      Left            =   4200
      TabIndex        =   16
      Top             =   960
      Width           =   2895
      Begin VB.PictureBox cmdRemove 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   28
         Top             =   480
         Width           =   1215
      End
      Begin VB.PictureBox cmdBrowse 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.Image imgProfile 
         Height          =   1095
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000011&
         FillColor       =   &H80000004&
         Height          =   1095
         Left            =   1440
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
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
         Left            =   1680
         TabIndex        =   17
         Top             =   720
         Width           =   825
      End
   End
   Begin VB.PictureBox cmdClose 
      Height          =   375
      Left            =   4320
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   30
      Top             =   5640
      Width           =   1335
   End
   Begin VB.PictureBox cmdOptions 
      Height          =   375
      Left            =   5760
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   31
      Top             =   5640
      Width           =   1335
   End
   Begin VB.PictureBox cmdSave 
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   32
      Top             =   5640
      Width           =   1335
   End
   Begin VB.PictureBox cmdNew 
      Height          =   375
      Left            =   2880
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   33
      Top             =   5640
      Width           =   1335
   End
   Begin VB.PictureBox cmdDelete 
      Height          =   375
      Left            =   4320
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   34
      Top             =   8400
      Width           =   1335
   End
   Begin VB.PictureBox cmdEdit 
      Height          =   375
      Left            =   5760
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   35
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double Click to view a Supplier Profile"
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
      TabIndex        =   25
      Top             =   8565
      Width           =   2670
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   0
      Picture         =   "frmSuppliers.frx":0C81
      Top             =   8400
      Width           =   480
   End
End
Attribute VB_Name = "OfrmSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'set variables
Option Explicit
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

Private Sub cmdNew_Click()
ClrFlds
txtCname.SetFocus
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrMsg
ExecuteSql "Select * from tblSuppliers where record_id = '" & lstSuppliers.SelectedItem.SubItems(1) & "'"
If rs.EOF = False Then
    i = MsgBox("Delivery schedule of this supplier will also be removed. Click OK to procceed", vbQuestion + vbOKCancel)
    Screen.MousePointer = 11
    If i = vbOK Then
        If rs.Fields!image_name <> Empty Then
            Kill App.Path & "\Images\Suppliers\" & rs.Fields!image_name
        End If
        rs.Delete
        ClrFlds
        MsgBox "Account id " & lblId.Caption & " has been deleted", vbInformation
    Else
        Exit Sub
    End If
Else
    MsgBox "No specified supplier account found. Please search for an account try again", vbInformation
End If
Screen.MousePointer = 0
Exit Sub
ErrMsg:
    Screen.MousePointer = 0
    MsgBox ExecErr("You cannot delete this supplier. This record is present on your item records.", _
        "ProductId", "tblItems", "supplier", lstSuppliers.SelectedItem.SubItems(2)), vbCritical
End Sub

Private Sub cmdEdit_Click()
If lstSuppliers.ListItems.Count <> 0 Then
    ExecSrch "record_no", lstSuppliers.SelectedItem
End If
txtCname.SetFocus
End Sub

Private Sub cmdOptions_Click()
If cmdOptions.Caption = "&Options >>" Then
    Me.Height = 9405
    cmdOptions.Caption = "&Options <<"
Else
    Me.Height = 6645
    cmdOptions.Caption = "&Options >>"
End If
End Sub

Private Sub cmdRemove_Click()
imgProfile.Picture = LoadPicture()
ImgName = Empty
End Sub

Private Sub cmdSave_Click()
'set trapping functions
If TextBoxEmpty(txtCname) Then Exit Sub
If TextBoxEmpty(txtCaddress) Then Exit Sub
If TextBoxEmpty(txtDescription) Then Exit Sub
If TextBoxEmpty(txtCtel) Then Exit Sub
If TextBoxNumberEmpty(txtPno) Then Exit Sub
Screen.MousePointer = 11
'set the table name where supplier id is equal to the lblid
ExecuteSql "Select * from tblSuppliers where record_id = '" & lblId.Caption & "'"
With rs
    If cmdSave.Caption <> "&Update" Then
        msg = "Supplier id " & lblId.Caption & " has been successfully added."
        .AddNew
        .Fields!record_no = RcrdId("tblSuppliers")
        If ImgName <> Empty Then
            FileCopy ImgSrc, App.Path & "\Images\Suppliers\" & ImgName
        End If
    Else
        If ImgName <> Empty And .Fields!image_name = Empty Then
            FileCopy ImgSrc, App.Path & "\Images\Suppliers\" & ImgName
        End If
        If .Fields!image_name <> Empty And .Fields!image_name <> ImgName Then
            Kill App.Path & "\Images\Suppliers\" & .Fields!image_name
            If ImgName <> Empty Then
                FileCopy ImgSrc, App.Path & "\Images\Suppliers\" & ImgName
            End If
        End If
        msg = "Supplier id " & lblId.Caption & " has been successfully updated."
    End If
    .Fields!record_id = lblId.Caption
    .Fields!company = txtCname.Text
    .Fields!address = txtCaddress.Text
    .Fields!Description = txtDescription.Text
    .Fields!tel_no = txtCtel.Text
    .Fields!productId = txtPname.Text
    .Fields!p_email = txtPemail.Text
    .Fields!sched_type = cboSchedType.Text
    .Fields!last_delivery = Format(Date, "mm/dd/yyyy")
    .Fields!image_name = ImgName
    .Fields!date_reg = Format(Date, "mm/dd/yyyy")
    .Update
End With
Screen.MousePointer = 0
MsgBox msg, vbInformation
ClrFlds
ViewSup "record_no", "%"
End Sub

Public Sub ExecSrch(RcrdFld As String, RcrdStr As String)
ExecuteSql "Select * from tblSuppliers where " & RcrdFld & " LIKE '" & RcrdStr & "%' Order by record_no Asc"
With rs
    If .EOF = False Then
        lblId.Caption = .Fields!record_id
        cboSchedType.Text = .Fields!sched_type
        txtCname.Text = .Fields!company
        txtCaddress.Text = .Fields!address
        txtDescription.Text = .Fields!Description
        txtCtel.Text = .Fields!tel_no
        txtPname.Text = .Fields!p_name
        txtPno.Text = .Fields!p_cpno
        txtPemail.Text = .Fields!p_email
        cmdSave.Caption = "&Update"
        'set image filenames and values
        If .Fields!image_name <> Empty Then
            imgProfile.Picture = LoadPicture(App.Path & "\Images\Suppliers\" & .Fields!image_name)
        Else
            imgProfile.Picture = LoadPicture()
        End If
        ImgName = .Fields!image_name
    Else
        ClrFlds
    End If
End With
End Sub

Public Sub ViewSup(RcrdFld As String, RcrdStr As String)
ExecuteSql "Select * from tblSuppliers where " & RcrdFld & " LIKE '" & RcrdStr & "%' Order By record_no ASC"
With rs
    lstSuppliers.ListItems.Clear
    While Not .EOF
        Set x = lstSuppliers.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 3)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 11
'frmMain.cmdWarnings.Caption = Warnings & " Warnings"
Screen.MousePointer = 0
End Sub

Private Sub lstSuppliers_DblClick()
cmdEdit_Click
End Sub

Private Sub txtCtel_GotFocus()
'highlight the tel # when got focus
txtCtel.SelStart = 0
If txtCtel.Text <> "   -   -    " Then
    txtCtel.SelLength = Len(txtCtel)
End If
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    ExecSrch cboFilter.Text, txtSrchStr.Text
Else
    ClrFlds
End If
End Sub

Public Sub ClrFlds()
lblId.Caption = RcrdId("tblSuppliers", StrConv(Format(Date, "mmm"), vbUpperCase) & "-", "record_no")
dlgPic.FileName = Empty
cboSchedType.ListIndex = 0
ImgName = Empty
ImgSrc = Empty
txtCname.Text = Empty
txtCtel.Text = "   -   -    "
txtCaddress.Text = Empty
txtDescription.Text = Empty
txtPname.Text = Empty
txtPno.Text = Empty
txtPemail.Text = Empty
cmdSave.Caption = "&Save"
imgProfile.Picture = LoadPicture()
ViewSup "record_no", "%"
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
    txtSrchStr.ForeColor = &H8000000B
End If
End Sub

Private Sub Form_Load()
'SetListView lstSuppliers, True, True
lblId.Caption = RcrdId("tblSuppliers", StrConv(Format(Date, "mmm"), vbUpperCase) & "-", "record_no")
ExecuteSql "Select * from tblSuppliers"
With rs
    cboFilter.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).name)
    Next i
End With
cboFilter.Text = "company"
ClrFlds
LoadCombo "tblDeliverySched", cboSchedType, "description"
End Sub

