VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form OfrmAccountProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts"
   ClientHeight    =   7935
   ClientLeft      =   3930
   ClientTop       =   1965
   ClientWidth     =   7350
   Icon            =   "OfrmAccountProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   7935
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   7575
      TabIndex        =   44
      Top             =   840
      Width           =   7575
   End
   Begin VB.Frame Frame7 
      Caption         =   "ID Count"
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
      Left            =   3120
      TabIndex        =   41
      Top             =   1800
      Width           =   1095
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   480
         TabIndex        =   43
         Top             =   240
         Width           =   120
      End
      Begin VB.Label Label13 
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
         Left            =   1620
         TabIndex        =   42
         Top             =   280
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog dlgPic 
      Left            =   6600
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      Caption         =   "PC NAME"
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
      TabIndex        =   37
      Top             =   960
      Width           =   1815
      Begin VB.Label lblPcId 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PC ID"
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
         Left            =   1050
         TabIndex        =   39
         Top             =   285
         Width           =   450
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   3240
         TabIndex        =   38
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "User ID"
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
      Left            =   2040
      TabIndex        =   32
      Top             =   960
      Width           =   2175
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
         Left            =   1620
         TabIndex        =   33
         Top             =   280
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Employee Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   16
      Top             =   5640
      Width           =   7095
      Begin VB.ComboBox cboEmpStatus 
         Height          =   315
         ItemData        =   "OfrmAccountProfile.frx":038A
         Left            =   1680
         List            =   "OfrmAccountProfile.frx":0397
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cboPosition 
         Height          =   315
         ItemData        =   "OfrmAccountProfile.frx":03B9
         Left            =   5040
         List            =   "OfrmAccountProfile.frx":03C0
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cboRemarks 
         Height          =   315
         ItemData        =   "OfrmAccountProfile.frx":03CC
         Left            =   5040
         List            =   "OfrmAccountProfile.frx":03D6
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   960
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpDHired 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-mmm-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   12
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   81920003
         CurrentDate     =   40071
      End
      Begin MSMask.MaskEdBox txtDhired 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblEmpStatus 
         AutoSize        =   -1  'True
         Caption         =   "* Emp. Stat.:"
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
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Position:"
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
         TabIndex        =   28
         Top             =   480
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "* Date Hired:"
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
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "* Remarks:"
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
         TabIndex        =   26
         Top             =   960
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Personal Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   7095
      Begin VB.TextBox txtContact 
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
         Left            =   5040
         TabIndex        =   10
         Top             =   2400
         Width           =   1815
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "OfrmAccountProfile.frx":03EE
         Left            =   5040
         List            =   "OfrmAccountProfile.frx":03F8
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox cboGender 
         Height          =   315
         ItemData        =   "OfrmAccountProfile.frx":040D
         Left            =   5040
         List            =   "OfrmAccountProfile.frx":041A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtAge 
         Alignment       =   1  'Right Justify
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
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   6
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtLname 
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
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtFname 
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
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtMname 
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
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Height          =   735
         Left            =   5040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpBdate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-mmm-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Top             =   1920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   81920003
         CurrentDate     =   40071
      End
      Begin MSMask.MaskEdBox txtBdate 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         TabIndex        =   30
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "* Birth Date:"
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
         TabIndex        =   25
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "* Gender:"
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
         Left            =   3600
         TabIndex        =   24
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "* Age:"
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
         Top             =   2400
         Width           =   525
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "* Surname:"
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
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "* Middle Name:"
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
         TabIndex        =   21
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "* First Name:"
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
         TabIndex        =   20
         Top             =   480
         Width           =   1095
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
         Left            =   3600
         TabIndex        =   19
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Contact #:"
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
         TabIndex        =   18
         Top             =   2400
         Width           =   885
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
      TabIndex        =   34
      Top             =   1800
      Width           =   2895
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
         Left            =   720
         TabIndex        =   0
         Text            =   "Search"
         Top             =   200
         Width           =   1695
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
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
         TabIndex        =   35
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   2400
         Picture         =   "OfrmAccountProfile.frx":0434
         Top             =   100
         Width           =   480
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Profile Picture"
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
      Left            =   4320
      TabIndex        =   36
      Top             =   960
      Width           =   2895
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000011&
         FillColor       =   &H80000004&
         Height          =   1095
         Left            =   1440
         Top             =   240
         Width           =   1335
      End
      Begin VB.Image imgProfile 
         Height          =   1095
         Left            =   1440
         Stretch         =   -1  'True
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
         TabIndex        =   40
         Top             =   720
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3000
      TabIndex        =   45
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Required field"
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
      TabIndex        =   29
      Top             =   7560
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "OfrmAccountProfile.frx":0CFE
      Top             =   7425
      Width           =   480
   End
End
Attribute VB_Name = "OfrmAccountProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdDelete_Click()
'locate a user id and delete if found
ExecuteSql "Select * from tblAccountProfile where id = '" & lblId.Caption & "'"
With rs
    'if no record...
    If .EOF Then
        MsgBox "No specified account to delete. Please search for an account", vbInformation
        ClrFlds
        txtSrchStr.SetFocus
    Else
    'else if a record found
        'test if the image is not empty then delete from accounts images
        If imgProfile.Picture <> LoadPicture() Then
            Me.MousePointer = 11
            Kill App.Path & "\Images\accounts\" & .Fields!image_name
        End If
        'delete the account
        .Delete
        'clear all
        ClrFlds
        MsgBox "Account id " & lblId.Caption & " has been deleted", vbInformation
    End If
End With
'test if there is no account left in the database, then exit system
ExecuteSql "Select * from tblAccountProfile"
If rs.RecordCount = 0 Then
    MsgBox "All accounts has been deleted. The system will now exit", vbCritical
    'exit
    End
End If
End Sub


Private Sub cmdRemove_Click()
imgProfile.Picture = LoadPicture()
ImgName = Empty
End Sub

Private Sub cmdSave_Click()
Dim msg As String

'set functions in trappings
If TextBoxEmpty(txtFname) Then Exit Sub
If TextBoxEmpty(txtMname) Then Exit Sub
If TextBoxEmpty(txtLname) Then Exit Sub
If ComboEmpty(cboGender) Then Exit Sub
If TextBoxEmpty(txtBdate) Then Exit Sub: SelectAll (txtBdate)
If TextBoxEmpty(txtAddress) Then Exit Sub
If TextBoxEmpty(txtDhired) Then Exit Sub
If ComboEmpty(ComboEmptyStatus) Then Exit Sub
If ComboEmpty(cboPosition) Then Exit Sub

'select an specified account
ExecuteSql "Select * from tblAccountProfile where id = '" & lblId.Caption & "'"
With rs
    'if a command button is Save then add new record
    'else update the record
    If cmdSave.Caption <> "&Update" Then
        msg = "Account id " & lblId.Caption & " has been successfully added."
        .AddNew
        .Fields!account_no = Val(RcrdId("tblAccountProfile", , "account_no"))
        If ImgName <> Empty Then
            FileCopy ImgSrc, App.Path & "\Images\Accounts\" & ImgName
        End If
    Else
        If ImgName <> Empty And .Fields!image_name = Empty Then
            FileCopy ImgSrc, App.Path & "\Images\Accounts\" & ImgName
        End If
        If .Fields!image_name <> Empty And .Fields!image_name <> ImgName Then
            Kill App.Path & "\Images\Accounts\" & .Fields!image_name
            If ImgName <> Empty Then
                FileCopy ImgSrc, App.Path & "\Images\Accounts\" & ImgName
            End If
        End If
        msg = "Account id " & lblId.Caption & " has been successfully updated."
    End If
    Me.MousePointer = 11
    .Fields!id = lblId.Caption
    .Fields!fname = txtFname.Text
    .Fields!mname = txtMname.Text
    .Fields!lname = txtLname.Text
    .Fields!age = Val(txtAge)
    .Fields!gender = cboGender.Text
    .Fields!bdate = txtBdate.Text
    .Fields!address = txtAddress.Text
    .Fields!Status = cboStatus.Text
    .Fields!contact = Val(txtContact)
    .Fields!date_hired = txtDhired.Text
    .Fields!Position = cboPosition.Text
    .Fields!remarks = cboRemarks.Text
    .Fields!emp_status = ComboEmptyStatus.Text
    .Fields!image_name = ImgName
    .Fields!date_reg = Format(Date, "mm/dd/yyyy")
    .Update
    i = MsgBox(msg & " Do you want to set the security of this Account?", vbQuestion + vbYesNo)
    If i = vbYes Then
        frmAcntManage.Show vbModal
    End If
    ClrFlds
End With
If CurrentUserAdmin Then
    Unload Me
End If
End Sub

Private Sub cmdSecurity_Click()
'show security manager and set security accounts
ExecuteSql "Select * from tblAccountProfile where id = '" & lblId.Caption & "'"
If rs.EOF = False Then
    frmAcntManage.txtId = lblId.Caption
    frmAcntManage.Show vbModal
Else
    frmAcntManage.Show vbModal
End If
End Sub

Private Sub dtpBdate_Change()
'format txtbdate as short date
txtBdate.Text = Format(dtpBdate.value, "mm/dd/yyyy")
End Sub

Private Sub dtpBdate_LostFocus()
'if age is below zero, invalid bdate
If Val(txtAge.Text) <= 10 Then
    MsgBox "Invalid date: Birth date must be less than current date", vbExclamation
    txtBdate.Text = "  /  /    "
    txtAge.Text = Empty
End If
End Sub

Private Sub dtpDHired_Change()
'change format to short date
txtDhired.Text = Format(dtpDHired.value, "mm/dd/yyyy")
End Sub

Private Sub Form_Load()
'pc name
''lblPcId.Caption = PcId
ClrFlds

'add the saved postions from database
ExecuteSql "Select * from tblAccountPosition"
While Not rs.EOF
    cboPosition.AddItem (rs.Fields!Description)
    rs.MoveNext
Wend

ExecuteSql "Select* from tblAccountProfile"
lblCount.Caption = rs.RecordCount
lblId.Caption = RcrdId("tblAccountProfile", "NEW-", "account_no")
End Sub
Public Sub ClrFlds()
'clear all
Me.MousePointer = 0
dlgPic.FileName = Empty
ImgName = Empty
ImgSrc = Empty
imgProfile.Picture = LoadPicture()
txtFname.Text = Empty
txtMname.Text = Empty
txtLname.Text = Empty
txtBdate.Text = "  /  /    "
txtAge.Text = Empty
txtAddress.Text = Empty
txtContact.Text = Empty
txtDhired.Text = "  /  /    "
cboGender.ListIndex = 0
cboStatus.ListIndex = 0
cboPosition.ListIndex = 0
cboRemarks.ListIndex = 0
ComboEmptyStatus.ListIndex = 0
cmdSave.Caption = "&Save"
lblId.Caption = RcrdId("tblAccountProfile", "NEW-", "account_no")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If CurrentUserAdmin Then
    ExecuteSql "Select * from tblAccountProfile"
    If rs.EOF Then
        MsgBox "System has failed to initialized. Please contact your administrator" & vbNewLine _
                & vbNewLine & "Status: analysing accounts configuration" & vbNewLine & _
                "Error: No account found", vbCritical
        End
    End If
End If
End Sub

Private Sub txtAddress_LostFocus()
'the function that converts text to proper cases
txtAddress.Text = StrConv(txtAddress, vbProperCase)
End Sub

Private Sub txtBdate_Change()
'birth date conditions and trappings
txtAge.Text = Val(Format(Date, "YYYY")) - Val(Format(txtBdate, "YYYY"))

If Format(txtBdate, "MM") = Format(Date, "MM") Then
    If Format(txtBdate, "DD") > Format(Date, "DD") Then
        txtAge.Text = txtAge.Text - 1
    End If
End If

If Format(txtBdate, "MM") > Format(Date, "MM") Then
    txtAge.Text = txtAge.Text - 1
End If
End Sub

'---------------------highlight if got focus-------------
Private Sub txtBdate_GotFocus()
txtBdate.SelStart = 0
If txtBdate.Text <> "  /  /    " Then
    txtBdate.SelLength = Len(txtBdate)
End If
End Sub

Private Sub txtDhired_GotFocus()
txtDhired.SelStart = 0
If txtDhired.Text = "  /  /    " Then
    txtDhired.SelLength = Len(txtDhired)
End If
End Sub

Private Sub txtFname_LostFocus()
txtFname.Text = StrConv(txtFname, vbProperCase)
End Sub

Private Sub txtLname_Change()
lblId.Caption = RcrdId("tblAccountProfile", StrConv(Left(txtLname.Text, 3), vbUpperCase) & "-", "account_no")
End Sub

Private Sub txtLname_LostFocus()
txtLname.Text = StrConv(txtLname, vbProperCase)
End Sub

Private Sub txtMname_LostFocus()
txtMname.Text = StrConv(txtMname, vbProperCase)
End Sub
'------------------end hghlights----------
Private Sub txtSrchStr_Change()
If Trim(txtSrchStr.Text) <> Empty Then
    ExecSrch (txtSrchStr.Text)
Else
    ClrFlds
End If
End Sub
Public Sub ExecSrch(ByVal id As String)
'execute the search for users
ExecuteSql "Select * From tblAccountProfile where id Like '" & id & "%'"
With rs
'the system has found something
If .EOF = False Then
    lblId.Caption = .Fields!id
    txtFname.Text = .Fields!fname
    txtMname.Text = .Fields!mname
    txtLname.Text = .Fields!lname
    txtAge.Text = .Fields!age
    cboGender.Text = .Fields!gender
    txtAddress.Text = .Fields!address
    txtBdate.Text = Format(.Fields!bdate, "mm/dd/yyyy")
    cboStatus.Text = .Fields!Status
    txtContact.Text = .Fields!contact
    txtDhired.Text = Format(.Fields!date_hired, "mm/dd/yyyy")
    ComboEmptyStatus.Text = .Fields!emp_status
    cboPosition.Text = .Fields!Position
    cboRemarks.Text = .Fields!remarks
    If .Fields!image_name <> Empty Then
        imgProfile.Picture = LoadPicture(App.Path & "\Images\accounts\" & .Fields!image_name)
    Else
        imgProfile.Picture = LoadPicture()
    End If
    ImgName = .Fields!image_name
    cmdSave.Caption = "&Update"
Else
    ClrFlds
End If
End With
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
