VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAcntManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Security"
   ClientHeight    =   6225
   ClientLeft      =   4905
   ClientTop       =   2160
   ClientWidth     =   4590
   Icon            =   "frmAddAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   4590
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lstAccounts 
      Height          =   1695
      Left            =   0
      TabIndex        =   15
      Top             =   3960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UserId"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   4695
      TabIndex        =   9
      Top             =   840
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "User information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4335
      Begin VB.TextBox txtFullname 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cboLevel 
         Height          =   315
         ItemData        =   "frmAddAccount.frx":038A
         Left            =   1800
         List            =   "frmAddAccount.frx":0391
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Full name:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
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
         TabIndex        =   8
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "New password"
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
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User level:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label lblId 
         AutoSize        =   -1  'True
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
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   45
      End
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   ">>"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER SECURITY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   720
      TabIndex        =   19
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set user security and system access"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   720
      TabIndex        =   18
      Top             =   480
      Width           =   2235
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Left            =   120
      MouseIcon       =   "frmAddAccount.frx":039D
      MousePointer    =   99  'Custom
      Picture         =   "frmAddAccount.frx":0C67
      Stretch         =   -1  'True
      ToolTipText     =   "View warnings"
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Click Edit to view"
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
      TabIndex        =   3
      Top             =   5760
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   0
      Picture         =   "frmAddAccount.frx":6879
      Top             =   5640
      Width           =   480
   End
End
Attribute VB_Name = "frmAcntManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
Dim SecId As String
Dim CurrentEditedUser As String
Private Sub cmdClear_Click()
'clear all
txtUsername.Text = Empty
txtUsername.SetFocus
ClearFields
End Sub

Private Sub cmdDelete_Click()
If NoRecords(lstAccounts, "Please add a user") Then Exit Sub
If MsgBox("Are you sure you want to delete this user?", vbExclamation + vbYesNo) = vbYes Then
    ExecuteSql "Select * from Users"
    If rs.RecordCount = 1 Then
        MsgBox "Delete Failed: You cannot delete this user.", vbCritical
        Exit Sub
    End If
    ExecuteSql "Select * from User where Username = '" & UserId & "' and level = 'Administrator'"
    With rs
        If .EOF = False Then
            MsgBox "You cannot delete this account", vbCritical
            Exit Sub
        End If
    End With
    ExecuteSql "delete * from Users where Username = '" & lstAccounts.SelectedItem & "'"
End If
End Sub

Private Sub cmdEdit_Click()
If NoRecords(lstAccounts, "No user found on the list. Please add a user account") Then Exit Sub
ExecuteSql "Select * from Users where Username = '" & lstAccounts.SelectedItem & "'"
SecId = rs.Fields!Username
txtUsername.Text = SecId
With rs
    If .EOF Then
        MsgBox "This user does not exist", vbInformation
        txtUsername.SetFocus
    Else
        txtUsername.Text = .Fields!Username
        CurrentEditedUser = txtUsername.Text
        txtPassword.Text = .Fields!Password
        txtFullname.Text = .Fields!Fullname
        cboLevel.Text = .Fields!Level
        cmdSave.Caption = "&Update"
    End If
End With
End Sub

Private Sub cmdSave_Click()
'if no account were selecte on the list or no user accounts on database
If SecId = Empty Then
    MsgBox "No specified Account found. Please add a new Account or select an Account from the list. Click Options", vbExclamation
    ClearFields
    Exit Sub
End If
'text trappings
If TxtEmp(txtUsername) Then Exit Sub
If TxtEmp(txtPassword) Then Exit Sub
If TxtEmp(txtFullname) Then Exit Sub
If CboEmp(cboLevel) Then Exit Sub

ExecuteSql "Select * from Users where Username = '" & SecId & "'"
With rs
    If cmdSave.Caption <> "&Update" Then
        If cboLevel.Text <> "Administrator" Then
            ExecuteSql2 "Select * from Users where level = 'Administrator' and id <> '" & SecId & "'"
            If rs2.EOF Then
                MsgBox "Update failed: No any Administrator found on accounts.  You are not allowed to change the level of this account", vbCritical
                Exit Sub
            End If
        End If
        If CurrentUserAdmin Then
            If cboLevel.Text <> "Administrator" Then
                MsgBox "You cannot add another level without 'Administrator'", vbInformation
                cboLevel.SetFocus
                Exit Sub
            End If
        End If
        .AddNew
        '.Fields!record_no = Val(RcrdId("Users", , "record_no"))
        msg = "Add security account of user " & SecId
    ElseIf CurrentEditedUser <> txtUsername Then
        ExecuteSql2 "Select * from Users where username = '" & txtUsername & "'"
        If rs2.EOF = False Then
            MsgBox "Username '" & txtUsername.Text & "' is already taken.", vbInformation
            txtUsername.SetFocus
            txtUsername.SelStart = 0
            txtUsername.SelLength = Len(txtUsername)
            Exit Sub
        End If
    Else
        msg = "Record has been successfully updated"
    End If
    .Fields!Username = txtUsername
    .Fields!Password = txtPassword
    .Fields!Level = cboLevel
    .Fields!Fullname = txtFullname
    .Update
End With
MsgBox msg, vbInformation
ClearFields
RecordView

If CurrentUserAdmin Then
    Unload Me
End If
End Sub

Public Sub RecordView()
ExecuteSql "Select * from Users"
lstAccounts.ListItems.Clear
With rs
    While Not .EOF
        Set x = lstAccounts.ListItems.Add(, , .Fields!Username)
        x.SubItems(1) = .Fields!Fullname
        .MoveNext
    Wend
End With
End Sub

Public Sub ClearFields()
txtUsername.Text = Empty
txtPassword.Text = Empty
txtFullname.Text = Empty
cboLevel.Text = cboLevel.List(0)
SecId = Empty
cmdSave.Caption = "&Save"
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOptions_Click()
'hide and show options
    If cmdOptions.Caption = "&Options >>" Then
        Me.Height = 6660
        cmdOptions.Caption = "&Options <<"
    Else
        Me.Height = 4680
        cmdOptions.Caption = "&Options >>"
    End If
End Sub

Private Sub Form_Load()
ExecuteSql "Select * from Levels"
While Not rs.EOF
    cboLevel.AddItem (rs.Fields!Level)
    rs.MoveNext
Wend
If CurrentUserAdmin Then
    cboLevel.Text = "Administrator"
Else
    cboLevel.ListIndex = 0
End If
RecordView
End Sub

Private Sub Form_Unload(Cancel As Integer)
If CurrentUserAdmin Then
    ExecuteSql "Select * from Users"
    If rs.EOF Then
        MsgBox "System has failed to initialized. Please contact your administrator" & vbNewLine _
                & vbNewLine & "Status: analysing accounts configuration" & vbNewLine & _
                "Error: No users found", vbCritical
        End
    End If
    'frmxSplash.tmrLoad.Enabled = True
End If
End Sub

Private Sub lstAccounts_DblClick()
Call cmdEdit_Click
End Sub

Private Sub txtFullname_GotFocus()
SelAll txtFullname
End Sub

Private Sub txtPassword_GotFocus()
SelAll txtPassword
End Sub

Private Sub txtUsername_GotFocus()
SelAll txtUsername
End Sub
