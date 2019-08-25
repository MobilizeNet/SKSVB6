VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsersManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users Management"
   ClientHeight    =   6210
   ClientLeft      =   4905
   ClientTop       =   2160
   ClientWidth     =   5175
   Icon            =   "frmUsersManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   5175
   Begin MSComctlLib.ListView lstAccounts 
      Height          =   1695
      Left            =   0
      TabIndex        =   7
      Top             =   3960
      Width           =   5055
      _ExtentX        =   8916
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UserId"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Level"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&New"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   4695
      TabIndex        =   15
      Top             =   840
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "User information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   4935
      Begin VB.TextBox txtFullname 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
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
         MaxLength       =   50
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   0
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox cboLevel 
         Height          =   315
         ItemData        =   "frmUsersManage.frx":5C12
         Left            =   1800
         List            =   "frmUsersManage.frx":5C14
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Full name: *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   16
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username: *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   14
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "New password: *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   13
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User level: *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   12
         Top             =   1800
         Width           =   1065
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
         TabIndex        =   11
         Top             =   360
         Width           =   45
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "* Required fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   120
      TabIndex        =   19
      Top             =   5880
      Width           =   1425
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   18
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set user information and access level"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   720
      TabIndex        =   17
      Top             =   480
      Width           =   2640
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      MouseIcon       =   "frmUsersManage.frx":5C16
      MousePointer    =   99  'Custom
      Picture         =   "frmUsersManage.frx":64E0
      Stretch         =   -1  'True
      ToolTipText     =   "View warnings"
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmUsersManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentEditedUser As String

Private Sub cmdClear_Click()
txtUsername.text = Empty
txtUsername.SetFocus
ClearFields
End Sub

Private Sub cmdDelete_Click()
If NoRecords(lstAccounts, "Please add/select a user") Then Exit Sub
If MsgBox("Are you sure you want to delete the user '" & lstAccounts.SelectedItem & "'?", vbExclamation + vbYesNo) = vbYes Then
    ExecuteSql "Select * from Users"
    If rs.RecordCount = 1 Then
        MsgBox "You cannot delete the last user", vbCritical, "Delete error"
        Exit Sub
    End If
    ExecuteSql "delete * from Users where Username = '" & lstAccounts.SelectedItem & "'"
    LoadUsers
End If
End Sub

Private Sub cmdEdit_Click()
If NoRecords(lstAccounts, "Please add/select a user") Then Exit Sub
ExecuteSql "Select * from Users where Username = '" & lstAccounts.SelectedItem & "'"
txtUsername.text = rs.Fields!UserName
If rs.EOF Then
    MsgBox "This user does not exist", vbInformation
    txtUsername.SetFocus
Else
    txtUsername.text = rs.Fields!UserName
    CurrentEditedUser = txtUsername.text
    txtPassword.text = rs.Fields!Password
    txtFullname.text = rs.Fields!Fullname
    cboLevel.text = rs.Fields!Level
    cmdSave.Caption = "&Update"
End If
End Sub

Private Sub cmdSave_Click()
Dim SecId As String
If TextBoxEmpty(txtUsername) Then Exit Sub
If TextBoxEmpty(txtPassword) Then Exit Sub
If TextBoxEmpty(txtFullname) Then Exit Sub
If ComboEmpty(cboLevel) Then Exit Sub

ExecuteSql "Select * from Users where Username = '" & txtUsername & "'"
If cmdSave.Caption <> "&Update" Then
    If Not rs.EOF Then
            MsgBox "Add failed: Username already exists", vbCritical
            Exit Sub
    End If
    
    If cboLevel.text <> "Administrator" Then
        ExecuteSql2 "Select * from Users where level = 'Administrator'"
        If rs2.EOF Then
            MsgBox "Update failed: No any Administrator found on accounts.  You are not allowed to change the level of this account", vbCritical
            Exit Sub
        End If
    End If
    If Not CurrentUserAdmin And cboLevel.text = "Administrator" Then
        MsgBox "You cannot add another level without being 'Administrator'", vbInformation
        cboLevel.SetFocus
        Exit Sub
    End If
    rs.AddNew
    msg = "Added new user " & txtUsername
ElseIf CurrentEditedUser <> txtUsername Then
    ExecuteSql2 "Select * from Users where username = '" & txtUsername & "'"
    If Not rs2.EOF Then
        MsgBox "Username '" & txtUsername & "' already exists.", vbInformation
        txtUsername.SetFocus
        SelectAll txtUsername
        Exit Sub
    End If
    msg = "Record for the user " & txtUsername & " has been successfully updated"
Else
    msg = "Record for the user " & txtUsername & " has been successfully updated"
End If
rs.Fields!UserName = txtUsername
rs.Fields!Password = txtPassword
rs.Fields!Level = cboLevel
rs.Fields!Fullname = txtFullname
rs.Update
LogStatus msg
ClearFields
LoadUsers

If CurrentUserAdmin Then
    Unload Me
End If
End Sub

Public Sub LoadUsers()
ExecuteSql "Select * from Users"
lstAccounts.ListItems.Clear
Dim x As ListItem
While Not rs.EOF
    Set x = lstAccounts.ListItems.Add(, , rs.Fields!UserName)
    x.SubItems(1) = rs.Fields!Fullname
    x.SubItems(2) = rs.Fields!Level
    rs.MoveNext
Wend
End Sub

Public Sub LoadUsersAvoidingWith()
ExecuteSql "Select * from Users"
lstAccounts.ListItems.Clear
Dim x As ListItem
While Not rs.EOF
    Set x = lstAccounts.ListItems.Add(, , rs.Fields!UserName)
    x.SubItems(1) = rs.Fields!Fullname
    x.SubItems(2) = rs.Fields!Level
    rs.MoveNext
Wend
End Sub


Public Sub ClearFields()
txtUsername.text = Empty
txtPassword.text = Empty
txtFullname.text = Empty
cboLevel.ListIndex = -1
cmdSave.Caption = "&Save"
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
ExecuteSql "Select * from Levels"
While Not rs.EOF
    cboLevel.AddItem (rs.Fields!Level)
    rs.MoveNext
Wend
If CurrentUserAdmin Then
    cboLevel.text = "Administrator"
Else
    cboLevel.ListIndex = -1
End If
LoadUsers
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
LogStatus ""
End Sub

Private Sub lstAccounts_DblClick()
Call cmdEdit_Click
End Sub

Private Sub txtFullname_GotFocus()
SelectAll txtFullname
End Sub

Private Sub txtPassword_GotFocus()
SelectAll txtPassword
End Sub

Private Sub txtUsername_GotFocus()
SelectAll txtUsername
End Sub
