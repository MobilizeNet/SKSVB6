VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2055
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4335
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1214.162
   ScaleMode       =   0  'User
   ScaleWidth      =   4070.33
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1770
      TabIndex        =   1
      Top             =   375
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2940
      TabIndex        =   5
      Top             =   1500
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   765
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1335
      TabIndex        =   4
      Top             =   1500
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   390
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   585
      TabIndex        =   2
      Top             =   780
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    ExecuteSql "SELECT * FROM Users WHERE username = '" & txtUserName.Text & "' and password = '" & txtPassword.Text & "'"
    If rs.EOF Then
        MsgBox "Invalid 'Username' or 'Password', please try again!", vbExclamation
        txtUserName.SetFocus
        SelectAll txtUserName
        Exit Sub
    End If
    UserFullname = rs.Fields!Fullname
    UserLevel = rs.Fields!Level
    CurrentUserAdmin = (UserLevel = "Administrator")
    Me.MousePointer = 0
    LoginSucceeded = True
    LogStatus "User : " & UserFullname & " logged at " & DateValue(Now) & "," & TimeValue(Now)
    Unload Me
End Sub
