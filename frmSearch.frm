VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   2595
   ClientLeft      =   7290
   ClientTop       =   4830
   ClientWidth     =   5670
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   5775
      TabIndex        =   6
      Top             =   840
      Width           =   5775
   End
   Begin VB.ComboBox cboSrchBy 
      Height          =   315
      ItemData        =   "frmSearch.frx":1681C
      Left            =   3120
      List            =   "frmSearch.frx":1681E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtSrchStr 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search for a specific item"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   735
      TabIndex        =   5
      Top             =   480
      Width           =   1785
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Left            =   735
      TabIndex        =   4
      Top             =   120
      Width           =   750
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   120
      Picture         =   "frmSearch.frx":16820
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Search by:"
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
      Left            =   1875
      TabIndex        =   2
      Top             =   2160
      Width           =   930
   End
   Begin VB.Label lblSrchBy 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Field"
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
      Left            =   105
      TabIndex        =   1
      Top             =   1080
      Width           =   1740
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SearchTable As String
Private Sub cboSrchBy_Click()
    lblSrchBy.Caption = cboSrchBy.text
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Public Sub Search(Table As String, fieldToSearch As String, itemToSearch As String)
If itemToSearch <> Empty Then
    Label20.Caption = "Search for a " & itemToSearch
End If
SearchTable = Table
ExecuteSql "Select * from " & Table & " limit 1;"
For i = 0 To (rs.Fields.Count - 1)
    cboSrchBy.AddItem (rs.Fields(i).Name)
Next i
cboSrchBy = fieldToSearch
End Sub

Private Sub cmdSearch_Click()
If Right(txtSrchStr.text, 1) = "'" Then
    txtSrchStr.text = Empty
End If
Dim txtToSearch As String

If Trim(txtSrchStr.text) <> Empty Then
    txtToSearch = txtSrchStr.text
Else
    txtToSearch = "%"
End If
If SearchTable = "Customers" Then
    SearchCriteriaCustomers lblSrchBy.Caption, txtToSearch
ElseIf SearchTable = "Products" Then
    SearchCriteriaProducts lblSrchBy.Caption, txtToSearch
ElseIf SearchTable = "Providers" Then
    SearchCriteriaProviders lblSrchBy.Caption, txtToSearch
End If
End Sub

'''
Public Sub SearchCriteriaCustomers(field As String, value As String)
ExecuteSql "Select * from Customers where " & field & " LIKE '" & value & "%'"
If rs.RecordCount = 0 Then
    MsgBox "There are no records with the selected criteria", vbInformation, "Search"
Else
    LogStatus "There are " & rs.RecordCount & " that meet with the selected criteria"
    Set frmCustomers.dcCustomers.Recordset = rs
End If
End Sub

Public Sub SearchCriteriaProducts(field As String, value As String)
ExecuteSql "Select * from Products where " & field & " LIKE '" & value & "%'"
If rs.RecordCount = 0 Then
    MsgBox "There are no records with the selected criteria", vbInformation, "Search"
Else
    Set frmProducts.dcProducts.Recordset = rs
End If
End Sub

Public Sub SearchCriteriaProviders(field As String, value As String)
ExecuteSql "Select * from Providers where " & field & " LIKE '" & value & "%'"
If rs.RecordCount = 0 Then
    MsgBox "There are no records with the selected criteria", vbInformation, "Search"
Else
    LogStatus "There are " & rs.RecordCount & " that meet with the selected criteria"
    Set frmProviders.dcProviders.Recordset = rs
End If
End Sub
