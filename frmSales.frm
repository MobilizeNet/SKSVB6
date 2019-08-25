VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9045
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   9045
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvwSales 
      Height          =   2895
      Left            =   0
      TabIndex        =   12
      Top             =   1800
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox ctrLine 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   9015
      TabIndex        =   11
      Top             =   1320
      Width           =   9015
   End
   Begin VB.ComboBox cboCashier 
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
      ItemData        =   "frmSales.frx":038A
      Left            =   5760
      List            =   "frmSales.frx":0391
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox cboMonth 
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
      ItemData        =   "frmSales.frx":039F
      Left            =   3240
      List            =   "frmSales.frx":03CA
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cboYear 
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
      ItemData        =   "frmSales.frx":0445
      Left            =   720
      List            =   "frmSales.frx":0447
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
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
      Left            =   4080
      TabIndex        =   10
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cashier:"
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
      TabIndex        =   7
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Month:"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   840
      Width           =   510
   End
   Begin VB.Label lblSellable 
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
      Left            =   6360
      TabIndex        =   4
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Most Sellable Item:"
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
      Left            =   4440
      TabIndex        =   3
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label lblTotalSales 
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
      Left            =   2520
      TabIndex        =   2
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Sales for the Month:"
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
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCashier_Click()
ViewSales "ProductId", "%"
End Sub

Private Sub cboMonth_Click()
ViewSales "ProductId", "%"
End Sub

Private Sub cboYear_Click()
ViewSales "ProductId", "%"
End Sub

Private Sub Form_Load()
For i = Format(Date, "yyyy") To 2000 Step -1
    cboYear.AddItem i
Next i
cboYear.Text = Val(Format(Date, "yyyy"))
cboMonth.ListIndex = Val(Format(Date, "mm"))
'SetListView lvwSales, True, True
LoadCombo "Users", cboCashier, "id"
'frmMain.tbrMenu.Buttons(4).Value = tbrPressed
End Sub

Private Sub Form_Resize()
lvwSales.Width = ScaleWidth - (lvwSales.Left + 100)
lvwSales.Height = ScaleHeight - (lvwSales.Top + 100)
ctrLine.Width = ScaleWidth - ctrLine.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
'frmMain.tbrMenu.Buttons(4).Value = tbrUnpressed
End Sub

Public Sub ViewSales(RcrdFld As String, RcrdStr As String)
If cboCashier.ListIndex = 0 Then
    ExecuteSql "Select record_no, ProductId, description, gross_amount, net_amount, vat, quantity from tblSales where " & RcrdFld & " LIKE '" & RcrdStr & "%' and format(date_sold, 'm') = " & _
            cboMonth.ListIndex & " and format(date_sold,'yyyy') = " & cboYear.Text & " Order by record_no ASC"
Else
    ExecuteSql "Select record_no, ProductId, description, gross_amount, net_amount, vat, quantity from tblSales where " & RcrdFld & " LIKE '" & RcrdStr & "%' and format(date_sold, 'm') = " & _
            cboMonth.ListIndex & " and format(date_sold,'yyyy') = " & cboYear.Text & " and cashier_id = '" & cboCashier.Text & "' Order by record_no ASC"
End If
lvwSales.ListItems.Clear
Dim x As ListItem
While Not rs.EOF
    Set x = lvwSales.ListItems.Add(, , rs.Fields(0))
    For i = 1 To (rs.Fields.Count - 1)
        x.SubItems(i) = rs.Fields(i)
    Next i
    rs.MoveNext
Wend
ExecuteSql "Select format(sum(net_amount),'#0.00') from tblSales where format(date_sold, 'm') = " & cboMonth.ListIndex & " and format(date_sold,'yyyy') = " & cboYear.Text
lblTotalSales.Caption = "P " & Val(Format(rs.Fields(0), "#,##0.00"))

ExecuteSql "Select ProductId from tblSales where quantity = (Select max(quantity) from tblSales) and format(date_sold, 'm') = " & cboMonth.ListIndex & " and format(date_sold,'yyyy') = " & cboYear.Text
lblSellable.Caption = rs.Fields(0)
End Sub


