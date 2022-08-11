VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReceptionApproval 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Stock to Inventory"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7845
   Begin VB.CommandButton cmdApprove 
      Caption         =   "&Approve"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&Info"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   6600
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid fgOrders 
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      BorderStyle     =   0
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   7065
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13785
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel "
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search supplier"
      Height          =   2295
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7575
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         ItemData        =   "frmOrderAproval.frx":0000
         Left            =   5040
         List            =   "frmOrderAproval.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkTo 
         Caption         =   "To:"
         Height          =   255
         Left            =   5040
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1200
         Width           =   615
      End
      Begin VB.CheckBox chkFrom 
         Caption         =   "From:"
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1200
         Width           =   705
      End
      Begin VB.TextBox txtProductID 
         Height          =   300
         Left            =   5040
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtOrderID 
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtContactLastName 
         Height          =   300
         Left            =   5040
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtContactName 
         Height          =   300
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdProviders 
         Caption         =   "..."
         Height          =   315
         Left            =   3600
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtProviderName 
         Height          =   300
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   300
         Left            =   2040
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   80871425
         CurrentDate     =   41323
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   300
         Left            =   5760
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   80871425
         CurrentDate     =   41323
      End
      Begin VB.Label Label5 
         Caption         =   "Status"
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   290
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Product"
         Height          =   255
         Left            =   3960
         TabIndex        =   21
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Order"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Dates"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Supplier"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmReceptionApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Id As String


Private Sub cmbStatus_Click()
DoSearchReception
End Sub

Private Sub cmdApprove_Click()
LoadActionOrderReception 1
End Sub

Private Sub cmdCancel_Click()
LoadActionOrderReception 2
End Sub

Private Sub cmdInfo_Click()
LoadActionOrderReception
End Sub

Private Sub LoadActionOrderReception(Optional Action As Integer)
If fgOrders.Row > 0 Then
    Dim OrderId As Integer
    With frmActionOrderReception
        OrderId = CInt(fgOrders.TextMatrix(fgOrders.Row, 1))
        .OrderId = OrderId
        .Action = Action
        .LoadData
        .Show 'vbModal
    End With
End If
End Sub

Private Sub dtFrom_Change()
chkFrom.value = 1
DoSearchReception
End Sub

Private Sub dtTo_Change()
chkTo.value = 1
DoSearchReception
End Sub


Private Sub fgOrders_DblClick()
cmdInfo_Click
End Sub

Private Sub Form_Load()
InitGrid
End Sub

Private Sub txtOrderID_Change()
DoSearchReception
End Sub

Private Sub txtProductID_Change()
DoSearchReception
End Sub

Private Sub txtProviderName_Change()
DoSearchReception
End Sub

Private Sub txtContactLastName_Change()
DoSearchReception
End Sub

Private Sub txtContactName_Change()
DoSearchReception
End Sub

Private Sub txtName_Change()
DoSearchReception
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdProviders_Click()
frmProviders.Show 'vbModal
txtProviderName = ""
txtContactLastName = ""
txtContactName = ""
DoSearchReception frmProviders.CurrentProviderID
End Sub

Private Sub DoSearchReception(Optional Id As Integer = -1)
Dim filter As String
filter = ""
If Id <> -1 Then
    filter = "o.ProviderID = " & Id
End If
If txtProviderName <> Empty Then
    AppendAND filter
    filter = "p.ProviderName LIKE '%" & txtProviderName & "%'"
End If
If txtContactName <> Empty Then
    AppendAND filter
    filter = filter & "p.ContactFirstName LIKE '%" & txtContactName & "%'"
End If
If txtContactLastName <> Empty Then
    AppendAND filter
    filter = filter & "p.ContactLastName LIKE '%" & txtContactLastName & "%'"
End If
If txtOrderID <> Empty Then
    AppendAND filter
    filter = filter & "o.OrderID = " & txtOrderID
End If
If txtProductID <> Empty Then
    AppendAND filter
    filter = filter & "d.ProductID LIKE '%" & txtProductID & "%'"
End If
If chkFrom.value = 1 Then
    AppendAND filter
    filter = filter & "o.OrderDate >= #" & Format(dtFrom.value, "mm/dd/yyyy") & "#"
End If
If chkTo.value = 1 Then
    AppendAND filter
    filter = filter & "o.OrderDate <= #" & Format(dtTo.value, "mm/dd/yyyy") & "#"
End If
If cmbStatus.ListIndex <> -1 And cmbStatus.text <> "All" Then
    AppendAND filter
    filter = filter & "o.Status = '" & cmbStatus.text & "'"
End If

Dim where As String
where = " Where o.OrderID = d.OrderID And p.ProviderID = o.ProviderID And u.Username = o.ReceivedBy "
If filter <> Empty Then
    filter = where & " AND " & filter
Else
    filter = where
End If

Dim sql As String

sql = "Select o.OrderDate, o.OrderID, p.ProviderName, p.ContactFirstName + ' ' + p.ContactLastName as ContactName, u.Fullname as [Received by], Sum(d.LineTotal) as Price, o.Status " & _
"From OrderReceptions as o, OrderReceptionDetails as d, Providers as p, Users as u " & _
filter & " Group by o.orderDate, o.OrderID, p.ProviderName, p.ContactFirstName + ' ' + p.ContactLastName, u.Fullname, o.Status "
ExecuteSql sql
LogStatus "There are " & rs.RecordCount & " records with the selected criteria", Me
With fgOrders
    .Rows = rs.RecordCount + 1
    If .Rows = 1 Then .FixedRows = 0 Else .FixedRows = 1
    Dim i As Integer
    Dim j As Integer
    i = 1
    While Not rs.EOF
        For j = 0 To rs.Fields.Count - 1
            If Not rs.Fields(j) Is Nothing Then
                .TextMatrix(i, j) = rs.Fields(j)
            End If
        Next j
        rs.MoveNext
        i = i + 1
    Wend
End With
End Sub

Private Sub InitGrid()
With fgOrders
    .Rows = 0
    .Cols = 7
    .FixedCols = 0
    .AddItem "Date" & vbTab & "Order" & vbTab & "Supplier" & vbTab & "Contact" & vbTab & "Received by" & vbTab & "Price" & vbTab & "Status"
    .Rows = 1
    .FixedRows = 0
    .SelectionMode = flexSelectionByRow
End With
End Sub

Private Sub MakeTextBoxVisible(txtBox As textbox, grid As MSFlexGrid)
With grid
    txtBox.text = .TextMatrix(.Row, .col)
    txtBox.Move .CellLeft + .Left, .CellTop + .Top, .CellWidth, .CellHeight
    txtBox.Visible = True
    DoEvents
    txtBox.SetFocus
End With
End Sub



