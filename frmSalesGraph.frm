VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmSalesGraph 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame GraphFrame 
      Caption         =   "Sales"
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6975
      Begin MSChart20Lib.MSChart Chart 
         Height          =   5055
         Left            =   120
         OleObjectBlob   =   "frmSalesGraph.frx":0000
         TabIndex        =   6
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Filter 
      Caption         =   "Date range: "
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.OptionButton Option1 
         Caption         =   "Previous month"
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "This month"
         Height          =   255
         Index           =   0
         Left            =   5400
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkFrom 
         Caption         =   "From:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   705
      End
      Begin VB.CheckBox chkTo 
         Caption         =   "To:"
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   100728833
         CurrentDate     =   41323
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   300
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   100728833
         CurrentDate     =   41323
      End
   End
End
Attribute VB_Name = "frmSalesGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dtFrom_Change()
chkFrom.value = 1
End Sub

Private Sub dtTo_Change()
chkTo.value = 1
End Sub

Private Sub DoGraph()
Dim length As Integer
Dim i As Integer

ExecuteSql "Select Top 6 ProductName from Products "
length = 6
Dim arrValues(1 To 6, 1 To 4)
If rs.RecordCount < 6 Then
    length = rs.RecordCount
End If

'ReDim Preserve arrValues(1 To length, 1 To 4)
For i = 1 To length
   arrValues(i, 1) = rs("ProductName") '"Product " & i ' Labels
   arrValues(i, 2) = 2 + i ' Series 1 values.
   arrValues(i, 3) = 2 * i ' Series 2 values.
   arrValues(i, 4) = Abs(10 * Sin(i)) ' Series 3 values.
   rs.MoveNext
Next i
If length < 6 Then
    For i = length + 1 To 6
        arrValues(i, 1) = "Product " & i ' Labels
        arrValues(i, 2) = 2 + i ' Series 1 values.
        arrValues(i, 3) = 2 * i ' Series 2 values.
        arrValues(i, 4) = Abs(10 * Sin(i)) ' Series 3 values.
    Next i
End If
Chart.ChartData = arrValues
End Sub

Private Sub Form_Load()
DoGraph
End Sub

Private Sub Form_Resize()
Filter.Width = ScaleWidth - (Filter.Left + 100)
GraphFrame.Width = ScaleWidth - (GraphFrame.Left + 100)
GraphFrame.Height = ScaleHeight - (GraphFrame.Top + 100)
Chart.Width = GraphFrame.Width - (Chart.Left + 100)
Chart.Height = GraphFrame.Height - (Chart.Top + 100)
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0:
    dtFrom.value = DateSerial(Year(Now), Month(Now), 1)
    dtTo.value = Now
Case 1:
    dtFrom.value = DateSerial(Year(Now), Month(Now) - 1, 1)
    dtTo.value = DateSerial(Year(Now), Month(Now), 0)
End Select
chkTo.value = 1
chkFrom.value = 1
DoGraph
End Sub
