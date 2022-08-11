VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCustomers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   6810
   Begin MSAdodcLib.Adodc dcCustomers 
      Height          =   330
      Left            =   120
      Top             =   6000
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   2
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   1
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Driver=SQLite3 ODBC Driver; Database=Orders.db;"
      OLEDBString     =   "Driver=SQLite3 ODBC Driver; Database=Orders.db;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Customers;"
      Caption         =   "Customers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer information"
      Height          =   5175
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   6495
      Begin VB.TextBox txtField 
         DataField       =   "PostalCode"
         DataSource      =   "dcCustomers"
         Height          =   300
         Index           =   4
         Left            =   1560
         TabIndex        =   7
         Top             =   2208
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         DataField       =   "CompanyName"
         DataSource      =   "dcCustomers"
         Height          =   300
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         DataField       =   "StateOrProvince"
         DataSource      =   "dcCustomers"
         Height          =   300
         Index           =   6
         Left            =   1560
         TabIndex        =   9
         Top             =   3132
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         DataField       =   "Country/Region"
         DataSource      =   "dcCustomers"
         Height          =   300
         Index           =   7
         Left            =   1560
         TabIndex        =   10
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         DataField       =   "EmailAddress"
         DataSource      =   "dcCustomers"
         Height          =   300
         Index           =   2
         Left            =   1560
         TabIndex        =   2
         Top             =   1284
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         DataField       =   "CompanyOrDepartment"
         DataSource      =   "dcCustomers"
         Height          =   300
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Top             =   822
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         DataField       =   "BillingAddress"
         DataSource      =   "dcCustomers"
         Height          =   300
         Index           =   3
         Left            =   1560
         TabIndex        =   6
         Top             =   1746
         Width           =   4695
      End
      Begin VB.TextBox txtField 
         DataField       =   "City"
         DataSource      =   "dcCustomers"
         Height          =   300
         Index           =   5
         Left            =   1560
         TabIndex        =   8
         Top             =   2670
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contact"
         Height          =   1455
         Left            =   3360
         TabIndex        =   28
         Top             =   200
         Width           =   2900
         Begin VB.TextBox txtField 
            DataField       =   "ContactTitle"
            DataSource      =   "dcCustomers"
            Height          =   300
            Index           =   11
            Left            =   960
            TabIndex        =   3
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtField 
            DataField       =   "ContactLastName"
            DataSource      =   "dcCustomers"
            Height          =   300
            Index           =   13
            Left            =   960
            TabIndex        =   5
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtField 
            DataField       =   "ContactFirstName"
            DataSource      =   "dcCustomers"
            Height          =   300
            Index           =   12
            Left            =   960
            TabIndex        =   4
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Title"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Last name"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1035
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "First name"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   637
            Width           =   855
         End
      End
      Begin VB.TextBox txtField 
         DataField       =   "PhoneNumber"
         DataSource      =   "dcCustomers"
         Height          =   300
         Index           =   8
         Left            =   4680
         TabIndex        =   11
         Top             =   2685
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         DataField       =   "FaxNumber"
         DataSource      =   "dcCustomers"
         Height          =   300
         Index           =   10
         Left            =   4680
         TabIndex        =   13
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         DataField       =   "Notes"
         DataSource      =   "dcCustomers"
         Height          =   780
         Index           =   14
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   4200
         Width           =   6015
      End
      Begin VB.TextBox txtField 
         DataField       =   "Extension"
         DataSource      =   "dcCustomers"
         Height          =   300
         Index           =   9
         Left            =   4680
         TabIndex        =   12
         Top             =   3135
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Zip"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   2208
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Country"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "State/Prov"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3132
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Notes:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Email"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1284
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Fax"
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Ext"
         Height          =   255
         Left            =   3480
         TabIndex        =   22
         Top             =   3135
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Phone"
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Top             =   2685
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "City"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2670
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1746
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Department"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   822
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Company"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomers.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomers.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomers.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomers.frx":09F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomers.frx":0D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomers.frx":109A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   1164
      ButtonWidth     =   1138
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Object.ToolTipText     =   "Create a new record"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Edit this record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Object.ToolTipText     =   "Save the current changes"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Object.ToolTipText     =   "Delete the current record"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Object.ToolTipText     =   "Search for a record"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            Object.ToolTipText     =   "Cancel edited changes"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NewMode As Boolean
Private EditMode As Boolean
Private CancellingMode As Boolean
Public CurrentCustomerID As String



Private Sub Form_Unload(Cancel As Integer)
    CurrentCustomerID = dcCustomers.Recordset.Fields("CustomerId")
End Sub


'Private Sub dcCustomers_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'NewMode = False
'EditMode = False
'CancellingMode = False
'End Sub

'Private Sub dcCustomers_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'CancellingMode = True
'End Sub

Private Sub Form_Load()
    InitForm
End Sub

Public Sub InitForm()
dcCustomers.ConnectionString = ConnectionString
NewMode = False
EditMode = False
CancellingMode = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim x As Variant
Dim newCustomerId As Integer
    Select Case Button.Caption
    Case "Add"
        'Add new record
        NewMode = True
        dcCustomers.Recordset.AddNew
    Case "Edit"
        'Edit mode
        EditMode = True
    Case "Save"
        'Save data
        dcCustomers.Recordset.Update
        dcCustomers.Recordset.Requery ' SQLite ODBC driver needs to requery the info
        EditMode = False
        NewMode = False
    Case "Delete"
        'Delete record
        If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Delete record") = vbYes Then
            dcCustomers.Recordset.Delete
            dcCustomers.Recordset.Requery
        End If
    Case "Search"
        'Search for records
        SearchShow "Customers", "CompanyName", "customer"
    Case "Cancel"
        CancellingMode = True
        'Cancel edited changes
        EditMode = False
        NewMode = False
        dcCustomers.Recordset.CancelUpdate
        dcCustomers.Recordset.Requery
        CancellingMode = False
        Unload Me
    End Select
End Sub

Private Sub txtField_Change(Index As Integer)
If Not CancellingMode Then
    EditMode = True
End If
End Sub

'Used already in frmSearch
Public Sub SearchCriteriaProducts(field As String, value As String)
ExecuteSql "Select * from Customers where " & field & " LIKE '" & value & "%'"
If rs.RecordCount = 0 Then
    MsgBox "There are no records with the selected criteria", vbInformation, "Search"
Else
    LogStatus "There are " & rs.RecordCount & " that meet with the selected criteria"
    Set dcCustomers.Recordset = rs
End If
End Sub
