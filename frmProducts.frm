VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmProducts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Products"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Product information"
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   6495
      Begin VB.TextBox txtField 
         DataField       =   "Unit"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "dcProducts"
         Height          =   300
         Index           =   6
         Left            =   3840
         TabIndex        =   20
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         DataField       =   "ProductID"
         DataSource      =   "dcProducts"
         Height          =   300
         Index           =   0
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         DataField       =   "QuantityPerUnit"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "dcProducts"
         Height          =   300
         Index           =   5
         Left            =   1560
         TabIndex        =   7
         Top             =   3600
         Width           =   1575
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1755
         Width           =   1815
      End
      Begin VB.TextBox txtField 
         DataField       =   "UnitPrice"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dcProducts"
         Height          =   300
         Index           =   4
         Left            =   1560
         TabIndex        =   6
         Top             =   3150
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         DataField       =   "SerialNumber"
         DataSource      =   "dcProducts"
         Height          =   300
         Index           =   3
         Left            =   1560
         TabIndex        =   4
         Top             =   2220
         Width           =   1815
      End
      Begin VB.TextBox txtField 
         DataField       =   "ProductDescription"
         DataSource      =   "dcProducts"
         Height          =   300
         Index           =   2
         Left            =   1560
         MaxLength       =   255
         TabIndex        =   2
         Top             =   1320
         Width           =   4575
      End
      Begin VB.TextBox txtField 
         DataField       =   "ProductName"
         DataSource      =   "dcProducts"
         Height          =   300
         Index           =   1
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         DataField       =   "Discontinued"
         DataSource      =   "dcProducts"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   2715
         Width           =   375
      End
      Begin VB.TextBox txtCategory 
         DataField       =   "CategoryID"
         DataSource      =   "dcProducts"
         Height          =   195
         Left            =   3000
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1800
         Width           =   150
      End
      Begin VB.Label Label7 
         Caption         =   "Unit"
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   3650
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Product "
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Qty per unit"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3650
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Desc"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1305
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Serial number"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2220
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Unit price"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3150
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Category"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1770
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Discontinued"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2685
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc dcProducts 
      Height          =   330
      Left            =   120
      Top             =   5040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Driver=SQLite3 ODBC Driver; Database=Orders.db"
      OLEDBString     =   "Driver=SQLite3 ODBC Driver; Database=Orders.db"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Products"
      Caption         =   "Products"
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
            Picture         =   "frmProducts.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProducts.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProducts.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProducts.frx":09F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProducts.frx":0D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProducts.frx":109A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6720
      _ExtentX        =   11853
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
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private NewMode As Boolean
Private EditMode As Boolean
Private CancellingMode As Boolean
Public CurrentProductID As String

'SKS Demo TODO: Go the the designer and change the data binding of _txtField_4 like this:
'_txtField_4.DataBindings.Add("Text", dcProducts, "UnitPrice", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged, null, "C2");


Private Sub cmbCategory_Click()
If cmbCategory.ListCount = 0 Or cmbCategory.ListIndex = -1 Then
    Exit Sub
End If
txtCategory = cmbCategory.ItemData(cmbCategory.ListIndex)
End Sub


Private Sub Form_Unload(Cancel As Integer)
CurrentProductID = dcProducts.Recordset.Fields("ProductId")
End Sub

Private Sub txtCategory_Change()
If cmbCategory.ListCount = 0 Then
    LoadCombo "Categories", cmbCategory, "CategoryName", "CategoryID"
End If
If txtCategory = Empty Then
    cmbCategory.ListIndex = -1
    Exit Sub
End If
Dim Index As Integer
Index = -1
For i = 0 To cmbCategory.ListCount
    If cmbCategory.ItemData(i) = txtCategory Then
        Index = i
        Exit For
    End If
Next
cmbCategory.ListIndex = i
End Sub

Private Sub Form_Load()
txtCategory.Height = 0
txtCategory.Width = 0
dcProducts.ConnectionString = ConnectionString
NewMode = False
EditMode = False
CancellingMode = False
If cmbCategory.ListCount = 0 Then
    LoadCombo "Categories", cmbCategory, "CategoryName", "CategoryID"
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim x As Variant
'SKS Demo TODO: dcProducts.SetFocus()
Select Case Button.Caption
Case "Add"
'Add new record
NewMode = True
dcProducts.Recordset.AddNew
dcProducts.Recordset("UnitsInStock") = 0
dcProducts.Recordset("UnitsOnOrder") = 0
dcProducts.Recordset("Discontinued") = 0
Case "Edit"
'Edit mode
EditMode = True
'dcProducts.Recordset.EditMode =
Case "Save"
'Save data
dcProducts.Recordset.Update
dcProducts.Recordset.Requery ' SQLite ODBC driver needs to requery the info
EditMode = False
NewMode = False
Case "Delete"
'Delete record
If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Delete record") = vbYes Then
    dcProducts.Recordset.Delete
    dcProducts.Recordset.Requery
End If
Case "Search"
'Search for records
SearchShow "Products", "ProductName", "product"
Case "Cancel"
    CancellingMode = True
    'Cancel edited changes
    EditMode = False
    NewMode = False
    dcProducts.Recordset.CancelUpdate
    dcProducts.Recordset.Requery
    CancellingMode = False
    Unload Me
End Select
End Sub

'Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
'If Index = 0 Then
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'ElseIf Index = 4 Or Index = 5 Then
'    Select Case KeyAscii
'        Case vbKey0 To vbKey9
'        Case vbKeyBack, vbKeyClear, vbKeyDelete
'        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
'        Case Else
'            KeyAscii = 0
'            Beep
'    End Select
'End If
'End Sub
