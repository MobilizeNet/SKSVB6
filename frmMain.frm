VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Sales Agent"
   ClientHeight    =   10260
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   16695
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9885
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23777
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "1:25 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2/21/2018"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCustomer 
         Caption         =   "&Manage Customers"
      End
      Begin VB.Menu mnuProviders 
         Caption         =   "Manage Su&ppliers "
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOrders 
      Caption         =   "&Orders"
      Begin VB.Menu mnuCreateOrderRequest 
         Caption         =   "Create Order"
      End
      Begin VB.Menu mnuOrderRequestsApproval 
         Caption         =   "Create Invoice"
      End
      Begin VB.Menu lExit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateOrderReception 
         Caption         =   "Add Stock Order"
      End
      Begin VB.Menu mnuOrderReceptionsApproval 
         Caption         =   "Add Stock to Inventory"
      End
   End
   Begin VB.Menu mnuMainInventory 
      Caption         =   "&Inventory"
      Begin VB.Menu mnuAddStockManually 
         Caption         =   "Inventory Update"
      End
      Begin VB.Menu mnuAdjustStockManually 
         Caption         =   "Inventory Adjust"
      End
   End
   Begin VB.Menu mnuAccounts 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuProducts 
         Caption         =   "Manage Products"
      End
      Begin VB.Menu mnuSecurity 
         Caption         =   "Manage Users"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    frmSplash.Show vbModal
    frmOrderRequest.Show
    
    
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuAddStockManually_Click()
frmAddStockManual.Show
End Sub

Private Sub mnuAdjustStockManually_Click()
frmAdjustStockManual.Show
End Sub

Private Sub mnuCreateOrderReception_Click()
frmOrderReception.Show
End Sub

Private Sub mnuCreateOrderRequest_Click()
frmOrderRequest.Show
End Sub

Private Sub mnuCustomer_Click()
frmCustomers.Show vbModal
frmCustomers.InitForm
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuOrderReceptionsApproval_Click()
frmReceptionApproval.Show
End Sub

Private Sub mnuOrderRequestsApproval_Click()
frmRequestApproval.Show
End Sub

Private Sub mnuProducts_Click()
frmProducts.Show vbModal
End Sub

Private Sub mnuProviders_Click()
frmProviders.Show vbModal
End Sub

Private Sub mnuSecurity_Click()
frmUsersManage.Show
End Sub


