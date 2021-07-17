VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8625
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16515
   LinkTopic       =   "Form3"
   Picture         =   "form3.frx":0000
   ScaleHeight     =   8625
   ScaleWidth      =   16515
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "    Welcome to the Cake Shop"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1575
      Left            =   6480
      TabIndex        =   0
      Top             =   840
      Width           =   10335
   End
   Begin VB.Menu stockinventory 
      Caption         =   "Stock Inventory"
      Begin VB.Menu addstock 
         Caption         =   "Add Stock"
      End
      Begin VB.Menu updatecake 
         Caption         =   "Update Cake  Inventory"
      End
      Begin VB.Menu stock 
         Caption         =   "Stock Details"
      End
   End
   Begin VB.Menu supplier 
      Caption         =   "Supplier  Details"
      Begin VB.Menu newsupplier 
         Caption         =   "New Supplier Details"
      End
      Begin VB.Menu editsupplier 
         Caption         =   "Edit Supplier Details"
      End
   End
   Begin VB.Menu customer 
      Caption         =   "Customer Order Details"
   End
   Begin VB.Menu purchase 
      Caption         =   "Purchase"
   End
   Begin VB.Menu reports 
      Caption         =   "Reports"
      Begin VB.Menu billreports 
         Caption         =   "Bill  Reports"
      End
      Begin VB.Menu supplierreports 
         Caption         =   "Supplier  Reports"
      End
      Begin VB.Menu sales 
         Caption         =   "Sales"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addcake_Click()
Form9.Show
End Sub

Private Sub addstock_Click()
Form4.Show

End Sub

Private Sub cakereports_Click()
Form9.Show
End Sub

Private Sub editcake_Click()
Form5.Show

End Sub

Private Sub editdetails_Click()
Form8.Show
End Sub

Private Sub newcake_Click()
Form4.Show

End Sub

Private Sub newdealer_Click()
Form7.Show
End Sub




Private Sub billreports_Click()
BillReport.Show
End Sub

Private Sub customer_Click()
Form9.Show
End Sub

Private Sub editstock_Click()
Form5.Show

End Sub

Private Sub editsupplier_Click()
Form8.Show
End Sub


Private Sub newsupplier_Click()
Form7.Show
End Sub

Private Sub purchase_Click()
Form10.Show
End Sub

Private Sub stock_Click()
Form6.Show

End Sub

Private Sub supplierreports_Click()
SupplierReport.Show
End Sub

Private Sub updatecake_Click()
Form5.Show

End Sub
