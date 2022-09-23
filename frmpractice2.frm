VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmProductdetails 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   7560
      TabIndex        =   2
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton cmdBill 
      Caption         =   "Bill"
      Height          =   855
      Left            =   10320
      TabIndex        =   1
      Top             =   7440
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   5655
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   9975
      _Version        =   393216
   End
End
Attribute VB_Name = "frmProductdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetGridProperties()
With grd
.Cols = 6
.Rows = 1
.FocusRect = flexFocusHeavy
.SelectionMode = flexSelectionFree


.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 1500
.ColWidth(4) = 1500

.Row = 0
.col = 0
.Text = "Product ID"

.col = 1
.Text = "Product Name"

.col = 2
.Text = "Product Company Name"

.col = 3
.Text = "Product Weight"

.col = 4
.Text = "Product Price"

.col = 5
.Text = "Product Quantity"

End With

End Sub

Private Sub cmdbill_Click()
FrmBill.Show
End Sub

Private Sub cmdcancel_Click()
con.Close
Unload Me
End Sub

Private Sub Form_Load()
Call SetGridProperties

con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BHAVESH\Desktop\Vb 5 sem project\Vb5sempro.mdb;Persist Security Info=False"
con.Open
'rs.Open "select * from Bill", con, adOpenStatic, adLockOptimistic
End Sub
