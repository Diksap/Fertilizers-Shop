VERSION 5.00
Begin VB.Form frmPurchase 
   Caption         =   "Form2"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16470
   LinkTopic       =   "Form2"
   Picture         =   "frmPurchase.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   16470
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combodealername 
      Height          =   315
      Left            =   3000
      TabIndex        =   18
      Text            =   "Select Dealer Name"
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox txtdealerid 
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3000
      Width           =   3255
   End
   Begin VB.CommandButton cmdbill 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Purchase Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdtotal 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   7800
      Width           =   3255
   End
   Begin VB.TextBox txtquantity 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   6960
      Width           =   3255
   End
   Begin VB.TextBox txtprice 
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6120
      Width           =   3255
   End
   Begin VB.ComboBox Comboproduct 
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Text            =   "Select Product Name"
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox txtproductid 
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Details"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   5160
      TabIndex        =   0
      Top             =   1200
      Width           =   6615
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
'display dealer name in combobox
Private Sub Combodealername_Click()
Set rs2 = New ADODB.Recordset
rs2.Open "select * from Dealer where D_Name='" & Combodealername.Text & "'", con, 3, 2

If Not rs2.EOF Then
txtdealerid.Text = rs2!D_ID
End If
End Sub
'display product name in combobox
Private Sub comboProduct_Click()
Set rs = New ADODB.Recordset
rs.Open "select * from Product where Product_Name='" & Comboproduct.Text & "'", con, 3, 2

If Not rs.EOF Then
Me.txtproductid.Text = rs!Product_No
Me.txtprice.Text = rs!Product_Price
End If
rs.Close
Set rs = Nothing
End Sub
'To display Product Name in comboproduct combobox
Sub fillcomboproduct()
rs.Open " select * from Product", con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
Comboproduct.AddItem rs!Product_Name
rs.MoveNext
Wend
rs.Close
Set rs = Nothing
End Sub
'To display Dealer Name in combodealer combobox
Sub fillcombodealer()
rs2.Open " select * from Dealer", con, adOpenDynamic, adLockOptimistic
While rs2.EOF = False
Combodealername.AddItem rs2!D_Name
rs2.MoveNext
Wend
rs2.Close
Set rs2 = Nothing
End Sub

Private Sub cmdbill_Click()
PurchaseDataReport1.Show
End Sub

Private Sub cmdcancel_Click()
con.Close
Unload Me
End Sub
'add record to Purchase Table
Private Sub cmdadd_Click()
Set rs1 = New ADODB.Recordset
rs1.Open "select * from Purchase", con, 3, 2
rs1.AddNew
rs1("Dealer_ID") = txtdealerid.Text
rs1("Dealer_Name") = Combodealername.Text
rs1("Product_ID") = txtproductid.Text
rs1("Product_Name") = Comboproduct.Text
rs1("Price") = txtprice.Text
rs1("Quantity") = txtquantity.Text
rs1("Total") = txttotal.Text
rs1.Update
MsgBox "Order is Placed...."
End Sub
'calculate total
Private Sub cmdtotal_Click()
txttotal.Text = Val(txtprice.Text) * Val(txtquantity.Text)
End Sub
Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BHAVESH\Desktop\Vb 5 sem project\Vb5sempro.mdb;Persist Security Info=False"
con.Open
Me.fillcombodealer
Me.fillcomboproduct
End Sub
'validation for price
Private Sub txtprice_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc("A") And KeyAscii < Asc("Z") Or KeyAscii > Asc("a") And KeyAscii < Asc("z") Then
KeyAscii = 0
End If
End Sub
'validation for quantity
Private Sub txtquantity_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc("A") And KeyAscii < Asc("Z") Or KeyAscii > Asc("a") And KeyAscii < Asc("z") Then
KeyAscii = 0
End If
End Sub
'validation for total
Private Sub txttotal_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc("A") And KeyAscii < Asc("Z") Or KeyAscii > Asc("a") And KeyAscii < Asc("z") Then
KeyAscii = 0
End If
End Sub
