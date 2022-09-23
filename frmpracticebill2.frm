VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.Ocx"
Begin VB.Form frmpracticebill2 
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form2"
   Picture         =   "frmpracticebill2.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00C0E0FF&
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
      Height          =   855
      Left            =   17520
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8040
      Width           =   2175
   End
   Begin VB.TextBox txtgrandtotal 
      Height          =   495
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   8040
      Width           =   2055
   End
   Begin VB.TextBox txtgst 
      Height          =   495
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   7320
      Width           =   2055
   End
   Begin VB.TextBox txtsubtotal 
      Height          =   495
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00C0E0FF&
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
      Height          =   855
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6720
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid grdbill 
      Height          =   5415
      Left            =   1680
      TabIndex        =   22
      Top             =   4200
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   10
      Cols            =   4
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8040
      Width           =   2175
   End
   Begin VB.ComboBox ComboFarmerName 
      Height          =   315
      Left            =   13440
      TabIndex        =   19
      Text            =   "Select Farmer name"
      Top             =   4200
      Width           =   3495
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   2175
   End
   Begin VB.TextBox txtqty 
      Height          =   495
      Left            =   13440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5640
      Width           =   2775
   End
   Begin VB.ComboBox ComboProductName 
      Height          =   315
      Left            =   13440
      TabIndex        =   0
      Text            =   "Select Product Name"
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   20280
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Line Line3 
      X1              =   8760
      X2              =   20280
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   20280
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   20280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   4200
      Picture         =   "frmpracticebill2.frx":1717F
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   3210
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Farmer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   20
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9840
      TabIndex        =   18
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Quantity"
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
      Left            =   9840
      TabIndex        =   17
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Label lblmobilenumber 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   16
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No."
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
      Left            =   12120
      TabIndex        =   15
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblfarmername 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label lbldate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   13
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   12
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Farmer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblbillno 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Mob No. 9898767890 / 9034543213"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   1440
      Width           =   7935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Farmer@gmail.com"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   7
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FARMERS FRIEND"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   5
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "GST 5%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   3
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   6600
      Width           =   2055
   End
End
Attribute VB_Name = "frmpracticebill2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Dim sql As String
Dim q As Integer

Private Sub cmdcancel_Click()
con.Close
Unload Me
End Sub

Private Sub cmdPrint_Click()
frmprintbill.Show
End Sub


'To save record in Bill table
Private Sub cmdsave_Click()
Dim r As Integer
Dim row As Integer
For row = 1 To grdbill.Rows - 1
rs.AddNew
rs("BillNo") = lblbillno.Caption
rs("Bill_Date") = lbldate.Caption
rs("FarmerName") = lblfarmername.Caption
rs("FarmerMobNo") = lblmobilenumber.Caption

rs("ProductName") = grdbill.TextMatrix(row, 0)
rs("Price") = grdbill.TextMatrix(row, 1)
rs("Quantity") = grdbill.TextMatrix(row, 2)
rs("Amount") = grdbill.TextMatrix(row, 3)

Next row

rs.AddNew
rs("BillNo") = lblbillno.Caption
rs("Sub_Total") = txtsubtotal.Text
rs("GST") = txtgst.Text
rs("Grand_Total") = txtgrandtotal.Text
rs.Update

MsgBox "Record has been Save Sucessfully", vbInformation
End Sub



'When we select Farmer Name from combofarmername then farmer name and mobile no is display in textbox
Private Sub ComboFarmerName_Click()
Set rs2 = New ADODB.Recordset
rs2.Open "select * from Farmer where F_Name='" & ComboFarmerName.Text & "'", con, adOpenDynamic, adLockOptimistic
If Not rs2.EOF Then
lblfarmername.Caption = rs2!F_Name
End If
rs2.Close

rs2.Open "Select F_MobileNo from Farmer where F_Name='" & ComboFarmerName.Text & "'", con, adOpenDynamic, adLockOptimistic
If Not rs2.EOF Then
lblmobilenumber.Caption = rs2!F_MobileNo
End If
rs2.Close
End Sub

Private Sub cmdadd_Click()
Set rs1 = New ADODB.Recordset
'when we select product and click on add button then the selected product price will be display in MS-Flex
rs1.Open "Select Product_Price from Product where Product_Name='" & ComboProductName.Text & "'", con, adOpenDynamic, adLockOptimistic
Dim a As Integer
Dim p As Integer
a = rs1!Product_Price * Val(txtqty.Text)  'calculate Quantity wise product price

'to add product price and quantity in grdbill
With grdbill
.AddItem ComboProductName.Text & vbTab & rs1!Product_Price & vbTab & txtqty.Text & vbTab & a
End With

'for amount total
Dim t As Integer
Dim s As Integer
Dim temp As Integer
temp = 0
Set rs = New ADODB.Recordset
rs.Open "select * from Bill where BillNo= " & lblbillno.Caption & "", con, adOpenStatic, adLockOptimistic
For i = 1 To grdbill.Rows - 1
temp = temp + Val(grdbill.TextMatrix(i, 3)) 'From Grid we collect amount column of all products
Next i
txtsubtotal.Text = temp
txtgst.Text = Math.Round(Val(txtsubtotal.Text) * 0.05)
txtgrandtotal.Text = Val(txtsubtotal.Text) + Val(txtgst.Text)


Call qty ' for minus qyantity from stock
txtqty.Text = ""
End Sub
'to deduct Quantity from product table
Public Sub qty()
Dim r As Integer
Set rs1 = New ADODB.Recordset
rs1.Open "select * from Product", con, adOpenStatic, adLockOptimistic
sql = "UPDATE Product set Stock = Stock - " & Val(txtqty.Text) & " WHERE Product_Name = '" & ComboProductName.Text & "'"
con.Execute (sql)
rs1.Update
End Sub


Private Sub Form_Load()
Me.fillcomboproduct
Me.fillcomboFarmer
Call SetGridProperties
rs.Open "select * from Bill", con, adOpenStatic, adLockOptimistic
rs.MoveLast
'Auto generate BillNo and Date
lblbillno.Caption = Val(rs(0)) + 1
lbldate.Caption = Now()
rs.Close
End Sub
'To display Product Name in comboproduct combobox
Sub fillcomboproduct()
Module1.main
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from Product", con, adOpenDynamic, adLockOptimistic
While rs1.EOF = False
ComboProductName.AddItem rs1!Product_Name
rs1.MoveNext
Wend
rs1.Close
End Sub
'To display Farmer Name in combofarmer combobox
Sub fillcomboFarmer()
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from Farmer", con, adOpenDynamic, adLockOptimistic
While rs2.EOF = False
ComboFarmerName.AddItem rs2!F_Name
rs2.MoveNext
Wend
rs2.Close
Set rs2 = Nothing
End Sub

Private Sub SetGridProperties()
With grdbill
.Cols = 4
.Rows = 1
.FocusRect = flexFocusHeavy
.SelectionMode = flexSelectionFree


.ColWidth(0) = 2500
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 1500

.row = 0
.Col = 0
.Text = "Product Name"

.Col = 1
.Text = "Product Price"

.Col = 2
.Text = "Quantity"

.Col = 3
.Text = "Amount"

End With
End Sub





