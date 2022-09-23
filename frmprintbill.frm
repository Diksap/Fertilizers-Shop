VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.Ocx"
Begin VB.Form frmprintbill 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "frmprintbill.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtgrandtotal 
      Height          =   495
      Left            =   15360
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   7680
      Width           =   2055
   End
   Begin VB.TextBox txtgst 
      Height          =   495
      Left            =   15360
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   6720
      Width           =   2055
   End
   Begin VB.TextBox txtsubtotal 
      Height          =   495
      Left            =   15360
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5760
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid grdbill 
      Height          =   2535
      Left            =   4440
      TabIndex        =   11
      Top             =   5640
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   10
      Cols            =   4
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   " *  NOTE : Goods Once Sold Can Not Be Returned."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   8760
      Width           =   9255
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
      Left            =   12480
      TabIndex        =   14
      Top             =   7680
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
      Left            =   12480
      TabIndex        =   13
      Top             =   6720
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
      Left            =   12480
      TabIndex        =   12
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   20280
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   20280
      Y1              =   5160
      Y2              =   5160
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
      Left            =   14640
      TabIndex        =   10
      Top             =   4200
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
      Left            =   12480
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
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
      Left            =   14640
      TabIndex        =   8
      Top             =   3240
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
      Left            =   12480
      TabIndex        =   7
      Top             =   3240
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
      Left            =   7560
      TabIndex        =   6
      Top             =   4200
      Width           =   2895
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
      Left            =   4560
      TabIndex        =   5
      Top             =   4200
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
      Left            =   7560
      TabIndex        =   4
      Top             =   3360
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
      Left            =   4560
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   20280
      Y1              =   2760
      Y2              =   2760
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
      Left            =   8640
      TabIndex        =   2
      Top             =   1920
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
      Left            =   10440
      TabIndex        =   1
      Top             =   1080
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
      Left            =   9000
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   5160
      Picture         =   "frmprintbill.frx":1717F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3210
   End
End
Attribute VB_Name = "frmprintbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call SetGridProperties
lblbillno.Caption = frmpracticebill2.lblbillno.Caption
lbldate.Caption = frmpracticebill2.lbldate.Caption
lblfarmername.Caption = frmpracticebill2.lblfarmername.Caption
lblmobilenumber.Caption = frmpracticebill2.lblmobilenumber.Caption
txtsubtotal.Text = frmpracticebill2.txtsubtotal.Text
txtgst.Text = frmpracticebill2.txtgst.Text
txtgrandtotal.Text = frmpracticebill2.txtgrandtotal.Text

Dim i As Integer
Dim j As Integer
grdbill.Rows = frmpracticebill2.grdbill.Rows
grdbill.Cols = frmpracticebill2.grdbill.Cols
For i = 0 To frmpracticebill2.grdbill.Rows - 1
For j = 0 To frmpracticebill2.grdbill.Cols - 1
grdbill.TextMatrix(i, j) = frmpracticebill2.grdbill.TextMatrix(i, j)
Next
Next
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


