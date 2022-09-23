VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmseeds 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdselectotherproduct 
      Caption         =   "Select Other Product"
      Height          =   855
      Left            =   11520
      TabIndex        =   19
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtriceseedqauntity 
      Height          =   495
      Left            =   6120
      TabIndex        =   18
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtcottonseedquantity 
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cndcancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   12240
      TabIndex        =   15
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdbill 
      Caption         =   "BILL"
      Height          =   855
      Left            =   12120
      TabIndex        =   14
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtriceseed 
      Height          =   615
      Left            =   4800
      TabIndex        =   13
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtwheatseed 
      Height          =   615
      Left            =   8400
      TabIndex        =   12
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtjawarseed 
      Height          =   615
      Left            =   4800
      TabIndex        =   11
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox txtgroundnutsseed 
      Height          =   615
      Left            =   8400
      TabIndex        =   10
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox txtsugarcaneseed 
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   6120
      Width           =   2655
   End
   Begin VB.CommandButton cmdgroundnutsseed 
      Height          =   1935
      Left            =   8400
      TabIndex        =   8
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdjawarseed 
      Height          =   1935
      Left            =   4800
      Picture         =   "frmGridbill.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdsugarcaneseed 
      Height          =   1935
      Left            =   960
      Picture         =   "frmGridbill.frx":B3D0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdwheatseed 
      Height          =   1935
      Left            =   8400
      Picture         =   "frmGridbill.frx":44F87
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton cmdriceseed 
      Height          =   1935
      Left            =   4800
      Picture         =   "frmGridbill.frx":49F52
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtcottonseed 
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdcottonseed 
      Height          =   1935
      Left            =   960
      Picture         =   "frmGridbill.frx":4D0F1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   2690
   End
   Begin MSFlexGridLib.MSFlexGrid grdSeed 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   6720
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   5741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Available Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "SEEDS"
      Height          =   735
      Left            =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmseeds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim i As Integer


Private Sub SetGridProperties()
With grdSeed
.Cols = 5
.Rows = 1
.FocusRect = flexFocusHeavy
.SelectionMode = flexSelectionFree


.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 1500
.ColWidth(4) = 1500

.Row = 0
.Col = 0
.Text = "Seed ID"

.Col = 1
.Text = "Seed Name"

.Col = 2
.Text = "Company Name"

.Col = 3
.Text = "Weight"

.Col = 4
.Text = "Price"

End With
End Sub


Private Sub cmdbill_Click()
FrmBill.Show


End Sub


Private Sub cmdcottonseed_Click()
With grdSeed
.AddItem "301" & vbTab & "Cotton Seed" & vbTab & "GreenGold Seeds" & vbTab & "5KG" & vbTab & "1200"

End With

With grdSeed
.Col = 1
.Row = 1
End With
End Sub

Private Sub cmdgroundnutsseed_Click()
With grdSeed
.AddItem "306" & vbTab & "Groundnuts Seed" & vbTab & "Mahabeej" & vbTab & "5KG" & vbTab & "1800"

End With

With grdSeed
.Col = 1
.Row = 1
End With
End Sub

Private Sub cmdjawarseed_Click()
With grdSeed
.AddItem "305" & vbTab & "Jawar Seed" & vbTab & "National Agro Industries" & vbTab & "5KG" & vbTab & "1100"

End With

With grdSeed
.Col = 1
.Row = 1
End With
End Sub

Private Sub cmdriceseed_Click()
With grdSeed
.AddItem "302" & vbTab & "Rice Seed" & vbTab & "Mahabeej" & vbTab & "5KG" & vbTab & "1400"

End With

With grdSeed
.Col = 1
.Row = 1
End With
End Sub

Private Sub cmdselectotherproduct_Click()
frmselectproduct.Show
End Sub

Private Sub cmdsugarcaneseed_Click()
With grdSeed
.AddItem "304" & vbTab & "Sugarcane Seed" & vbTab & "GreenGold Seeds" & vbTab & "5KG" & vbTab & "1600"

End With

With grdSeed
.Col = 1
.Row = 1
End With
End Sub

Private Sub cmdwheatseed_Click()
With grdSeed
.AddItem "303" & vbTab & "Wheat Seed" & vbTab & "Mahyco" & vbTab & "5KG" & vbTab & "1500"

End With

With grdSeed
.Col = 1
.Row = 1
End With
End Sub

Private Sub cndcancel_Click()
con.Close
Unload Me
End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BHAVESH\Desktop\Vb 5 sem project\Vb5sempro.mdb;Persist Security Info=False"
con.Open
rs.Open "select * from AddSeeds", con, adOpenDynamic, adLockOptimistic

'rs.Open "Select SeedNo,SStock from Bill where SeedNo='301'", con, adOpenDynamic, adLockPessimistic
'If rs.BOF Then

txtcottonseed.Text = rs("SStock")
txtriceseed.Text = rs("SStock")
txtwheatseed = rs("SStock")
txtsugarcaneseed = rs("SStock")
txtjawarseed = rs("SStock")
txtgroundnutsseed = rs("SStock")

Call SetGridProperties
End Sub

