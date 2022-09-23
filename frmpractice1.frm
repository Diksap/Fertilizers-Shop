VERSION 5.00
Begin VB.Form frmselectproduct 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "frmpractice1.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   44
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   43
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   735
      Left            =   9360
      TabIndex        =   42
      Top             =   10080
      Width           =   1815
   End
   Begin VB.CommandButton cmdCarbendazim 
      Height          =   1935
      Left            =   12600
      TabIndex        =   41
      Top             =   7920
      Width           =   2175
   End
   Begin VB.CommandButton cmdNeemleafmanure 
      Height          =   1935
      Left            =   10080
      Picture         =   "frmpractice1.frx":89610
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CommandButton cmdPalash 
      Height          =   1935
      Left            =   7560
      TabIndex        =   39
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CommandButton cmdSuphala 
      Height          =   1935
      Left            =   5040
      TabIndex        =   38
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CommandButton cmdChrysoperla 
      Height          =   1935
      Left            =   2640
      TabIndex        =   37
      Top             =   7920
      Width           =   2175
   End
   Begin VB.CommandButton cmdTricordrama 
      Height          =   1935
      Left            =   120
      TabIndex        =   36
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   -360
      TabIndex        =   28
      Top             =   6600
      Width           =   14895
   End
   Begin VB.CommandButton cmdRockPhosphate 
      Height          =   1935
      Left            =   12480
      Picture         =   "frmpractice1.frx":8B439
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdPSBManure 
      Height          =   1935
      Left            =   9960
      Picture         =   "frmpractice1.frx":8DEC0
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdNitroganFixation 
      Height          =   1935
      Left            =   7440
      Picture         =   "frmpractice1.frx":8FDD7
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdPhospherous 
      Height          =   1935
      Left            =   5040
      Picture         =   "frmpractice1.frx":9146B
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdJeevamrut 
      Height          =   1935
      Left            =   2640
      Picture         =   "frmpractice1.frx":94939
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdBijamrut 
      Height          =   1935
      Left            =   120
      Picture         =   "frmpractice1.frx":964CA
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   14895
   End
   Begin VB.CommandButton cmdgroundnutsseed 
      Height          =   1935
      Left            =   12480
      Picture         =   "frmpractice1.frx":9E5A4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdjawarseed 
      Height          =   1935
      Left            =   9960
      Picture         =   "frmpractice1.frx":A1325
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdsugarcaneseed 
      Height          =   1935
      Left            =   7440
      Picture         =   "frmpractice1.frx":A55DA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdwheatseed 
      Height          =   1935
      Left            =   5040
      Picture         =   "frmpractice1.frx":A8CAA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdriceseed 
      Height          =   1935
      Left            =   2640
      Picture         =   "frmpractice1.frx":AC2C4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdcottonseed 
      Height          =   1935
      Left            =   120
      Picture         =   "frmpractice1.frx":AF71F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2325
   End
   Begin VB.CommandButton cmdaddproduct 
      Caption         =   "Add Product"
      Height          =   855
      Left            =   12000
      TabIndex        =   0
      Top             =   9960
      Width           =   2175
   End
   Begin VB.Label Label21 
      Caption         =   "Carbendazim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12720
      TabIndex        =   35
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label20 
      Caption         =   "Neem Leaf"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10440
      TabIndex        =   34
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label19 
      Caption         =   "Suphala"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   33
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "Palash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   32
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "Chrysoperla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "Tricordrama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   30
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "INSECTICIDES PRODUCTS"
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
      Left            =   4920
      TabIndex        =   29
      Top             =   6840
      Width           =   4335
   End
   Begin VB.Label Label14 
      Caption         =   "Rock-Phosphate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   21
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "PSB Manure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   20
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "Nitrogan Fixation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   19
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Phospherous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Jeevamrut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Bijamrut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "FERTILIZERS PRODUCTS"
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
      Left            =   4680
      TabIndex        =   15
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label Label7 
      Caption         =   "Groundnuts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   13
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Jowar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   12
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Sugarcane "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   11
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Wheat "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Rice "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Cotton "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "SEEDS PRODUCTS"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmselectproduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim i As Integer

Private Sub cmdaddproduct_Click()
frmProductdetails.Show
End Sub

Private Sub cmdBijamrut_Click()
With frmProductdetails.grd
.AddItem "101" & vbTab & "Bijamrut" & vbTab & "National Fertilizer Ltd" & vbTab & "5KG" & vbTab & "1200"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
ind = btn
frmProductdetails.grd.TextMatrix(2, 5) = Text2.Text
End Sub

Private Sub cmdcancel_Click()
con.Close
Unload Me
End Sub

Private Sub cmdCarbendazim_Click()
With frmProductdetails.grd
.AddItem "206" & vbTab & "Carbendazim" & vbTab & "Indogulf CropSciences Ltd" & vbTab & "5KG" & vbTab & "1800"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With

End Sub

Private Sub cmdChrysoperla_Click()
With frmProductdetails.grd
.AddItem "202" & vbTab & "Chrysoperla" & vbTab & "Bayer House" & vbTab & "5KG" & vbTab & "1400"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With

End Sub

Private Sub cmdcottonseed_Click()

With frmProductdetails.grd
.AddItem "301" & vbTab & "Cotton Seed" & vbTab & "GreenGold Seeds" & vbTab & "5KG" & vbTab & "1200"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With

frmProductdetails.grd.TextMatrix(1, 5) = Text1.Text
End Sub

Private Sub cmdgroundnutsseed_Click()
With frmProductdetails.grd
.AddItem "306" & vbTab & "Groundnuts Seed" & vbTab & "Mahabeej" & vbTab & "5KG" & vbTab & "1800"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdjawarseed_Click()
With frmProductdetails.grd
.AddItem "305" & vbTab & "Jowar Seed" & vbTab & "National Agro Industries" & vbTab & "5KG" & vbTab & "1100"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdJeevamrut_Click()
With frmProductdetails.grd
.AddItem "102" & vbTab & "Jeevamrut" & vbTab & "National Fertilizer Ltd" & vbTab & "5KG" & vbTab & "1400"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdNeemleafmanure_Click()
With frmProductdetails.grd
.AddItem "205" & vbTab & "Neem Leaf manure" & vbTab & "Jain Agro Industries" & vbTab & "5KG" & vbTab & "1100"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdNitroganFixation_Click()
With frmProductdetails.grd
.AddItem "104" & vbTab & "Nitrogan Fixation" & vbTab & "Zuri Agro Chemicals Ltd" & vbTab & "5KG" & vbTab & "1600"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdPalash_Click()
With frmProductdetails.grd
.AddItem "204" & vbTab & "Palash" & vbTab & "Kapoor Pestisides" & vbTab & "5KG" & vbTab & "1600"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdPhospherous_Click()
With frmProductdetails.grd
.AddItem "103" & vbTab & "Phospherous" & vbTab & "Petrochemical Corporation Ltd" & vbTab & "5KG" & vbTab & "1500"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdPSBManure_Click()
With frmProductdetails.grd
.AddItem "105" & vbTab & "PSB Manure(Compost)" & vbTab & "National Fertilizer Ltd" & vbTab & "5KG" & vbTab & "1100"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdriceseed_Click()
With frmProductdetails.grd
.AddItem "302" & vbTab & "Rice Seed" & vbTab & "Mahabeej" & vbTab & "5KG" & vbTab & "1400"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub


Private Sub cmdRockPhosphate_Click()
With frmProductdetails.grd
.AddItem "106" & vbTab & "Rock-Phosphate" & vbTab & "Mangalore Chemical & fertilizer Ltd" & vbTab & "5KG" & vbTab & "1800"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdsugarcaneseed_Click()
With frmProductdetails.grd
.AddItem "304" & vbTab & "Sugarcane Seed" & vbTab & "GreenGold Seeds" & vbTab & "5KG" & vbTab & "1600"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdSuphala_Click()
With frmProductdetails.grd
.AddItem "203" & vbTab & "Suphala" & vbTab & "Insecticides Indis Ltd" & vbTab & "5KG" & vbTab & "1500"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdTricordrama_Click()
With frmProductdetails.grd
.AddItem "201" & vbTab & "Tricordrama" & vbTab & "Sameer Agro" & vbTab & "5KG" & vbTab & "1200"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub

Private Sub cmdwheatseed_Click()
With frmProductdetails.grd
.AddItem "303" & vbTab & "Wheat Seed" & vbTab & "Mahyco" & vbTab & "5KG" & vbTab & "1500"

End With

With frmProductdetails.grd
.Col = 1
.row = 1
End With
End Sub




Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BHAVESH\Desktop\Vb 5 sem project\Vb5sempro.mdb;Persist Security Info=False"
con.Open


End Sub
