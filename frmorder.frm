VERSION 5.00
Begin VB.Form frmselectproduct 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   17.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   10080
      TabIndex        =   4
      Top             =   6120
      Width           =   3015
   End
   Begin VB.CommandButton cmdinsecticides 
      Caption         =   "Insecticides"
      Height          =   975
      Left            =   8160
      TabIndex        =   2
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton cmdfertilizer 
      Caption         =   "Fertilizers"
      Height          =   855
      Left            =   4320
      TabIndex        =   1
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton cmdseeds 
      Caption         =   "Seeds"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Select Product"
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "frmselectproduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdcancel_Click()
'con.Close
Unload Me
End Sub

Private Sub cmdfertilizer_Click()
FrmFertilizers.Show
End Sub

Private Sub cmdinsecticides_Click()
Frminsecticides.Show
End Sub

Private Sub cmdseeds_Click()
FrmSeeds.Show
End Sub


