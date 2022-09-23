VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmpracticebill 
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   5295
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9340
      _Version        =   393216
   End
End
Attribute VB_Name = "frmpracticebill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetGridProperties()
With grid
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
.Text = "Product ID"

.Col = 1
.Text = "Product Name"

.Col = 2
.Text = "Product Company Name"

.Col = 3
.Text = "Product Weight"

.Col = 4
.Text = "Product Price"

End With
End Sub

Private Sub Form_Load()
Call SetGridProperties


End Sub


