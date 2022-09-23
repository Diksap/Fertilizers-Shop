VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmbilldetails 
   Caption         =   "Form1"
   ClientHeight    =   9675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19545
   LinkTopic       =   "Form1"
   Picture         =   "frmbilldetails.frx":0000
   ScaleHeight     =   9675
   ScaleWidth      =   19545
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcancel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   16080
      Picture         =   "frmbilldetails.frx":3E11F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtsearch 
      Height          =   615
      Left            =   9480
      TabIndex        =   2
      Top             =   3000
      Width           =   3855
   End
   Begin VB.CommandButton cmdsearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14160
      Picture         =   "frmbilldetails.frx":3F55D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   17520
      ScaleHeight     =   795
      ScaleWidth      =   2835
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmbilldetails.frx":3FF36
      Height          =   3735
      Left            =   5280
      TabIndex        =   0
      Top             =   4440
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   6588
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   9
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Search "
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
      Left            =   14400
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   16200
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "All Bill Records"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   10680
      TabIndex        =   5
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   3
      Top             =   3000
      Width           =   2895
   End
End
Attribute VB_Name = "frmbilldetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Private Sub cmdcancel_Click()
con.Close
Unload Me
End Sub

Private Sub cmdsearch_Click()
Dim no As Double
no = Val(txtsearch.Text)
rs.MoveFirst
Do Until rs.EOF
If no = rs.Fields(3) Then
MsgBox "Record has been Found", vbInformation, "Found"

Exit Sub
End If
rs.MoveNext
Loop
MsgBox "Record Not Found", vbExclamation + vbOKOnly, "Not Found"
End Sub
'validation for search textbox
Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc("A") And KeyAscii < Asc("Z") Or KeyAscii > Asc("a") And KeyAscii < Asc("z") Then
KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BHAVESH\Desktop\Vb 5 sem project\Vb5sempro.mdb;Persist Security Info=False"
con.Open
rs.CursorLocation = adUseClient
rs.Open "select * from Bill", con, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub
