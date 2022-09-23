VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAddInsecticides 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   Picture         =   "frmAddInsecticides.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtistock 
      Height          =   495
      Left            =   3240
      TabIndex        =   19
      Top             =   6960
      Width           =   3255
   End
   Begin VB.TextBox txtiquantityperbag 
      Height          =   495
      Left            =   3240
      TabIndex        =   18
      Top             =   6120
      Width           =   3255
   End
   Begin VB.TextBox txtiprice 
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   5280
      Width           =   3255
   End
   Begin VB.TextBox txticompanyname 
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   4440
      Width           =   3255
   End
   Begin VB.TextBox txtinsecticidename 
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox txtinsecticideno 
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton cmdiaddnew 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Add New"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtisearch 
      Height          =   495
      Left            =   12360
      TabIndex        =   11
      Top             =   4080
      Width           =   3975
   End
   Begin VB.CommandButton cmdisearch 
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
      Left            =   16560
      Picture         =   "frmAddInsecticides.frx":1A528
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdilast 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Last"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdiprevious 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Previous"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdinext 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Next"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdifirst 
      BackColor       =   &H00C0FFC0&
      Caption         =   "First"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdicancel 
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
      Left            =   18240
      Picture         =   "frmAddInsecticides.frx":1AF01
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdidelete 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Delete"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdisubmit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Submit"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdiUpdate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Update"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   2055
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   19080
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   26
      Top             =   7800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAddInsecticides.frx":1C33F
      Height          =   1815
      Left            =   12000
      TabIndex        =   9
      Top             =   5160
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3201
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
   Begin VB.Label Label4 
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   13320
      TabIndex        =   29
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   16680
      TabIndex        =   28
      Top             =   3120
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   18360
      TabIndex        =   27
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Per Bag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Insecticides Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Insecticides No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   " Insecticides Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12600
      TabIndex        =   12
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Insecticides Details"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   6720
      TabIndex        =   0
      Top             =   1200
      Width           =   8415
   End
End
Attribute VB_Name = "frmAddInsecticides"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Private Sub cmdidelete_Click()
rs.Delete
MsgBox "Record has been Deleted Succesfully", vbInformation
rs.MoveNext
End Sub

Private Sub cmdiaddnew_Click()
txtinsecticideno.Text = ""
txtinsecticidename.Text = ""
txticompanyname.Text = ""
txtiprice.Text = ""
txtiquantityperbag.Text = ""
txtistock.Text = ""
rs.MoveLast
'auto generate ID
txtinsecticideno.Text = Val(rs(0)) + 1
rs.AddNew
End Sub

Private Sub cmdicancel_Click()
con.Close
Unload Me
End Sub

Private Sub cmdisearch_Click()
Dim a As String
a = txtisearch.Text
rs.MoveFirst
Do Until rs.EOF
If a = rs.Fields(1) Then
MsgBox "Record has been Found", vbInformation, "Found"
Call display
Exit Sub
End If
rs.MoveNext
Loop
MsgBox "Record Not Found", vbExclamation + vbOKOnly, "Not Found"
End Sub
'validation for search text box
Private Sub txtisearch_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc(0) And KeyAscii < Asc(9) Then
KeyAscii = 0
End If
End Sub

Private Sub cmdifirst_Click()
rs.MoveFirst
Call display
End Sub

Private Sub cmdilast_Click()
rs.MoveLast
Call display
End Sub

Private Sub cmdinext_Click()
If Not rs.EOF Then
rs.MoveNext
If rs.EOF Then
rs.MoveLast
MsgBox "You Reached on Last Record", vbInformation
Else
Call display
End If
End If
End Sub

Private Sub cmdiprevious_Click()
If Not rs.BOF Then
rs.MovePrevious
If rs.BOF Then
rs.MoveFirst
MsgBox "You Reached on Last Record", vbInformation
Else
Call display
End If
End If
End Sub

Private Sub cmdisubmit_Click()
rs("InsecticideNo") = txtinsecticideno.Text
rs("InsecticideName") = txtinsecticidename.Text
rs("ICompanyName") = txticompanyname.Text
rs("IQuantityPerBag") = txtiquantityperbag.Text
rs("IPrice") = txtiprice.Text
rs("IStock") = txtistock.Text
rs.Update
MsgBox "Record has been Save Succesfully", vbInformation
End Sub

Private Sub cmdiUpdate_Click()
rs("InsecticideName") = txtinsecticidename.Text
rs("ICompanyName") = txticompanyname.Text
rs("IQuantityPerBag") = txtiquantityperbag.Text
rs("IPrice") = txtiprice.Text
rs("IStock") = txtistock.Text
rs.Update
MsgBox "Record has been Update Succesfully", vbInformation
End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BHAVESH\Desktop\Vb 5 sem project\Vb5sempro.mdb;Persist Security Info=False"
con.Open
rs.CursorLocation = adUseClient
rs.Open "select * from AddInsecticides", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub

Public Sub display()
txtinsecticideno.Text = rs("InsecticideNo")
txtinsecticidename.Text = rs("InsecticideName")
txticompanyname.Text = rs("ICompanyName")
txtiquantityperbag.Text = rs("IQuantityPerBag")
txtiprice.Text = rs("IPrice")
txtistock.Text = rs("IStock")
End Sub
'company name valdation
Private Sub txticompanyname_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc(0) And KeyAscii < Asc(9) Then
KeyAscii = 0
End If
End Sub
'price validation
Private Sub txtiprice_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc("A") And KeyAscii < Asc("Z") Or KeyAscii > Asc("a") And KeyAscii < Asc("z") Then
KeyAscii = 0
End If
End Sub
'Insecticides name validation
Private Sub txtinsecticidename_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc(0) And KeyAscii < Asc(9) Then
KeyAscii = 0
End If
End Sub
'stock validation
Private Sub txtistock_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc("A") And KeyAscii < Asc("Z") Or KeyAscii > Asc("a") And KeyAscii < Asc("z") Then
KeyAscii = 0
End If
End Sub



