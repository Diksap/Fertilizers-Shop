VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAddFertilizers 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "frmAddFertilizers.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtfstock 
      Height          =   495
      Left            =   6600
      TabIndex        =   19
      Top             =   6720
      Width           =   3015
   End
   Begin VB.TextBox txtfprice 
      Height          =   495
      Left            =   6600
      TabIndex        =   18
      Top             =   5280
      Width           =   3015
   End
   Begin VB.TextBox txtfquantityperbag 
      Height          =   495
      Left            =   6600
      TabIndex        =   17
      Top             =   6000
      Width           =   3015
   End
   Begin VB.TextBox txtfcompanyname 
      Height          =   495
      Left            =   6600
      TabIndex        =   16
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox txtfertilizername 
      Height          =   495
      Left            =   6600
      TabIndex        =   15
      Top             =   3840
      Width           =   3015
   End
   Begin VB.TextBox txtfertilizerno 
      Height          =   495
      Left            =   6600
      TabIndex        =   14
      Top             =   3120
      Width           =   3015
   End
   Begin VB.CommandButton cmdfaddnew 
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
      Height          =   735
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdfcancel 
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
      Left            =   15600
      Picture         =   "frmAddFertilizers.frx":2D47E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdfdelete 
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
      Height          =   735
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtfsearch 
      Height          =   615
      Left            =   9120
      TabIndex        =   9
      Top             =   7680
      Width           =   3975
   End
   Begin VB.CommandButton cmdfsearch 
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
      Left            =   13440
      Picture         =   "frmAddFertilizers.frx":2E8BC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdflast 
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
      Height          =   735
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdfprevious 
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
      Height          =   735
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdfnext 
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
      Height          =   735
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdffirst 
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
      Height          =   735
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdfsubmit 
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
      Height          =   735
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdfupdate 
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
      Height          =   735
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAddFertilizers.frx":2F295
      Height          =   1815
      Left            =   3960
      TabIndex        =   8
      Top             =   8400
      Width           =   9135
      _ExtentX        =   16113
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
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   16560
      ScaleHeight     =   1155
      ScaleWidth      =   3075
      TabIndex        =   26
      Top             =   9360
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   15720
      TabIndex        =   28
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   13560
      TabIndex        =   27
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label15 
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
      Height          =   375
      Left            =   3840
      TabIndex        =   25
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label14 
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
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label13 
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
      Left            =   3840
      TabIndex        =   23
      Top             =   5280
      Width           =   2295
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
      Left            =   3840
      TabIndex        =   22
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fertilizer Name"
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
      Left            =   3840
      TabIndex        =   21
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fertilizer No"
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
      Left            =   3840
      TabIndex        =   20
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By Fertilizers Name"
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
      Left            =   4080
      TabIndex        =   10
      Top             =   7680
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fertilizers Details"
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
      Left            =   8040
      TabIndex        =   0
      Top             =   2040
      Width           =   7935
   End
End
Attribute VB_Name = "frmAddFertilizers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection

Private Sub cmdfdelete_Click()
rs.Delete
MsgBox "Record has been Deleted Succesfully", vbInformation
rs.MoveNext
End Sub

Private Sub cmdfaddnew_Click()
txtfertilizerno.Text = ""
txtfertilizername.Text = ""
txtfcompanyname.Text = ""
txtfprice.Text = ""
txtfquantityperbag.Text = ""
txtfstock.Text = ""
rs.MoveLast
'Auto generate Fertilizer ID
txtfertilizerno.Text = Val(rs(0)) + 1
rs.AddNew
End Sub

Private Sub cmdfcancel_Click()
con.Close
Unload Me
End Sub

Private Sub cmdfsearch_Click()
Dim a As String
a = txtfsearch.Text
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

'Public Sub display()
'txtfertilizerno.Text = rs("F_ID")
'txtfertilizername.Text = rs("F_Name")
'txtfcompanyname.Text = rs("F_Address")
'txtfprice.Text = rs("F_MobileNo")
'txtfquantityperbag.Text = rs("F_EmailId")
'txtfstock.Text = rs("F_AdharCardno")
'End Sub
'Search textbox validation
Private Sub txtfsearch_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc(0) And KeyAscii < Asc(9) Then
KeyAscii = 0
End If
End Sub

Private Sub cmdffirst_Click()
rs.MoveFirst
Call display
End Sub

Private Sub cmdflast_Click()
rs.MoveLast
Call display
End Sub

Private Sub cmdfnext_Click()
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

Private Sub cmdfprevious_Click()
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

Private Sub cmdfsubmit_Click()
rs("FertilizerNo") = txtfertilizerno.Text
rs("FertilizerName") = txtfertilizername.Text
rs("FCompanyName") = txtfcompanyname.Text
rs("FQuantityPerBag") = txtfquantityperbag.Text
rs("FPrice") = txtfprice.Text
rs("FStock") = txtfstock.Text
rs.Update
MsgBox "Record has been Save Succesfully", vbInformation
End Sub

Private Sub cmdfUpdate_Click()
rs("FertilizerName") = txtfertilizername.Text
rs("FCompanyName") = txtfcompanyname.Text
rs("FQuantityPerBag") = txtfquantityperbag.Text
rs("FPrice") = txtfprice.Text
rs("FStock") = txtfstock.Text
rs.Update
MsgBox "Record has been Update Succesfully", vbInformation
End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BHAVESH\Desktop\Vb 5 sem project\Vb5sempro.mdb;Persist Security Info=False"
con.Open
rs.CursorLocation = adUseClient
rs.Open "select * from AddFertilizers", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub

Public Sub display()
txtfertilizerno.Text = rs("FertilizerNo")
txtfertilizername.Text = rs("FertilizerName")
txtfcompanyname.Text = rs("FCompanyName")
txtfquantityperbag.Text = rs("FQuantityPerBag")
txtfprice.Text = rs("FPrice")
txtfstock.Text = rs("FStock")
End Sub
'Company name validation
Private Sub txtfcompanyname_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc(0) And KeyAscii < Asc(9) Then
KeyAscii = 0
End If
End Sub
'Price validation
Private Sub txtfprice_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc("A") And KeyAscii < Asc("Z") Or KeyAscii > Asc("a") And KeyAscii < Asc("z") Then
KeyAscii = 0
End If
End Sub
'Name validation
Private Sub txtfertilizername_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc(0) And KeyAscii < Asc(9) Then
KeyAscii = 0
End If
End Sub
'Stock validation
Private Sub txtfstock_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc("A") And KeyAscii < Asc("Z") Or KeyAscii > Asc("a") And KeyAscii < Asc("z") Then
KeyAscii = 0
End If
End Sub
