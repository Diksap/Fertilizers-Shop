VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAddSeed 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "frmAddSeed.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   3840
      TabIndex        =   25
      Top             =   8520
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3413
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
   Begin VB.TextBox txtsstock 
      Height          =   495
      Left            =   7440
      TabIndex        =   18
      Top             =   6360
      Width           =   3495
   End
   Begin VB.TextBox txtsprice 
      Height          =   495
      Left            =   7440
      TabIndex        =   17
      Top             =   4920
      Width           =   3495
   End
   Begin VB.TextBox txtsquantityperbag 
      Height          =   495
      Left            =   7440
      TabIndex        =   16
      Top             =   5640
      Width           =   3495
   End
   Begin VB.TextBox txtscompanyname 
      Height          =   495
      Left            =   7440
      TabIndex        =   15
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox txtseedname 
      Height          =   495
      Left            =   7440
      TabIndex        =   14
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox txtseedno 
      Height          =   495
      Left            =   7440
      TabIndex        =   13
      Top             =   2760
      Width           =   3495
   End
   Begin VB.CommandButton cmdsaddnew 
      BackColor       =   &H8000000E&
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdslast 
      BackColor       =   &H8000000E&
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
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   2055
   End
   Begin VB.TextBox txtssearch 
      Height          =   615
      Left            =   8400
      TabIndex        =   9
      Top             =   7800
      Width           =   3855
   End
   Begin VB.CommandButton cmdssearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13080
      Picture         =   "frmAddSeed.frx":3EEC0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdsdelete 
      BackColor       =   &H8000000E&
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
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdsprevious 
      BackColor       =   &H8000000E&
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdsnext 
      BackColor       =   &H8000000E&
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
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdsfirst 
      BackColor       =   &H8000000E&
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdscancel 
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
      Left            =   15240
      Picture         =   "frmAddSeed.frx":3F899
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdssubmit 
      BackColor       =   &H8000000E&
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
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdsupdate 
      BackColor       =   &H8000000E&
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   2055
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
      Height          =   375
      Left            =   13320
      TabIndex        =   27
      Top             =   7920
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
      Left            =   15480
      TabIndex        =   26
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
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
      Left            =   4560
      TabIndex        =   24
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Per Bag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4560
      TabIndex        =   23
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   4560
      TabIndex        =   22
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
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
      Left            =   4560
      TabIndex        =   21
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Seed Name"
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
      Left            =   4560
      TabIndex        =   20
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Seed No"
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
      Left            =   4560
      TabIndex        =   19
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By Seed Name"
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
      Left            =   4200
      TabIndex        =   10
      Top             =   7800
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seeds Details"
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
      Height          =   1215
      Left            =   8640
      TabIndex        =   0
      Top             =   1080
      Width           =   6975
   End
End
Attribute VB_Name = "frmAddSeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Private Sub cmdsdelete_Click()
rs.Delete
MsgBox "Record has been Deleted Succesfully", vbInformation
rs.MoveNext
End Sub

Private Sub cmdsaddnew_Click()
txtseedno.Text = ""
txtseedname.Text = ""
txtscompanyname.Text = ""
txtsprice.Text = ""
txtsquantityperbag.Text = ""
txtsstock.Text = ""
rs.MoveLast
'auto generate seed id
txtseedno.Text = Val(rs(0)) + 1
rs.AddNew
End Sub

Private Sub cmdscancel_Click()
con.Close
Unload Me
End Sub

Private Sub cmdssearch_Click()
Dim a As String
a = txtssearch.Text
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
'validation for search
Private Sub txtssearch_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc(0) And KeyAscii < Asc(9) Then
KeyAscii = 0
End If
End Sub

Private Sub cmdsfirst_Click()
rs.MoveFirst
Call display
End Sub

Private Sub cmdslast_Click()
rs.MoveLast
Call display
End Sub

Private Sub cmdsnext_Click()
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
Private Sub cmdsprevious_Click()
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

Private Sub cmdssubmit_Click()
rs("SeedNo") = txtseedno.Text
rs("SeedName") = txtseedname.Text
rs("SCompanyName") = txtscompanyname.Text
rs("SQuantityPerBag") = txtsquantityperbag.Text
rs("SPrice") = txtsprice.Text
rs("SStock") = txtsstock.Text
rs.Update
MsgBox "Record has been Save Succesfully", vbInformation
End Sub

Private Sub cmdsUpdate_Click()
rs("SeedName") = txtseedname.Text
rs("SCompanyName") = txtscompanyname.Text
rs("SQuantityPerBag") = txtsquantityperbag.Text
rs("SPrice") = txtsprice.Text
rs("SStock") = txtsstock.Text
rs.Update
MsgBox "Record has been Update Succesfully", vbInformation
End Sub
Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BHAVESH\Desktop\Vb 5 sem project\Vb5sempro.mdb;Persist Security Info=False"
con.Open
rs.CursorLocation = adUseClient
rs.Open "select * from AddSeeds", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub
Public Sub display()
txtseedno.Text = rs("SeedNo")
txtseedname.Text = rs("SeedName")
txtscompanyname.Text = rs("SCompanyName")
txtsquantityperbag.Text = rs("SQuantityPerBag")
txtsprice.Text = rs("SPrice")
txtsstock.Text = rs("SStock")
End Sub
'company name validation
Private Sub txtscompanyname_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc(0) And KeyAscii < Asc(9) Then
KeyAscii = 0
End If
End Sub
'price validation
Private Sub txtsprice_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc("A") And KeyAscii < Asc("Z") Or KeyAscii > Asc("a") And KeyAscii < Asc("z") Then
KeyAscii = 0
End If
End Sub
'seed name validation
Private Sub txtseedname_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc(0) And KeyAscii < Asc(9) Then
KeyAscii = 0
End If
End Sub
'stock validation
Private Sub txtsstock_KeyPress(KeyAscii As Integer)
If KeyAscii > Asc("A") And KeyAscii < Asc("Z") Or KeyAscii > Asc("a") And KeyAscii < Asc("z") Then
KeyAscii = 0
End If
End Sub

