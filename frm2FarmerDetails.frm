VERSION 5.00
Begin VB.Form frmFarmerRegistration 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Farmer Registration"
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19545
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "frm2FarmerDetails.frx":0000
   ScaleHeight     =   9060
   ScaleWidth      =   19545
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Update"
      Height          =   735
      Left            =   17040
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdAddnew 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Add New"
      Height          =   735
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmssearch 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Search"
      Height          =   735
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Delete"
      Height          =   735
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2280
      Width           =   3975
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Last"
      Height          =   735
      Left            =   17040
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Previous"
      Height          =   735
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Next"
      Height          =   735
      Left            =   17040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0FFC0&
      Caption         =   "First"
      Height          =   735
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cancel"
      Height          =   735
      Left            =   17040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdsubmit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Submit"
      Height          =   735
      Left            =   17040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtaddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   4800
      Width           =   3975
   End
   Begin VB.TextBox txtadharcardno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      MaxLength       =   14
      TabIndex        =   8
      Top             =   6480
      Width           =   3975
   End
   Begin VB.TextBox txtemailid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   5640
      Width           =   3975
   End
   Begin VB.TextBox txtmobileno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   6
      Top             =   3960
      Width           =   3975
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Adhar-Card No Format  : 0000 0000 0000"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2520
      TabIndex        =   23
      Top             =   7200
      Width           =   6015
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Farmer ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   17
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Farmer Information"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   6000
      TabIndex        =   9
      Top             =   360
      Width           =   7815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Adhar-Card No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile  No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Farmer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   3120
      Width           =   2655
   End
End
Attribute VB_Name = "frmFarmerRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection

Private Sub cmdaddnew_Click()
txtid.Text = ""
txtname.Text = ""
txtaddress.Text = ""
txtmobileno.Text = ""
txtemailid.Text = ""
txtadharcardno.Text = ""
rs.MoveLast
'to get auto generate farmer ID
txtid.Text = Val(rs(0)) + 1
rs.AddNew
End Sub

Private Sub cmdcancel_Click()
con.Close
Unload Me
End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BHAVESH\Desktop\Vb 5 sem project\Vb5sempro.mdb;Persist Security Info=False"
con.Open
rs.Open "select * from Farmer", con, adOpenDynamic, adLockOptimistic
End Sub

Private Sub cmdsubmit_Click()
rs("F_id") = txtid.Text
rs("F_Name") = txtname.Text
rs("F_Address") = txtaddress.Text
rs("F_MobileNo") = txtmobileno.Text
rs("F_EmailId") = txtemailid.Text
rs("F_AdharCardno") = txtadharcardno.Text
rs.Update
MsgBox "Record has been Save Sucessfully", vbInformation
End Sub



Private Sub cmdupdate_Click()
rs("F_Name") = txtname.Text
rs("F_Address") = txtaddress.Text
rs("F_MobileNo") = txtmobileno.Text
rs("F_EmailId") = txtemailid.Text
rs("F_AdharCardno") = txtadharcardno.Text
rs.Update
MsgBox "Record has been Update Sucessfully ", vbInformation
End Sub

Private Sub cmddelete_Click()
rs.Delete
MsgBox "Record has been Delete Sucessfully ", vbInformation
rs.MoveNext
End Sub

Private Sub cmdFirst_Click()
rs.MoveFirst
Call display
End Sub

Private Sub cmdLast_Click()
rs.MoveLast
Call display
End Sub

Private Sub cmdNext_Click()
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

Private Sub cmdPrevious_Click()
If Not rs.BOF Then
rs.MovePrevious
If rs.BOF Then
rs.MoveFirst
MsgBox "You Reached on First Record", vbInformation
Else
Call display
End If
End If
End Sub

Private Sub cmssearch_Click()
Dim no As Double
no = Val(InputBox("Search By Mobile No"))
rs.MoveFirst
Do Until rs.EOF
If no = rs.Fields(3) Then
MsgBox "Record Has Been Found", vbInformation, "Found"
Call display
Exit Sub
End If
rs.MoveNext
Loop
MsgBox "Record Not Found", vbExclamation + vbOKOnly, "Not Found"

End Sub

Public Sub display()
txtid.Text = rs("F_ID")
txtname.Text = rs("F_Name")
txtaddress.Text = rs("F_Address")
txtmobileno.Text = rs("F_MobileNo")
txtemailid.Text = rs("F_EmailId")
txtadharcardno.Text = rs("F_AdharCardno")
End Sub
'Email Validation
Private Sub txtemailid_Validate(Cancel As Boolean)
Dim str As String
Dim a As String
Dim b As String
                str = txtemailid.Text
                a = InStr(1, str, "@")
                b = InStr(a + 2, str, ".")
                If a = 0 Or b = 0 Then
                MsgBox ("Invalid email")
                End If

End Sub
'Mobile No Validation
Private Sub txtmobileno_Change()
Dim numval As String
If IsNumeric(txtmobileno.Text) Then
numval = txtmobileno.Text
Else
txtmobileno.Text = CStr(numval)
End If
End Sub
'Mobile No Validation
Private Sub txtmobileno_Validate(Cancel As Boolean)
If KeyAscii > Asc(0) And KeyAscii < Asc(9) Then
KeyAscii = 0
End If
If Len(txtmobileno.Text) < 10 Then
MsgBox ("Please enter 10 digit Mobile Number")
End If
End Sub
'Address Validation
Private Sub txtaddreess_KeyPress(KeyAscii As Integer)
KeyAscii = add(KeyAscii)
End Sub
'use for address validation
Public Function add(KeyAscii As Integer) As Integer
Dim add1 As Boolean
add1 = Chr(KeyAscii) Like "[""a-z A-Z 0-9,.\/]"
If add1 = False And KeyAscii <> 8 Then
KeyAscii = 0
End If
address = KeyAscii
End Function
'Adharcardno validation
Private Sub txtadharcardno_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57, 8, 32
    '8 is allow Backspace
    '32 allow space
    Case Else
    KeyAscii = 0
    End Select
End Sub
'use for adharcard validation
Public Function adharcardno(KeyAscii As Integer) As Integer
Dim stat As String
stat = "0123456789"
If InStr(stat, Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
KeyAscii = 0
End If
adharcardno = KeyAscii
End Function
'Name validation
Private Sub txtname_KeyPress(KeyAscii As Integer)
KeyAscii = character(KeyAscii)
End Sub
'use for name validation
Public Function character(KeyAscii As Integer) As Integer
Dim var As Boolean
var = Chr(KeyAscii) Like "[""a-z A-Z]"
If var = False And KeyAscii <> 8 Then
KeyAscii = 0
End If
character = KeyAscii
End Function

