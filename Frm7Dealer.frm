VERSION 5.00
Begin VB.Form FrmDealerRegistration 
   Caption         =   "Form1"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18045
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   Picture         =   "Frm7Dealer.frx":0000
   ScaleHeight     =   9630
   ScaleWidth      =   18045
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8160
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0FFC0&
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton cmdsubmit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3840
      Width           =   2055
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
      Left            =   6480
      MaxLength       =   14
      TabIndex        =   14
      Top             =   8160
      Width           =   3735
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
      Left            =   6480
      TabIndex        =   13
      Top             =   7320
      Width           =   3735
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton cmdaddnew 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1935
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
      Left            =   6480
      MaxLength       =   10
      TabIndex        =   8
      Top             =   5640
      Width           =   3735
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
      Left            =   6480
      TabIndex        =   7
      Top             =   6480
      Width           =   3735
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
      Left            =   6480
      TabIndex        =   6
      Top             =   4800
      Width           =   3735
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3960
      Width           =   3735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Adhar-Card No Format  : 0000 0000 0000"
      Height          =   375
      Left            =   5040
      TabIndex        =   23
      Top             =   8880
      Width           =   5535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Adhar-card No"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   8160
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer ID"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Email-Id"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   2880
      TabIndex        =   1
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Information"
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
      Height          =   1095
      Left            =   7200
      TabIndex        =   0
      Top             =   2160
      Width           =   10575
   End
End
Attribute VB_Name = "FrmDealerRegistration"
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
'for auto DealerId generation
txtid.Text = Val(rs(0)) + 1
rs.AddNew
End Sub
Private Sub cmdsubmit_Click()
rs("D_id") = txtid.Text
rs("D_Name") = txtname.Text
rs("D_Address") = txtaddress.Text
rs("D_MobileNo") = txtmobileno.Text
rs("D_EmailId") = txtemailid.Text
rs("D_AdharCardno") = txtadharcardno.Text
rs.Update
MsgBox "Record has been Save Sucessfully", vbInformation
End Sub
Private Sub cmdcancel_Click()
con.Close
Unload Me
End Sub
Private Sub cmdupdate_Click()
rs("D_Name") = txtname.Text
rs("D_Address") = txtaddress.Text
rs("D_MobileNo") = txtmobileno.Text
rs("D_EmailId") = txtemailid.Text
rs("D_AdharCardno") = txtadharcardno.Text
rs.Update
MsgBox "Record has been Update Sucessfully", vbInformation
End Sub
Private Sub cmddelete_Click()
rs.Delete
MsgBox "Record has been Delete Sucessfully", vbInformation
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
Private Sub cmdsearch_Click()
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

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BHAVESH\Desktop\Vb 5 sem project\Vb5sempro.mdb;Persist Security Info=False"
con.Open
rs.Open "select * from Dealer", con, adOpenDynamic, adLockOptimistic
End Sub

Public Sub display()
txtid.Text = rs("D_ID")
txtname.Text = rs("D_Name")
txtaddress.Text = rs("D_Address")
txtmobileno.Text = rs("D_MobileNo")
txtemailid.Text = rs("D_EmailId")
txtadharcardno.Text = rs("D_AdharCardno")
End Sub



'Email validation
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
'AdharcardNo validation
Private Sub txtadharcardno_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57, 8, 32
    '8 is allow Backspace
    '32 allow space
    Case Else
    KeyAscii = 0
    End Select
End Sub
'Use for Mobile No validation
Private Sub txtmobileno_Change()
Dim numval As String
If IsNumeric(txtmobileno.Text) Then
numval = txtmobileno.Text
Else
txtmobileno.Text = CStr(numval)
End If
End Sub
'Mobile No validation
Private Sub txtmobileno_Validate(Cancel As Boolean)
If KeyAscii > Asc(0) And KeyAscii < Asc(9) Then
KeyAscii = 0
End If
'10 digit limit
If Len(txtmobileno.Text) < 10 Then
MsgBox ("Please enter 10 digit Mobile Number")
End If
End Sub
'Validation for Name
Private Sub txtname_KeyPress(KeyAscii As Integer)
KeyAscii = character(KeyAscii)
End Sub
Public Function character(KeyAscii As Integer) As Integer
Dim var As Boolean
var = Chr(KeyAscii) Like "[""a-z A-Z]"
If var = False And KeyAscii <> 8 Then
KeyAscii = 0
End If
character = KeyAscii
End Function







