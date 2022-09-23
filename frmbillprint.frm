VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmbillprint 
   Caption         =   "Form2"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13710
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      Top             =   3960
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   2990
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
End
Attribute VB_Name = "frmbillprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Form_Load()
Dim MyPass As String
Dim MyDataFile As String
Dim InvSql As String
Dim strCon As String

    'MyPass = ""
    'MyDataFile = App.Path & "\DataFile\Northwind.mdb"
    'strCon = "provider=microsoft.jet.oledb.4.0;data source=" _
    '& MyDataFile & ";" & "Jet OLEDB:Database Password=" & MyPass & ";"

con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BHAVESH\Desktop\Vb 5 sem project\Vb5sempro.mdb;Persist Security Info=False"
con.Open
    rs.Open "SELECT Farmer.F_Name, Farmer.F_MobNo, " _
    & "Employees.FirstName & Space(1) & Employees.LastName AS Salesperson, " _
    & "Orders.OrderID, Orders.OrderDate, " _
    & "[Order Details].ProductID, Products.ProductName, [Order Details].UnitPrice, " _
    & "[Order Details].Quantity, [Order Details].Discount, " _
    & "CCur([Order Details].UnitPrice*_"
    [Quantity]*(1-[Discount])/100)*100 AS ExtendedPrice, " _
    & "Orders.Freight " _
    & "FROM Products INNER JOIN ((Employees INNER JOIN " _
    & "(Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID) " _
    & "ON Employees.EmployeeID = Orders.EmployeeID) " _
    & "INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID) " _
    & "ON Products.ProductID = [Order Details].ProductID;"

    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open strCon

    Set rs = New ADODB.Recordset
    rs.Open InvSql, cn, adOpenStatic, adLockOptimistic

    Set datGrid.DataSource = rs
End Sub
