VERSION 5.00
Begin VB.Form frmReport 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox lstProduct 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   270
      TabIndex        =   3
      Top             =   2655
      Width           =   4215
   End
   Begin VB.ListBox lstCategory 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      ItemData        =   "frmReport.frx":0000
      Left            =   270
      List            =   "frmReport.frx":0007
      TabIndex        =   2
      Top             =   585
      Width           =   4215
   End
   Begin VB.TextBox txtCategory 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1530
      TabIndex        =   1
      Top             =   225
      Width           =   2955
   End
   Begin VB.CommandButton cmdProductList 
      Caption         =   "Show Productsreport"
      Height          =   510
      Left            =   225
      TabIndex        =   0
      Top             =   5400
      Width           =   4245
   End
   Begin VB.Label Label1 
      Caption         =   "Category :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   330
      TabIndex        =   5
      Top             =   225
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Products :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   330
      TabIndex        =   4
      Top             =   2325
      Width           =   1215
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConnection As ADODB.Connection

Public Function openTheDatabase() As Boolean

Dim sConnectionString As String
On Error GoTo dbError
'-- connection --
Set adoConnection = New ADODB.Connection
sConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;" _
                 & "Data Source=" & App.Path & "\Nwind.mdb"
adoConnection.Open sConnectionString
openTheDatabase = True
Exit Function
dbError:
MsgBox (Err.Description)
openTheDatabase = False
End Function

Private Sub cmdProductList_Click()
Dim adoRS As ADODB.Recordset
Dim strSql As String
Dim strWhere As String

strSql = "SELECT Products.ProductID, " 'ProductID           = Value (0)
strSql = strSql & "Products.ProductName, " 'ProductName     = Value (1)
strSql = strSql & "Categories.CategoryName " 'CategoryName  = Value (2)
strSql = strSql & "FROM Categories RIGHT JOIN "
strSql = strSql & "Products ON Categories.CategoryID= "
strSql = strSql & "Products.CategoryID "

'Find What? :
strWhere = strWhere & "Categories.CategoryName= '" & txtCategory & "'"

'Join the sql:
strSql = strSql & " WHERE " & strWhere

Set adoRS = adoConnection.Execute(strSql)

With DataReport1
    .DataMember = vbNullString
    Set .DataSource = adoRS
    .Caption = "Productlist..."

'Detail:
    With .Sections("Bereich1").Controls
        .Item("rptArtNr").DataField = adoRS.Fields(0).Name
        .Item("rptArtName").DataField = adoRS.Fields(1).Name
        .Item("rptArtKat").DataField = adoRS.Fields(2).Name
    End With
    
'Header
    With .Sections("Bereich2").Controls
    .Item("Be1").Caption = adoRS.Fields(2).Value
    End With
    .Show
End With
Set adoRS = Nothing
End Sub

Private Sub Form_Load()
DoEvents

    If (Not openTheDatabase()) Then
    MsgBox "No Database found!"
    Exit Sub
End If

Call updateListBoxes
txtCategory = lstCategory.Text

End Sub



'// Listboxes:
Public Sub updateListBoxes()
Dim adoTempRecordset As ADODB.Recordset
Dim sSql As String
'--------------------------------------
'-- Categories Listbox --
'--------------------------------------
Set adoTempRecordset = New ADODB.Recordset
adoTempRecordset.CursorLocation = adUseClient
adoTempRecordset.Open _
   "SELECT * FROM Categories ORDER BY CategoryName", _
   adoConnection
lstCategory.Clear
With adoTempRecordset
  If .RecordCount > 0 Then .MoveFirst
  While Not .EOF
    lstCategory.AddItem !CategoryName
    lstCategory.ItemData(lstCategory.NewIndex) = !CategoryID
    .MoveNext
  Wend
End With
lstCategory.ListIndex = 0
adoTempRecordset.Close

'--------------------------
'Products Listbox
'--------------------------
sSql = "SELECT ProductName, ProductID FROM Products"
sSql = sSql & " WHERE CategoryID = " & lstCategory.ItemData(lstCategory.ListIndex)
adoTempRecordset.Open sSql, adoConnection
lstProduct.Clear
With adoTempRecordset
   If .RecordCount > 0 Then .MoveFirst
   While Not .EOF
     lstProduct.AddItem !ProductName
     lstProduct.ItemData(lstProduct.NewIndex) = !ProductID
     .MoveNext
   Wend
End With
lstProduct.ListIndex = 0
End Sub

Private Sub lstCategory_Click()
Dim adoTempProduct As ADODB.Recordset

Dim sSql As String

sSql = "SELECT ProductName, ProductID FROM Products"
sSql = sSql & " WHERE CategoryID = "
sSql = sSql & lstCategory.ItemData(lstCategory.ListIndex)

Set adoTempProduct = New ADODB.Recordset
adoTempProduct.CursorLocation = adUseClient
adoTempProduct.Open sSql, adoConnection


If adoTempProduct.RecordCount > 0 Then _
    adoTempProduct.MoveLast
adoTempProduct.MoveFirst
lstProduct.Clear
With adoTempProduct
   If .RecordCount > 0 Then .MoveFirst
   While Not .EOF
     lstProduct.AddItem !ProductName
     lstProduct.ItemData(lstProduct.NewIndex) = !ProductID
     .MoveNext
   Wend
End With
lstProduct.ListIndex = 0
adoTempProduct.Close
txtCategory = lstCategory.Text

End Sub

'\\ Listboxes End
