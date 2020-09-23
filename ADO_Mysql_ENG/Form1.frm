VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "ADO & mySQL"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Searching"
      Height          =   1215
      Left            =   3600
      TabIndex        =   8
      Top             =   3720
      Width           =   6375
      Begin VB.CommandButton Command4 
         Caption         =   "&Find"
         Height          =   615
         Left            =   3480
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Surname:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Records"
      Height          =   3615
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete selected"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1215
      End
      Begin MSComctlLib.ListView LV 
         Height          =   2775
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4895
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adding Records"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Surname:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   675
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Connecting to MySQL 4.0.2 using myODBC 3.5 & ADO 2.7
'without creating DSN in ODBC DataSources in ControlPanel
'by Josip Pejakovic - jpejakovic@yahoo.com - http://jp-net.web1000.com - ICQ# 127475388
'THIS CODE IS CREATED BY CROATIAN PROGRAMMER

'Instructions
'----------------
'- create database dbSample
'- create table Table
'- create two fields; Name (Varchar 20) and Surname (Varchar 25)

'Features
'------------
'- adding records
'- showing records in listview
'- removing selected records from list
'- searching records (you can search for records by name or by surname or both

Dim konekcija As ADODB.Connection
Dim rsTable As ADODB.Recordset
Dim ik
Dim l_item As ListItem

Public Sub AddRecord()
Set rsTable = New ADODB.Recordset
rsTable.Source = "Table"
rsTable.CursorLocation = adUseClient
rsTable.CursorType = adOpenDynamic
rsTable.LockType = adLockOptimistic
Set rsTable.ActiveConnection = konekcija
rsTable.Open

 
rsTable.AddNew
rsTable!Name = Text1
rsTable!Surname = Text2
rsTable.Update

Text1 = ""
Text2 = ""
Text1.SetFocus

End Sub

Private Sub Command1_Click()
AddRecord
LoadData
End Sub

Private Sub Command2_Click()
LoadData
End Sub

Private Sub Command3_Click()

rsTable.Close
Set rsTable = New ADODB.Recordset
rsTable.Source = "DELETE FROM Table WHERE Name = '" & LV.SelectedItem & "' AND Surname = '" & LV.SelectedItem.SubItems(1) & "'"
rsTable.CursorLocation = adUseClient
rsTable.CursorType = adOpenDynamic
rsTable.LockType = adLockOptimistic
Set rsTable.ActiveConnection = konekcija
rsTable.Open
rsTable.Close

LoadData
End Sub

Private Sub Command4_Click()
LV.ListItems.Clear

Set rsTable = New ADODB.Recordset
rsTable.CursorLocation = adUseClient
rsTable.CursorType = adOpenStatic
rsTable.LockType = adLockReadOnly
Set rsTable.ActiveConnection = konekcija

If Text3 <> "" Then
rsTable.Source = "SELECT * FROM Table WHERE Name = '" & Text3 & "'"
End If


If Text4 <> "" Then
rsTable.Source = "SELECT * FROM Table WHERE Surname = '" & Text4 & "'"
End If

If Text3 <> "" And Text4 <> "" Then
rsTable.Source = "SELECT * FROM Table WHERE Name = '" & Text3 & "' AND Surname = '" & Text4 & "'"
End If

rsTable.Open

Do Until rsTable.EOF
Set l_item = LV.ListItems.Add(, , rsTable!Name)
l_item.SubItems(1) = rsTable!Surname
rsTable.MoveNext
Loop
End Sub

Private Sub Form_Load()
'look here; we don't need DSN from ODBC control panel. We will put DSN in code
'In my opinion it's better because no one can purposely make changes in DSN options and
'you don't need to wrote code for creating DSN.

'It very simple to understand!!!

db_name = "dbSample"
db_server = "localhost"
db_port = ""    'default port is 3306
db_user = "root"
db_pass = ""

ik = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=;Extended Properties=" & Chr$(34) & "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & db_name & ";SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
Set konekcija = New ADODB.Connection
konekcija.Open ik
Set rsTable = New ADODB.Recordset
Set rsTable.ActiveConnection = konekcija
rsTable.CursorLocation = adUseClient
'####################################

With LV
.ColumnHeaders.Add , , "Name"
.ColumnHeaders.Add , , "Surname"
End With

LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set konekcija = Nothing
Set rsTable = Nothing
End Sub

Sub LoadData()
LV.ListItems.Clear
Set rsTable = New ADODB.Recordset
rsTable.CursorLocation = adUseClient
rsTable.CursorType = adOpenStatic
rsTable.LockType = adLockReadOnly
Set rsTable.ActiveConnection = konekcija

rsTable.Source = "SELECT * FROM Table ORDER BY Surname ASC"
rsTable.Open

Do Until rsTable.EOF
Set l_item = LV.ListItems.Add(, , rsTable!Name)
l_item.SubItems(1) = rsTable!Surname
rsTable.MoveNext
Loop

End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
Text1 = LV.SelectedItem
Text2 = LV.SelectedItem.SubItems(1)
End Sub

Private Sub LV_LostFocus()
Text1 = ""
Text2 = ""
End Sub
