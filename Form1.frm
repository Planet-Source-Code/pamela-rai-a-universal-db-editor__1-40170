VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "RG - Universal DB Editor"
   ClientHeight    =   4920
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8070
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSQL 
      Caption         =   "Run"
      Height          =   315
      Left            =   7440
      TabIndex        =   5
      Top             =   4590
      Width           =   600
   End
   Begin VB.TextBox txtSQL 
      Height          =   300
      Left            =   1620
      TabIndex        =   4
      Text            =   "Enter sql command here..."
      Top             =   4605
      Width           =   5805
   End
   Begin VB.ListBox lstTables 
      Height          =   4350
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1605
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4365
      Left            =   1605
      TabIndex        =   0
      Top             =   240
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   7699
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   315
      Left            =   7410
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   255
      ScaleWidth      =   270
      TabIndex        =   2
      Top             =   855
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblSQL 
      Caption         =   "SQLcommand:"
      Height          =   240
      Left            =   255
      TabIndex        =   3
      Top             =   4665
      Width           =   1125
   End
   Begin VB.Menu Database 
      Caption         =   "Database"
      Begin VB.Menu Connect 
         Caption         =   "Connect"
      End
      Begin VB.Menu sp 
         Caption         =   "-"
      End
      Begin VB.Menu MNUrecent 
         Caption         =   "recent0"
         Index           =   0
      End
      Begin VB.Menu MNUrecent 
         Caption         =   "recent1"
         Index           =   1
      End
      Begin VB.Menu MNUrecent 
         Caption         =   "recent2"
         Index           =   2
      End
      Begin VB.Menu MNUrecent 
         Caption         =   "recent3"
         Index           =   3
      End
   End
   Begin VB.Menu SQLcommands 
      Caption         =   "SQLcommands"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbx As DAO.Database
Dim tdx As DAO.TableDef
Dim fdx As DAO.Field
Dim rsx As DAO.Recordset
Private Sub Form_Load()
updaterecent
    hMenu& = GetMenu(Form1.hWnd)
    hSubMenu& = GetSubMenu(hMenu&, 0)
    hID& = GetMenuItemID(hSubMenu&, 0)
    SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, _
    Picture1.Picture, _
    Picture1.Picture


End Sub
Private Sub Connect_Click()
Dim saveit As Boolean
createConnection
openConnection
Set dbx = OpenDatabase(App.Path & "\recent.mdb")
Set rsx = dbx.OpenRecordset("SELECT * FROM resentconnection")
teller = 0
rsx.MoveFirst
Do While Not rsx.EOF


saveit = True
If rsx.Fields!ConnString = StrConnection Then
saveit = False
Exit Do
End If
rsx.MoveNext
teller = teller + 1
Loop
'''''''''''''''
'get DBQ


If saveit = True Then
resx = InStr(1, StrConnection, ":\", vbTextCompare)
resxx = InStr(resx + 1, StrConnection, ".", vbTextCompare)
resxx = resxx + 4
DBQ = Mid(StrConnection, resx - 1, (resxx + 1) - (resx))

Debug.Print DBQ
Debug.Print StrConnection

rsx.MoveLast
rsx.MovePrevious
holder = rsx.Fields!ConnString
holder1 = rsx.Fields!Path
rsx.MoveLast
rsx.Edit
rsx.Fields!Path = holder1
rsx.Fields!ConnString = holder
rsx.Update

rsx.MoveLast
rsx.MovePrevious
rsx.MovePrevious
holder = rsx.Fields!ConnString
holder1 = rsx.Fields!Path
rsx.MoveNext
rsx.Edit
rsx.Fields!Path = holder1
rsx.Fields!ConnString = holder
rsx.Update

rsx.MoveLast
rsx.MovePrevious
rsx.MovePrevious
rsx.MovePrevious
holder = rsx.Fields!ConnString
holder1 = rsx.Fields!Path
rsx.MoveNext

rsx.Edit
rsx.Fields!Path = holder1
rsx.Fields!ConnString = holder
rsx.Update

rsx.MoveFirst
rsx.Edit
rsx.Fields!Path = DBQ
rsx.Fields!ConnString = StrConnection
rsx.Update
'''''''''''
End If
dbx.Close
getTables lstTables
updaterecent
closeConnection
End Sub
Private Sub lstTables_Click()
openConnection
SQLstatement = "select * from [" & lstTables.List(lstTables.ListIndex) & "]"
executeSQL
setfieldNames MSF
closeConnection
AutosizeGridColumns MSF, 100, 2000, Form1
End Sub




Public Sub updaterecent()
Set dbx = OpenDatabase(App.Path & "\recent.mdb")
Set rsx = dbx.OpenRecordset("SELECT * FROM resentconnection")
teller = 0
rsx.MoveFirst
Do While Not rsx.EOF

Form1.MNUrecent(teller).Caption = rsx.Fields!Path
If rsx.Fields!Path = "x" Then
Form1.MNUrecent(teller).Visible = False
Else
Form1.MNUrecent(teller).Visible = True
End If
rsx.MoveNext
teller = teller + 1
Loop
dbx.Close
End Sub

Private Sub MNUrecent_Click(Index As Integer)
Set dbx = OpenDatabase(App.Path & "\recent.mdb")
Set rsx = dbx.OpenRecordset("SELECT * FROM resentconnection")
teller = 0
rsx.MoveFirst
Do While Not rsx.EOF

If rsx.Fields!Path = MNUrecent(Index).Caption Then
StrConnection = rsx.Fields!ConnString
End If
rsx.MoveNext
Loop
openConnection
getTables lstTables
closeConnection
End Sub



