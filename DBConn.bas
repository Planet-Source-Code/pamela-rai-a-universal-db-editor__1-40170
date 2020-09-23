Attribute VB_Name = "DBConn"
Public SQLstatement As String
Public conn As ADODB.Connection
Public rs As ADODB.Recordset
Public connObj As MSDASC.DataLinks
Public StrConnection As String


Public Function createConnection()
Set connObj = New MSDASC.DataLinks
StrConnection = connObj.PromptNew
End Function

Public Function openConnection()
Set conn = New ADODB.Connection
    conn.ConnectionString = StrConnection
conn.Open
'Debug.Print conn.ConnectionString
End Function

Public Function closeConnection()
conn.Close
End Function

Public Function getTables(xlist As ListBox)
xlist.Clear
Set rs = conn.OpenSchema(adSchemaTables)
    Do While Not rs.EOF
        res = InStr(1, rs!TABLE_NAME, "MSys", vbTextCompare)
        If Not res = 1 Then
        xlist.AddItem rs!TABLE_NAME
        End If
        rs.MoveNext
    Loop
End Function

Public Function executeSQL()

Set rs = conn.Execute(SQLstatement, , adCmdText)

End Function

Public Function setfieldNames(grid As MSFlexGrid)
For I = 0 To rs.Fields.Count - 1

grid.Cols = rs.Fields.Count
grid.Row = 0
grid.Col = I
grid = rs.Fields(I).Name
'grid.ColWidth = TextWidth(grid)
'Debug.Print rs.Fields(i).Name
Next
End Function
