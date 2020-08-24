Attribute VB_Name = "mMain"
Option Explicit

Sub Main()

    ' - - - getting connection string for Excel 2013 from https://www.connectionstrings.com/ace-oledb-12-0/
    Dim wbPath          As String:          wbPath = ThisWorkbook.FullName
    Dim connStr         As String:          connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & wbPath & _
                                                                    ";Extended Properties=""Excel 12.0;HDR=YES"";"
    
    ' - - - data source file where script sends select query
    Dim sourceFile      As String:          sourceFile = ThisWorkbook.Path & "\" & "data.xlsx"
    
    
    ' - - - database source string
    Dim sourceDb        As String:          sourceDb = "[Excel 12.0;HDR=YES;DATABASE=" & sourceFile & "]"
    
    
    ' - - - worksheet of data source file
    Dim sourceWs        As String:          sourceWs = "[Sheet1$]"
    
    
    ' - - - string concatination of database source string and worksheet of data source file
    Dim sourceStr       As String:          sourceStr = sourceDb & "." & sourceWs
    
    
    ' - - - query SELECT with source string appended
    Dim qry             As String:          qry = "SELECT * FROM " & sourceStr
    
    ' - - - setting connection
    Dim conn            As Object:          Set conn = CreateObject("ADODB.Connection")
    Dim rs              As Object:          Set rs = CreateObject("ADODB.Recordset")
    Dim i               As Long
    
    
    ' - - - getting connected to db
    
    On Error GoTo ConnectionClose

    conn.Open connStr
    rs.Open qry, conn
    
    On Error GoTo 0
    
    ' - - - clearing sheet once connected
    
    Sheet1.Cells.ClearContents
    
    ' - - - getting headers
    
    For i = 0 To rs.Fields.Count - 1
        Sheet1.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i
    
    ' - - - getting db records
    
    If rs.EOF Then
        GoTo ConnectionClose
    Else
        Sheet1.Range("A2").CopyFromRecordset rs
    End If
    
ConnectionClose:

    Set rs = Nothing
    conn.Close

End Sub

