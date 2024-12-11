Attribute VB_Name = "Module1"
'declaration zone

Public conn As ADODB.Connection
Public RcrdSet As ADODB.Recordset
Public sql As String    'container for SQL statement

Public ConnStr As String  'container for connection string
'example connString = "Provider=MSDAORA.1;User ID=PRJ2432N/PRJ2432N;Persist Security Info=False"

Public filePath As String 'container for file/directory Address

'end of declaration zone

'creating a function for backend Integration
Public Sub bckend(connString As String, sql As String)

    'Error raise if Error Occurs
    On Error GoTo bckend_errHndlZone
    
    'connection creation and open
    Set conn = New ADODB.Connection
    conn.Open connString
    
    'recordset creation
    Set RcrdSet = New ADODB.Recordset
    
    'Execution of SQL Statement
    conn.Execute sql
    
    'Informing User about the Progress
    MsgBox "SQL executed successfully!"
    
    'Closing the Connection
    conn.Close
    
    'free the space occupied by container C
    Set conn = Nothing
    
    'Everything ran smoothly, Success!
    Exit Sub
    
    'If runtime error occurs in any part of bckend sub
    'the compiler will jump directly to the below part of code
bckend_errHndlZone:

    'vbcrlf means Visual Basic Carriage Return Line Feed
    'chr(13)Carriage Return
    'chr(10) Line Feed
    '(_) is for continuos vb statement with a line break
    
    'Informing User about the Runtime Error
    MsgBox ("Something Unexpected has Occured! " & vbCrLf & _
    "Error Number: " & Err.Number & vbCrLf & _
    "Error Description: " & Err.Description & Chr(13) & Chr(10) & _
    "Error Source: " & Err.Source)
    
    'clear the Error
    Err.Clear
    
    'Error Handled Successfully, Application will not Terminate!!!
End Sub

'end of backend integration sub
