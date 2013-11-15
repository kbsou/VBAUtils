VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xlQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************************************** Class xlQuery **********************************************+
'                                                                                                   '
'   Version 0.1                                                                                     '
'   Contains methods and functions for using SQL Queries on diverse data sources                    '
'                                                                                                   '
'======================================== License ==================================================+
'                                                                                                   '
'   This piece of code is licensed MIT License                                                      '
'   http://opensource.org/licenses/MIT                                                              '
' Contributors:                                                                                     '
'   kbsou - https://github.com/kbsou - xlQuery v0.1                                                 '
'                                                                                                   '
'======================================== IMPORTANT ================================================+
' This class references: (Added automatically)                                                      '
'       Microsoft ActiveX Data Objects 2.6 or higher                                                '
'       Microsoft Scripting Runtime                                                                 '
'===================================================================================================+

Option Explicit
Option Base 1

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H1

Private isConnected As Boolean

Dim con As Object
Dim rs As Object

'Separator for csv files
Public Separator As String

Private Sub Class_Initialize()
    
    require "adodb"
    require "scripting"
    init
    
End Sub

Private Sub init()

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    isConnected = False
    Separator = "," 'Default separator for csv files
    
End Sub

Public Sub Connect(path As String, Optional fileType As String = "", Optional password = "")
'* Sets up a connection to the data source
 ' Parameters:
  ' path: Path to the file or folder in case of text files
  ' filetype: optional - used in case of text files to specify the type (csv or fixed lenght)
  ' password: optional - used if the database requires a password to connect
  ' When the files are in csv or fixed lenght format you must pass the directory path instead
  
  
    con.Open CreateConnectionString(path, fileType, password)
    isConnected = True
    
End Sub

Private Function CreateConnectionString(path As String, Optional fileType As String = "", Optional password = "") As String
'Creates a connection string based on the parameters
'For more info on connection strings visit www.connectionstrings.com

    Dim conStr As String
    
    fileType = LCase(fileType)
    
    ' Tries to infer the filetype by the extension
    If path Like "*.xls?" Then
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                 "Data Source='" & path & "';" & _
                 "Extended Properties=""Excel 12.0;" & _
                 "HDR=YES;IMEX=1;"";"
    
    ElseIf path Like "*.xls" Then
        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                 "Data Source='" & path & "';" & _
                 "Extended Properties=""Excel 8.0;" & _
                 "HDR=Yes;IMEX=1;"";"
    
    ElseIf path Like "*.mdb" Then
        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                 "Data Source='" & path & "';"
        If password = "" Then
            conStr = conStr & "User Id=admin;Password=;"
        Else
            conStr = conStr & "Jet OLEDB:Database Password=" & password & ";"
        End If
    
    ElseIf path Like "*.accdb" Then
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                 "Data Source='" & path & "';"
        If password = "" Then
            conStr = conStr & "User Id=admin;Password=;"
        Else
            conStr = conStr & "Jet OLEDB:Database Password=" & password & ";"
        End If
        
    ElseIf fileType = "csv" Then
        If Separator = "" Then Err.Raise 589, "CSV separator not found"
        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & replace(path, "\", "\\") & ";" & _
                "Extended Properties=""text;" & _
                "HDR=Yes;FMT=Delimited(" & Separator & ");IMEX=1;"";"
                
    ElseIf fileType = "fixed" Then
        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                 "Data Source='" & replace(path, "\", "\\") & "';" & _
                 "Extended Properties=""text;HDR=No;FMT=Fixed"";"
    Else
        Err.Raise vbObjectError + 599, "xlQuery.Connect", _
              "File type missing or not supported: " & fileType
    End If

    CreateConnectionString = conStr
    
End Function


Public Function OpenRecordset(query As String) As Object
'* Returns a recordset based to a SQL query
' Parameters:
 ' query: string with the SQL Query
 ' WARNING: Risk of SQL injection. Be careful with this parameter

    'You must configure the connection with the method connect first
    If Not isConnected Then
        Err.Raise vbObjectError + 599, "xlQuery.OpenRecordset", _
              "Connection not configured"
    End If
    
    'If the recordset is open closes it
    If rs.State = adStateOpen Then rs.Close
        
    rs.Open query, con, adOpenStatic, adLockOptimistic, adCmdText
    Set OpenRecordset = rs
    
End Function

Public Sub CreateCsvSchema(files As Variant, fields As Variant, types As Variant, _
                            Optional delimiter As String = ",")
    
    If UBound(fields) <> UBound(types) Then _
        Err.Raise vbObjectError + 589, "xlQuery.CreateCsvSchema", _
              "The fields array and the types array must have the same size"
    
    Separator = delimiter
    Call CreateSchema(files, "Delimited(" & delimiter & ")", fields, types)
    
End Sub

Public Sub CreateFixedLenghtSchema(files As Variant, fields As Variant, types As Variant, _
                                    sizes As Variant)
    
    If UBound(fields) <> UBound(types) Or UBound(fields) <> UBound(sizes) Then _
        Raise vbObjectError + 589, "xlQuery.CreateFixedLenghtSchema", _
              "The arrays fields, types and sizes must have the same size"
    
    Call CreateSchema(files, "fixed", fields, types, sizes)
    
End Sub

'* Sub CreateSchema:
' Creates a table definition in schema.ini file in the directory of files()
' Parameters:
'   files() : Variant Array with full path of the files to be added to the table definitions
'             of the schema.ini file
'             The files must be on the same directory.
'   filetype: Either "delimited([Separator])", for delimited files (like csv) or "Fixed" for
'               fixed lenght files.
'   fields(): Variant Array with the fields.
'   types() : Variant Array with the type of the fields. See [1] for info about types
'   sizes() : Variant Array Size of the fields. Only required for Fixed lenght files
' More info on schema files:
'  [1] http://msdn.microsoft.com/en-us/library/ms709353%28VS.85%29.aspx
'***********************************************************************************************
'                   WARNING: Previous schema.ini files will be overwritten                     *
'***********************************************************************************************

Private Sub CreateSchema(files As Variant, fileType As String, fields As Variant, _
                        types As Variant, sizes As Variant)

    Dim fso As Scripting.FileSystemObject, _
        fschema As Scripting.TextStream
    Dim dirpath As String, _
        fieldDefs As String, _
        i As Integer
    
    dirpath = fso.GetAbsolutePathName(files(LBound(files)))
    
    'Warns if the file already exists
    If fso.FileExists(dirpath & "\schema.ini") Then Debug.Print dirpath & "\schema.ini will be overwritten"
        
    Set fschema = fso.OpenTextFile(dirpath & "\schema.ini", ForAppending, Create:=True)
    
    'Creates a string with the field definitions
    For i = 1 To UBound(fields)
        fieldDefs = fieldDefs & vbCrLf & _
                    "col" & i & ":" & fields(i) & " " & types(i)
        If fileType = "fixed" Then fieldDefs = fieldDefs & "Width " & sizes(i)
    Next i
    
    For i = 1 To UBound(files)
        Dim fname As String
        fname = fso.GetFileName(files(i))
        
        fschema.WriteLine "[" & fname & "]"
        fschema.WriteLine "Format=" & fileType
        fschema.Write fieldDefs & vbCrLf & vbCrLf
    Next i
    
End Sub

Private Sub require(reference As String, Optional path As String = "", Optional guid As String = "", Optional major As Long = 1, Optional minor As Long = 0)
'Sets up a reference based on the parameters.

    Dim ref As Object, _
        app As Object
    
    If Application.Name Like "*Excel*" Then Set app = ThisWorkbook.VBProject
    If Application.Name Like "*Access*" Then Set app = Application
    
    reference = LCase(reference)
    
    'Verifies if the reference already exists
    For Each ref In app.References
        If LCase(ref.Name) Like "*" & reference & "*" Then Exit Sub
    Next ref
    
    If guid = "" And path = "" Then
        'Some defaults
        Select Case reference
        Case "regexp"
            guid = "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}": major = 5: minor = 5
        Case "scripting"
            guid = "{420B2830-E718-11CF-893D-00A0C9054228}": major = 1: minor = 0
        Case "adodb"
            guid = "{B691E011-1797-432E-907A-4D8C69339129}": major = 2: minor = 6
        Case "vbide"
            guid = "{0002E157-0000-0000-C000-000000000046}": major = 5: minor = 3
        Case Else
            Err.Raise _
                vbObjectError + 559, _
                "require", _
                "Require requires a GUID or a path to add the reference to " & reference
        End Select
    End If
      
    If guid <> "" Then
        Set ref = app.References.AddFromGuid(guid, major, minor)
    ElseIf path <> "" Then
        Set ref = app.References.AddFromFile(path)
    End If
    
    Debug.Print "Ref added: " & reference, ref.Name, ref.Description
    
End Sub

Public Sub CloseRecordset()
    On Error Resume Next
    rs.Close
End Sub

Public Sub CloseConnection()
    On Error Resume Next
    rs.Close
    con.Close
    isConnected = False
End Sub

Private Sub Class_Terminate()
    
    On Error Resume Next
    rs.Close
    con.Close
    Set rs = Nothing
    Set con = Nothing
    
End Sub


