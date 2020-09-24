VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''This class module will handle
''''Access to any db operations
''''class module properties
''''=======================
''''1. error code = 300000

Option Explicit

''write code for clsSuper
Implements clsSuper

Dim dbAccess As ADODB.Connection
Dim rsAccess As ADODB.Recordset

Dim catAccess As ADOX.Catalog
Dim tblAccess As ADOX.Table

Dim strCon As String            ''Connection string
Dim strSQL As String            ''SQL string

Dim lngTotalRecords As Long     ''Store Total no. of Records

''Dim intPrevDestDatabaseType As Integer      ''Identify previsouly created db

Private Sub Class_Initialize()

On Error GoTo ehClassInit
    
    ErrorCode = 0
    
    Set dbAccess = New ADODB.Connection
    Set rsAccess = New ADODB.Recordset

    ''intPrevDestDatabaseType = 0
    
PROC_EXIT:
    Exit Sub

ehClassInit:
    ErrorCode = 300000
End Sub

Private Sub Class_Terminate()

On Error Resume Next

If rsAccess.State = 1 Then rsAccess.Close
If dbAccess.State = 1 Then dbAccess.Close

Set rsAccess = Nothing
Set dbAccess = Nothing

End Sub

Public Function clsSuper_GetTables( _
                Optional ByVal strSourceDatabaseName As String, _
                Optional intSourceDatabaseType As Integer, _
                Optional ByVal strSourceDatabaseFile As String, _
                Optional ByVal strSourceDatabaseFilePath As String, _
                Optional ByVal strSourceDatabaseUserName As String, _
                Optional ByVal strSourceDatabasePassword As String, _
                Optional ByVal strSourceDatabaseServerName As String) As String()

On Error GoTo ehGetTables

Dim strResult() As String
Dim i As Integer

''reset error code
ErrorCode = 0

If dbAccess Is Nothing Then Set dbAccess = New ADODB.Connection
''close the database
If dbAccess.State <> 0 Then dbAccess.Close

''strCon = ""

strCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
         "Data Source=" & strSourceDatabaseFilePath & strSourceDatabaseFile & ";" & _
         "Persist Security Info=False"
        
dbAccess.Open strCon


''retreive the tables from the database

Set catAccess = New ADOX.Catalog
Set tblAccess = New ADOX.Table

catAccess.ActiveConnection = strCon

i = 0

''determine the number of tables
''in this database so as to redim
''the strResult() variable

For Each tblAccess In catAccess.Tables
    If tblAccess.Type = "TABLE" Then
        i = i + 1
    End If
Next

ReDim strResult(1 To i)

i = 1

For Each tblAccess In catAccess.Tables
    If tblAccess.Type = "TABLE" Then
        strResult(i) = tblAccess.Name
        i = i + 1
    End If
Next

''check for ubound
If UBound(strResult) > 0 Then clsSuper_GetTables = strResult()

PROC_EXIT:
    Set catAccess = Nothing
    Set tblAccess = Nothing
    Exit Function

ehGetTables:
    ErrorCode = 300100
End Function

Public Function clsSuper_GetFields( _
            Optional ByVal strSourceDatabaseName As String, _
            Optional intSourceDatabaseType As Integer, _
            Optional ByVal strSourceDatabaseFile As String, _
            Optional ByVal strSourceDatabaseFilePath As String, _
            Optional ByVal strSourceDatabaseTableName As String, _
            Optional ByVal strSourceDatabaseUserName As String, _
            Optional ByVal strSourceDatabasePassword As String, _
            Optional ByVal strSourceDatabaseServerName As String) As ADODB.Recordset
            
''return available fields
On Error GoTo ehGetFields

''reset error code
ErrorCode = 0

If dbAccess Is Nothing Then
    Set dbAccess = New ADODB.Connection
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & strSourceDatabaseFilePath & strSourceDatabaseFile & ";" & _
             "Persist Security Info=False"
    dbAccess.Open strCon
End If

If rsAccess Is Nothing Then Set rsAccess = New ADODB.Recordset
''close existing recordset
If rsAccess.State <> 0 Then rsAccess.Close

rsAccess.Open "SELECT * FROM " & strSourceDatabaseTableName, dbAccess, adOpenStatic, adLockReadOnly

If rsAccess.RecordCount > 0 Then
    rsAccess.MoveLast
    lngTotalRecords = rsAccess.RecordCount
Else
    lngTotalRecords = 0
End If
        
''return rsAccess recordset
Set clsSuper_GetFields = rsAccess

PROC_EXIT:
    Exit Function

ehGetFields:
    ErrorCode = 300200
End Function

Public Function clsSuper_GetTotalRecords() As Long
''return total number of records
clsSuper_GetTotalRecords = lngTotalRecords
End Function

Public Function clsSuper_Convert(strFields() As String, _
                        Optional ByVal strSourceDatabaseName As String, _
                        Optional intSourceDatabaseType As Integer, _
                        Optional ByVal strSourceDatabaseFile As String, _
                        Optional ByVal strSourceDatabaseFilePath As String, _
                        Optional ByVal strSourceDatabaseTableName As String, _
                        Optional ByVal strSourceDatabaseUserName As String, _
                        Optional ByVal strSourceDatabasePassword As String, _
                        Optional ByVal strSourceDatabaseServerName As String, _
                        Optional ByVal strDestDatabaseName As String, _
                        Optional ByVal intDestDatabaseType As Integer, _
                        Optional ByVal strDestDatabaseFile As String, _
                        Optional ByVal strDestDatabaseFilePath As String, _
                        Optional ByVal strDestDatabaseUserName As String, _
                        Optional ByVal strDestDatabasePassword As String, _
                        Optional ByVal strDestDatabaseServerName As String) As Boolean
                        
On Error GoTo ehConvert

''reset error code
ErrorCode = 0
                        
strSQL = ""
                                          
Select Case intDestDatabaseType

''=====================================================================
    Case Is = 1     ''Access Output
    
        ''Create a blank mdb file now (dest database)
        ''concat filepath and filename
        CreateEmptyMDB strDestDatabaseFilePath & strDestDatabaseFile
        
        ''Prepare sql
        strSQL = "SELECT " & ParseStringArray(strFields) & vbCrLf & _
                 "INTO [" & strSourceDatabaseTableName & "]" & vbCrLf & _
                 "IN '" & strDestDatabaseFilePath & strDestDatabaseFile & "'" & vbCrLf & _
                 "FROM [" & strSourceDatabaseTableName & "]"
        
        ''strSQL = strSQL & "FROM [" & strSourceDatabaseTableName & "]"
        
''=====================================================================
 
     Case Is = 11, 12, 13, 21, 22, 23, 24, 25, 31, 41, 42, 43, 51 ''All other formats
        
        ''===========================================================
        ''if source is access replace strsourcedatabasefile with
        ''strsourcedatabasetablename
        ''===========================================================
        
        strSourceDatabaseFile = strSourceDatabaseTableName
        
        Select Case intDestDatabaseType
        
            Case Is = 11, 12, 13, 41, 42, 43, 51 'dBASE, Paradox & Text output
            
                strSQL = "SELECT " & ParseStringArray(strFields) & vbCrLf & _
                        "INTO [" & GetDatabaseType(intDestDatabaseType) & ";Database=" & strDestDatabaseFilePath & "]." & _
                        "[" & strDestDatabaseFile & "]" & vbCrLf & _
                        "FROM [" & strSourceDatabaseFile & "]"
                        
            Case Is = 21, 22, 23, 24, 25   ''Excel output
                
                strDestDatabaseFile = Mid(strDestDatabaseFile, 1, InStr(1, strDestDatabaseFile, ".") - 1)
                
                strSQL = "SELECT " & ParseStringArray(strFields) & vbCrLf & _
                        "INTO [" & GetDatabaseType(intDestDatabaseType) & ";Database=" & strDestDatabaseFilePath & strDestDatabaseFile & "]." & _
                        "[" & strDestDatabaseFile & "]" & vbCrLf & _
                        "FROM [" & strSourceDatabaseFile & "]"
            
            Case Is = 31    'HTML output
            
                strDestDatabaseFile = Mid(strDestDatabaseFile, 1, InStr(1, strDestDatabaseFile, ".") - 1)
                strDestDatabaseFile = strDestDatabaseFile & ".htm"
                    
                strSQL = "SELECT " & ParseStringArray(strFields) & vbCrLf & _
                        "INTO [" & GetDatabaseType(intDestDatabaseType) & ";Database=" & strDestDatabaseFilePath & "]." & _
                        "[" & strDestDatabaseFile & "]" & vbCrLf & _
                        "FROM [" & strSourceDatabaseFile & "]"
        End Select
''=====================================================================
End Select

If Not dbAccess.State = 0 And strSQL <> "" Then
    dbAccess.Execute strSQL
    
    ''====================================================
    ''remember last used dest database type (no need here)
    ''====================================================
    ''intPrevDestDatabaseType = intDestDatabaseType
    clsSuper_Convert = True
    GoTo PROC_EXIT
End If

PROC_EXIT:
    Exit Function

ehConvert:
    ErrorCode = 300300
    clsSuper_Convert = False
End Function

Public Sub CreateEmptyMDB(strMdbFileName As String)
''create a empty access database file

On Error GoTo ehCreateEmptyMDB

If Dir(strMdbFileName) = "" Then
    
    Set catAccess = New ADOX.Catalog
    
    catAccess.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & strMdbFileName
                  
    Set catAccess = Nothing
    
End If

PROC_EXIT:
    Exit Sub
    
ehCreateEmptyMDB:
    ErrorCode = 300400
End Sub
