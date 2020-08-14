Option Explicit

Public colLabelEvent As Collection 'Collection of labels for event handling
Public colLabels As Collection     'Collection of the date labels
Public bSecondDate As Boolean      'True if finding second date
Public sActiveDay As String        'Last day selected
Public lDays As Long               'Number of days in month
Public lFirstDay As Long           'Day selected, e.g. 19th
Public lStartPos As Long
Public lSelMonth As Long           'The selected month
Public lSelYear As Long            'The selected year
Public lSelMonth1 As Long          'Used to check if same date is selected twice
Public lSelYear1 As Long           'Used to check if same date is selected twice

Public cnn As New Connection

Public rs As New Recordset
Public rs2 As New Recordset

Public strSQL As String

Public lastID As String

Public g_objFSO As Scripting.FileSystemObject
Public g_scrText As Scripting.TextStream

Public Sub OpenDB()
    If cnn.State = adStateOpen Then cnn.Close
    On Error GoTo Handle
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & _
        ThisWorkbook.Sheets("L").Range("DB_PATH")
        cnn.Open
    Exit Sub
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, "Mulo2", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Public Sub closeRS()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If cnn.State = adStateOpen Then cnn.Close
    Set cnn = Nothing
End Sub
    
Public Function LogFile_WriteError(ByVal sRoutineName As String, _
                                    ByVal sMessage As String)
    Dim logPath As String
    
    logPath = ThisWorksheet.Sheets("L").Range("LOG_PATH")
    
    Dim sText As String
    On Error GoTo ErrorHandler
    If (g_objFSO Is Nothing) Then
        Set g_objFSO = New FileSystemObject
    End If
    If (g_scrText Is Nothing) Then
        If (g_objFSO.FileExists(logPath) = False) Then
            Set g_scrText = g_objFSO.OpenTextFile(logPath, IOMode.ForWriting, True)
        Else
            Set g_scrText = g_objFSO.OpenTextFile(logPath, IOMode.ForAppending)
        End If
    End If
    sText = sText & Format(Date, "DD/MM/YYYY") & " " & Time() & "|"
    sText = sText & sRoutineName & "|"
    sText = sText & sMessage & "|"
    g_scrText.WriteLine sText
    g_scrText.Close
    Set g_scrText = Nothing
    Exit Function
ErrorHandler:
    Set g_scrText = Nothing
    Call MsgBox("No se pudo escribir en el fichero log", vbCritical, "LogFile_WriteError")
End Function

Public Sub Error_Handle(ByVal sRoutineName As String, _
                         ByVal sObject As String, _
                         ByVal currentStrSQL As String, _
                         ByVal sErrorNo As String, _
                         ByVal sErrorDescription As String)
    Dim sMessage As String
    sMessage = sObject & "|" & currentStrSQL & "|" & sErrorNo & "|" & sErrorDescription & "|" & Application.UserName
    Call MsgBox(sErrorNo & vbCrLf & sErrorDescription, vbCritical, sRoutineName & " - " & sObject & " - Error")
    Call LogFile_WriteError(sRoutineName, sMessage)
End Sub


