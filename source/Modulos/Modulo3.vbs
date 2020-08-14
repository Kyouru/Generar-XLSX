Sub actualizarTodo()

    ActualizarCargas
    actualizarPendientes
    
    Hoja1.Range("INTERVALO").NumberFormat = "@"
    
End Sub

Sub AbrirXLSX()
    Dim id_cuenta As Integer
    Dim strDir As String
    Dim correlativo As String
    Dim tmpSplit As Variant
    
    If ActiveSheet.Name = Hoja1.Name Then
            
        If Selection.Column >= 2 And Selection.Column <= 4 And Selection.Row >= 4 And Range("B" & Selection(1, 1).Row).Value <> "" Then
            strDir = ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format([FECHA_HISTORICO], "YYYY") & "\" & Format([FECHA_HISTORICO], "MM") & "\" & Format([FECHA_HISTORICO], "YYYY.MM.DD") & "\" & Range("B" & Selection(1, 1).Row).Value & ".xlsx"
            If Dir(strDir) <> "" Then
                Workbooks.Open (strDir)
            Else
                tmpSplit = Split(Range("B" & Selection(1, 1).Row).Value, " ")
                If tmpSplit(0) = "20BCP" Then
                    strDir = ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format([FECHA_HISTORICO], "YYYY") & "\" & Format([FECHA_HISTORICO], "MM") & "\" & Format([FECHA_HISTORICO], "YYYY.MM.DD") & "\BCP20 " & tmpSplit(1) & " " & tmpSplit(2) & ".xlsx"
                    If Dir(strDir) <> "" Then
                        Workbooks.Open (strDir)
                    Else
                        MsgBox "Error: No se encontr・el XLSX" & vbCrLf & " >> Ruta: " & strDir
                    End If
                Else
                    MsgBox "Error: No se encontr・el XLSX" & vbCrLf & " >> Ruta: " & strDir
                End If
            End If
        Else
            If Selection.Column >= 6 And Selection.Column <= 8 And Selection.Row >= 6 And Range("F" & Selection(1, 1).Row).Value <> "" Then
                strSQL = "SELECT ID_PARAM_CUENTA, MAX(CORRELATIVO) FROM PARAM_CUENTA LEFT JOIN CARGA_XLSX ON CARGA_XLSX.ID_PARAM_CUENTA_FK = PARAM_CUENTA.ID_PARAM_CUENTA WHERE NOMBRE_CORTO+' '+NUMERO_CORTO = '" & Range("F" & Selection(1, 1).Row).Value & "' AND FORMAT(FECHA,'YYYY-MM-DD') = '" & Format([FECHA_HISTORICO], "YYYY-MM-DD") & "' GROUP BY ID_PARAM_CUENTA"
                
                OpenDB
                On Error Resume Next
                rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
                On Error GoTo 0
                    
                'Log del Error
                If cnn.Errors.count > 0 Then
                    Call Error_Handle(cnn.Errors.Item(0).Source, "Modulo3 - AbrirXLSX", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
                    cnn.Errors.Clear
                    closeRS
                    Exit Sub
                End If
                
                If rs.RecordCount > 0 Then
                    id_cuenta = rs.Fields(0).Value
                    correlativo = rs.Fields(1).Value
                Else
                    correlativo = 1
                End If
                
                closeRS
                
                If Len(correlativo) = 1 Then
                    correlativo = "0" & correlativo
                End If
                
                strDir = ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format([FECHA_HISTORICO], "YYYY") & "\" & Format([FECHA_HISTORICO], "MM") & "\" & Format([FECHA_HISTORICO], "YYYY.MM.DD") & "\" & Range("F" & Selection(1, 1).Row).Value & " " & correlativo & ".xlsx"
                If Dir(strDir) <> "" Then
                    Workbooks.Open (strDir)
                Else
                    tmpSplit = Split(Range("F" & Selection(1, 1).Row).Value, " ")
                    If tmpSplit(0) = "20BCP" Then
                        strDir = ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format([FECHA_HISTORICO], "YYYY") & "\" & Format([FECHA_HISTORICO], "MM") & "\" & Format([FECHA_HISTORICO], "YYYY.MM.DD") & "\BCP20 " & tmpSplit(1) & " " & correlativo & ".xlsx"
                        If Dir(strDir) <> "" Then
                            Workbooks.Open (strDir)
                        Else
                            MsgBox "Error: No se encontr・el XLSX" & vbCrLf & " >> Ruta: " & strDir
                        End If
                    Else
                        MsgBox "Error: No se encontr・el XLSX" & vbCrLf & " >> Ruta: " & strDir
                    End If
                    
                End If
            ElseIf Selection.Column >= 12 And Selection.Column <= 14 And Selection.Row >= 10 And Range("K" & Selection(1, 1).Row).Value <> "" Then
                strDir = ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format([FECHA_HISTORICO], "YYYY") & "\" & Format([FECHA_HISTORICO], "MM") & "\" & Format([FECHA_HISTORICO], "YYYY.MM.DD") & "\" & Range("L" & Selection(1, 1).Row).Value & ".xlsx"
                If Dir(strDir) <> "" Then
                    Workbooks.Open (strDir)
                End If
            End If
        End If
        
    End If
End Sub


Sub ActualizarCargas()

    'Remover Temporizador
    On Error Resume Next
    Application.OnTime Hoja1.proxActualizacion, Procedure:="ActualizarCargas", Schedule:=False
    On Error GoTo 0
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    ThisWorkbook.Activate
    ThisWorkbook.Sheets(1).Activate
    
    'Consulta Historica por fecha
    
    strSQL = "SELECT NOMBRE_CORTO + ' ' + NUMERO_CORTO + ' ' + IIF(CORRELATIVO <= 9, '0' + CSTR(CORRELATIVO), CSTR(CORRELATIVO)), FECHA, IIF(ALIAS.NOMBRE_ALIAS IS NOT NULL, ALIAS.NOMBRE_ALIAS, CARGA_XLSX.USUARIO) FROM ((CARGA_XLSX LEFT JOIN ALIAS ON ALIAS.NOMBRE = CARGA_XLSX.USUARIO) LEFT JOIN PARAM_CUENTA ON PARAM_CUENTA.ID_PARAM_CUENTA = CARGA_XLSX.ID_PARAM_CUENTA_FK) WHERE FORMAT(FECHA,'YYYY-MM-DD') = '" & Format([FECHA_HISTORICO], "YYYY-MM-DD") & "' ORDER BY FECHA, FECHA_GEN"
    
    Range(Range("dataSet"), Range("dataSet").End(xlDown)).ClearContents
    
    OpenDB
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, " Modulo3 - ActualizarCargas (1)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    If rs.RecordCount > 0 Then
        Range("dataSet").CopyFromRecordset rs
        Range(Split(Range("dataSet")(1, 2).Address, "$")(1) & ":" & Split(Range("dataSet")(1, 2).Address, "$")(1)).NumberFormat = "[$-F400]h:mm:ss AM/PM"
        Range("FECHA_HISTORICO").NumberFormat = "DD/MM/YYYY"
    End If
    
    'Ultima Carga de la fecha
    strSQL = "SELECT NOMBRE_CORTO + ' ' + NUMERO_CORTO, C.FECHA, IIF(ALIAS.NOMBRE_ALIAS IS NOT NULL, ALIAS.NOMBRE_ALIAS, C.USUARIO) FROM (((SELECT ID_PARAM_CUENTA_FK, MAX(CORRELATIVO) AS MCORRELATIVO FROM CARGA_XLSX WHERE FORMAT(FECHA,'yyyy-mm-dd') = '" & Format([FECHA_HISTORICO], "YYYY-MM-DD") & "' GROUP BY ID_PARAM_CUENTA_FK) AS R LEFT JOIN (SELECT * FROM CARGA_XLSX WHERE '" & Format([FECHA_HISTORICO], "YYYY-MM-DD") & "' = FORMAT(FECHA,'yyyy-mm-dd')) AS C ON R.ID_PARAM_CUENTA_FK = C.ID_PARAM_CUENTA_FK AND R.MCORRELATIVO = C.CORRELATIVO) LEFT JOIN ALIAS ON ALIAS.NOMBRE = C.USUARIO) LEFT JOIN PARAM_CUENTA ON PARAM_CUENTA.ID_PARAM_CUENTA = C.ID_PARAM_CUENTA_FK" & _
    " WHERE NOMBRE_CORTO <> 'HIST'" & _
    " ORDER BY NOMBRE_CORTO + ' ' + NUMERO_CORTO"
    
    Set rs = Nothing
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, " Modulo3 - ActualizarCargas (2)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    Range(Range("ultimasCargas"), Range("ultimasCargas").End(xlDown)).ClearContents
    
    If rs.RecordCount > 0 Then
        Range("ultimasCargas").CopyFromRecordset rs
        Range(Split(Range("ultimasCargas")(1, 2).Address, "$")(1) & ":" & Split(Range("ultimasCargas")(1, 2).Address, "$")(1)).NumberFormat = "[$-F400]h:mm:ss AM/PM"
    End If
    
    Hoja1.verificarVersion
    
    If Hoja1.cbActualizar And IsNumeric([INTERVALO]) Then
        If [INTERVALO] >= 1 Then
            Call timerSeg
        End If
    End If
    
    [ULTIMA_ACTUALIZACION] = Now()
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    closeRS
End Sub

Sub timerSeg()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim invalido As Boolean
    invalido = False
    Hoja1.Range("INTERVALO").NumberFormat = "@"
    If Hoja1.cbActualizar And IsNumeric([INTERVALO]) Then
        If [INTERVALO] >= 1 And [INTERVALO] < 60 Then
            If [INTERVALO] = CInt([INTERVALO]) Then
                Hoja1.proxActualizacion = Now + TimeValue("00:" & [INTERVALO] & ":00")
                Application.OnTime Hoja1.proxActualizacion, Procedure:="actualizarTodo"
                [PROXIMA_ACTUALIZACION] = Hoja1.proxActualizacion
            Else
                invalido = True
            End If
        Else
            invalido = True
        End If
    Else
        invalido = True
    End If
    If invalido Then
        [PROXIMA_ACTUALIZACION] = ""
        On Error Resume Next
        Application.OnTime Hoja1.proxActualizacion, Procedure:="actualizarTodo", Schedule:=False
        On Error GoTo 0
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'Busca los ficheros pendientes de procesar por el encriptador
Sub actualizarPendientes()
    Dim file As Variant
    
    file = Dir(ThisWorkbook.Sheets("L").Range("INPUT_PATH"))
    strSQL = "UPDATE CARGA_XLSX SET EN_INPUT = FALSE, FECHA_SALIDA = #" & Format(Now(), "YYYY-MM-DD HH:MM:SS") & "# WHERE EN_INPUT = TRUE"
    
    OpenDB
    
    On Error Resume Next
    cnn.Execute strSQL
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, "actualizarPendientes", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    strSQL = "UPDATE CARGA_XLSX SET EN_INPUT = TRUE WHERE"
    
    Dim tmpSplit As Variant
    Dim tmpStr As String
    Dim InUse As Boolean
    Dim swPendiente As Boolean
    
    InUse = False
    swPendiente = False
    While (file <> "")
        
        tmpStr = ""
        tmpSplit = Split(file, ".")(0)
        
        While Right(tmpSplit, 1) <> "_" And Len(tmpSplit) > 1
            tmpStr = Right(tmpSplit, 1) & tmpStr
            tmpSplit = Left(tmpSplit, Len(tmpSplit) - 1)
        Wend
        
        If IsNumeric(tmpStr) Then
            swPendiente = True
            strSQL = strSQL & " ID_CARGA_XLSX = " & tmpStr & " OR"
        End If
        
        If FileAlreadyOpen(ThisWorkbook.Sheets("L").Range("INPUT_PATH") & CStr(file)) Then
            InUse = True
        End If
        file = Dir
    Wend
    
    If InUse Or Not swPendiente Then
        [EN_PROCESO] = "EN PROCESO"
    Else
        [EN_PROCESO] = "ERROR"
    End If
    
    If Right(strSQL, 3) = " OR" Then
        strSQL = Left(strSQL, Len(strSQL) - 3)
    End If
    
    If swPendiente Then
        
        On Error Resume Next
        cnn.Execute strSQL
        On Error GoTo 0
        
        'Log del Error
        If cnn.Errors.count > 0 Then
            Call Error_Handle(cnn.Errors.Item(0).Source, "actualizarPendientes (1)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
            cnn.Errors.Clear
            closeRS
            Exit Sub
        End If
    End If
    
    strSQL = "SELECT ID_CARGA_XLSX, NOMBRE_CORTO + ' ' + NUMERO_CORTO + ' ' + IIF(CORRELATIVO < 10, '0' + CSTR(CORRELATIVO), CSTR(CORRELATIVO)), FECHA_GEN, FECHA_SALIDA - FECHA_GEN FROM CARGA_XLSX LEFT JOIN PARAM_CUENTA ON PARAM_CUENTA.ID_PARAM_CUENTA = CARGA_XLSX.ID_PARAM_CUENTA_FK WHERE EN_INPUT = TRUE ORDER BY ID_CARGA_XLSX ASC"
    
    Set rs = Nothing
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, "actualizarPendientes (2)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    Hoja1.Range(Hoja1.Range("dataSetProceso"), Hoja1.Range("dataSetProceso").End(xlDown)).ClearContents
    
    If rs.RecordCount > 0 Then
        Hoja1.Range("dataSetProceso").CopyFromRecordset rs
    End If
    
    Hoja1.Range("dataSetProceso").Offset(0, 2).EntireColumn.NumberFormat = "hh:mm:ss am/pm"
    Hoja1.Range("dataSetProceso").Offset(0, 3).EntireColumn.NumberFormat = "hh:mm:ss"
    
    closeRS
End Sub

Public Function fechaDateStr(fechaDate As Date)
    fechaDateStr = Format(fechaDate, "YYYY") & "-" & Format(fechaDate, "MM") & "-" & Format(fechaDate, "DD")
End Function

Function FileAlreadyOpen(FullFileName As String) As Boolean
' returns True if FullFileName is currently in use by another process
' example: If FileAlreadyOpen("C:\FolderName\FileName.xls") Then...
Dim f As Integer
    f = FreeFile
    On Error Resume Next
    Open FullFileName For Binary Access Read Write Lock Read Write As #f
    Close #f
    ' If an error occurs, the document is currently open.
    If Err.Number <> 0 Then
        FileAlreadyOpen = True
        Err.Clear
        'MsgBox "Error #" & Str(Err.Number) & " - " & Err.Description
    Else
        FileAlreadyOpen = False
    End If
    On Error GoTo 0
End Function
