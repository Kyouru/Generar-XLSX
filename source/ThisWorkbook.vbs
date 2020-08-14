
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.ThisWorkbook.Saved = True
End Sub

Private Sub Workbook_Open()
    Dim alias As String
    Dim Ret As Variant
    
    [FECHA_HISTORICO] = Format(Now(), "yyyy-mm-dd")
    Hoja1.cbActualizar = False
    
    
    strSQL = "SELECT * FROM ALIAS WHERE NOMBRE = '" & Application.UserName & "'"
    
    OpenDB
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - Workbook_Open", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    If rs.RecordCount > 0 Then
        [DIRECTORIO] = rs.Fields("RUTA_DESCARGA")
    Else
        alias = ""
        Do While alias = ""
            alias = InputBox("Usuario """ & Application.UserName & """ no encontrado." & vbCrLf & "Por favor, ingrese un Alias para este usuario nuevo (Preferiblemente corto):", "Nuevo Usuario")
        Loop
        
        Ret = False
        Do While Ret = False
            Ret = Hoja1.BrowseForFolder()
        Loop
        
        [DIRECTORIO] = Ret
        strSQL = "INSERT INTO ALIAS (NOMBRE, NOMBRE_ALIAS, RUTA_DESCARGA) VALUES ('" & Application.UserName & "', '" & alias & "', '" & Ret & "')"
        
        On Error Resume Next
        cnn.Execute strSQL
        On Error GoTo 0
        
        'Log del Error
        If cnn.Errors.count > 0 Then
            Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - Workbook_Open (2)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
            cnn.Errors.Clear
            closeRS
            Exit Sub
        End If
    End If
    closeRS
    actualizarTodo
End Sub

