
Private Sub btAgregar_Click()
    If cmbBanco.ListIndex <> -1 And cmbTipoMovimiento.ListIndex <> -1 And cmbConcidencia.ListIndex <> -1 And tbDescripcion.Text <> "" Then
        strSQL = "INSERT INTO PARAM_MOVIMIENTO (ID_TIPO_MOVIMIENTO_FK, ID_BANCO_FK, ID_COINCIDENCIA_FK, DESCRIPCION, COMENTARIO) VALUES ("
        strSQL = strSQL & cmbTipoMovimiento.List(cmbTipoMovimiento.ListIndex, 1) & ", " & cmbBanco.List(cmbBanco.ListIndex, 1) & ", " & cmbConcidencia.List(cmbConcidencia.ListIndex, 1) & ", '" & Replace(tbDescripcion.Text, "'", "''") & "'"
        If Me.tbComentario <> "" Then
            strSQL = strSQL & ", '" & Replace(tbComentario.Text, "'", "''") & "')"
        Else
            strSQL = strSQL & ", NULL)"
        End If
        
        OpenDB
        
        On Error Resume Next
        cnn.Execute strSQL
        On Error GoTo 0
        
        'Log del Error
        If cnn.Errors.count > 0 Then
            Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btAgregar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
            cnn.Errors.Clear
            closeRS
            Exit Sub
        End If
        
        closeRS
        
        MsgBox "Agregado Exitosamente"
        frmParamMov.UserForm_Initialize
        Unload Me
    End If
End Sub

Private Sub UserForm_Initialize()
    strSQL = "SELECT * FROM BANCO"
    
    OpenDB
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    If rs.RecordCount > 0 Then
        cmbBanco.Clear
        While Not rs.EOF
            cmbBanco.AddItem rs.Fields("NOMBRE_BANCO")
            cmbBanco.List(cmbBanco.ListCount - 1, 1) = rs.Fields("ID_BANCO")
            rs.MoveNext
        Wend
    End If
    
    Set rs = Nothing
    
    strSQL = "SELECT * FROM TIPO_MOVIMIENTO"
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize (2)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    If rs.RecordCount > 0 Then
        cmbTipoMovimiento.Clear
        While Not rs.EOF
            cmbTipoMovimiento.AddItem rs.Fields("NOMBRE_TIPO_MOVIMIENTO")
            cmbTipoMovimiento.List(cmbTipoMovimiento.ListCount - 1, 1) = rs.Fields("ID_TIPO_MOVIMIENTO")
            rs.MoveNext
        Wend
    End If
    
    Set rs = Nothing
    
    strSQL = "SELECT * FROM COINCIDENCIA"
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize (3)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    If rs.RecordCount > 0 Then
        cmbConcidencia.Clear
        While Not rs.EOF
            cmbConcidencia.AddItem rs.Fields("NOMBRE_COINCIDENCIA")
            cmbConcidencia.List(cmbConcidencia.ListCount - 1, 1) = rs.Fields("ID_COINCIDENCIA")
            rs.MoveNext
        Wend
    End If
    
    closeRS
    
End Sub
