
Private Sub btGuardar_Click()
    If cmbBanco.ListIndex <> -1 And cmbTipoMovimiento.ListIndex <> -1 And cmbConcidencia.ListIndex <> -1 And tbDescripcion.Text <> "" Then
        strSQL = "UPDATE PARAM_MOVIMIENTO SET" & _
        " ID_TIPO_MOVIMIENTO_FK = " & cmbTipoMovimiento.List(cmbTipoMovimiento.ListIndex, 1) & _
        ", ID_BANCO_FK = " & cmbBanco.List(cmbBanco.ListIndex, 1) & _
        ", ID_COINCIDENCIA_FK = " & cmbConcidencia.List(cmbConcidencia.ListIndex, 1) & _
        ", DESCRIPCION = '" & Replace(tbDescripcion.Text, "'", "''") & "'"
        If Me.tbComentario <> "" Then
            strSQL = strSQL & ", COMENTARIO = '" & Replace(tbComentario.Text, "'", "''") & "'"
        Else
            strSQL = strSQL & ", COMENTARIO = NULL"
        End If
        strSQL = strSQL & " WHERE ID_PARAM_MOVIMIENTO = " & Me.lbIdParamMov.Caption
        OpenDB
        
        On Error Resume Next
        cnn.Execute strSQL
        On Error GoTo 0
        
        'Log del Error
        If cnn.Errors.count > 0 Then
            Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
            cnn.Errors.Clear
            closeRS
            Exit Sub
        End If
        
        closeRS
        MsgBox "Modificado Exitosamente"
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
    
    Set rs = Nothing
    
    Me.lbIdParamMov.Caption = frmParamMov.ListBox1.List(frmParamMov.ListBox1.ListIndex, 0)
    
    strSQL = "SELECT * FROM PARAM_MOVIMIENTO WHERE ID_PARAM_MOVIMIENTO = " & Me.lbIdParamMov.Caption
    
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
        Dim i As Integer
        i = 0
        While i < cmbBanco.ListCount
            If cmbBanco.List(i, 1) = CInt(rs.Fields("ID_BANCO_FK")) Then
                cmbBanco.ListIndex = i
                i = cmbBanco.ListCount
            End If
            i = i + 1
        Wend
        
        i = 0
        While i < cmbTipoMovimiento.ListCount
            If cmbTipoMovimiento.List(i, 1) = CInt(rs.Fields("ID_TIPO_MOVIMIENTO_FK")) Then
                cmbTipoMovimiento.ListIndex = i
                i = cmbTipoMovimiento.ListCount
            End If
            i = i + 1
        Wend
        
        i = 0
        While i < cmbConcidencia.ListCount
            If cmbConcidencia.List(i, 1) = CInt(rs.Fields("ID_COINCIDENCIA_FK")) Then
                cmbConcidencia.ListIndex = i
                i = cmbConcidencia.ListCount
            End If
            i = i + 1
        Wend
        
        tbDescripcion.Text = rs.Fields("DESCRIPCION")
        
        If Not IsNull(rs.Fields("COMENTARIO")) Then
            tbComentario.Text = rs.Fields("COMENTARIO")
        End If
        
    End If
    
    closeRS
    
End Sub
