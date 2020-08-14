
Private Sub btGuardar_Click()
    If cmbBanco.ListIndex <> -1 And tbColSaldo.Text <> "" And tbFilaSaldo.Text <> "" And tbTratamiento.Text <> "" And tbComentario.Text <> "" Then
        If IsNumeric(tbColSaldo.Text) And IsNumeric(tbFilaSaldo.Text) Then
            strSQL = "INSERT INTO PARAM_SALDO (ID_BANCO_FK, SALDO_COL, SALDO_ROW, TRATAMIENTO, COMENTARIO) VALUES (" & _
                    cmbBanco.List(cmbBanco.ListIndex, 1) & ", " & tbColSaldo.Text & ", " & tbFilaSaldo.Text & ", '" & tbTratamiento.Text & "', " & tbComentario.Text & ")"
            
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
            
            Unload Me
            frmParamSaldo.UserForm_Initialize
            
        End If
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
    
    closeRS
    
End Sub
