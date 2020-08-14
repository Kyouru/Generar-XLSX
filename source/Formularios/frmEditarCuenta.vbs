
Private Sub btEditarCampos_Click()
    frmParamCampo.Show
End Sub

Private Sub btGuardar_Click()
    strSQL = "UPDATE PARAM_CUENTA SET NOMBRE_CORTO = '" & Me.tbNombreCorto.Text & "', NUMERO_CORTO = '" & Me.tbNumeroCorto.Text & "', RANGO_CUENTA_COL = '" & Me.tbColCuenta.Text & "', RANGO_CUENTA_ROW = '" & Me.tbFilaCuenta.Text & "', INICIO_COL = '" & Me.tbColInicio.Text & "', INICIO_ROW = '" & Me.tbFilaInicio.Text & "', NOMBRE_ARCHIVO = '" & Me.tbNombreArchivo.Text & "', EXTENSION = '" & Me.tbExtArchivo.Text & "', IDENTIFICADOR_CUENTA = '" & Me.tbIdentificador.Text & "', ANULADO = "
    
    If Me.cbAnulado.Value = False Then
        strSQL = strSQL & "FALSE"
    Else
        strSQL = strSQL & "TRUE"
    End If
    strSQL = strSQL & " WHERE ID_PARAM_CUENTA = " & Me.lbIdParamCuenta.Caption
    
    OpenDB
    
    cnn.Execute strSQL
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    closeRS
    Unload Me
    frmParamCuenta.UserForm_Initialize
    
End Sub

Private Sub UserForm_Initialize()
    Me.lbIdParamCuenta.Caption = frmParamCuenta.ListBox1.List(frmParamCuenta.ListBox1.ListIndex, 0)
    
    strSQL = "SELECT NOMBRE_CORTO, NUMERO_CORTO, RANGO_CUENTA_COL, RANGO_CUENTA_ROW, INICIO_COL, INICIO_ROW, NOMBRE_ARCHIVO, EXTENSION, IDENTIFICADOR_CUENTA, ANULADO FROM PARAM_CUENTA WHERE ID_PARAM_CUENTA = " & Me.lbIdParamCuenta.Caption
    
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
        Me.tbNombreCorto.Text = rs.Fields("NOMBRE_CORTO")
        Me.tbNumeroCorto.Text = rs.Fields("NUMERO_CORTO")
        Me.tbColCuenta.Text = rs.Fields("RANGO_CUENTA_COL")
        Me.tbFilaCuenta.Text = rs.Fields("RANGO_CUENTA_ROW")
        Me.tbColInicio.Text = rs.Fields("INICIO_COL")
        Me.tbFilaInicio.Text = rs.Fields("INICIO_ROW")
        
        If Not IsNull(rs.Fields("NOMBRE_ARCHIVO")) Then
            Me.tbNombreArchivo.Text = rs.Fields("NOMBRE_ARCHIVO")
        End If
        If Not IsNull(rs.Fields("EXTENSION")) Then
            Me.tbExtArchivo.Text = rs.Fields("EXTENSION")
        End If
        
        Me.tbIdentificador.Text = rs.Fields("IDENTIFICADOR_CUENTA")
        
        Me.cbAnulado = rs.Fields("ANULADO")
    End If
    closeRS
End Sub
