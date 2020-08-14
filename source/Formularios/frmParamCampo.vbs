
Private Sub btAgregar_Click()
    If Me.cmbCampo.ListIndex <> -1 And Me.tbColumna.Text <> "" Then
        If IsNumeric(Me.tbColumna.Text) Then
            strSQL = "INSERT INTO PARAM_CAMPO (ID_PARAM_CUENTA_FK, ID_CAMPO_FK, COLUMNA) VALUES (" & frmEditarCuenta.lbIdParamCuenta.Caption & ", " & Me.cmbCampo.List(Me.cmbCampo.ListIndex, 1) & ", " & Me.tbColumna.Text & ")"
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
            
            tbColumna.Text = ""
            
            ActualizarComboBox
            ActualizarHoja
            ActualizarLista
        End If
    End If
End Sub

Private Sub btEliminar_Click()
    If Me.ListBox1.ListIndex <> -1 Then
        strSQL = "DELETE FROM PARAM_CAMPO WHERE ID_PARAM_CAMPO = " & Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
        OpenDB
    
        On Error Resume Next
        cnn.Execute strSQL
        On Error GoTo 0
        
        'Log del Error
        If cnn.Errors.count > 0 Then
            Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btEliminar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
            cnn.Errors.Clear
            closeRS
            Exit Sub
        End If
        
        closeRS
        
        ActualizarHoja
        ActualizarLista
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox1.ListIndex <> -1 Then
        Dim resp As Variant
        Dim nuevaColumna As Integer
        resp = InputBox("Introdusca la nueva columna")
        If resp <> "" Then
            If IsNumeric(resp) Then
                nuevaColumna = CInt(resp)
                strSQL = "UPDATE PARAM_CAMPO SET COLUMNA = " & nuevaColumna & " WHERE ID_PARAM_CAMPO = " & Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
                OpenDB
            
                On Error Resume Next
                cnn.Execute strSQL
                On Error GoTo 0
                
                'Log del Error
                If cnn.Errors.count > 0 Then
                    Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ListBox1_DblClick", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
                    cnn.Errors.Clear
                    closeRS
                    Exit Sub
                End If
                closeRS
                ActualizarHoja
                ActualizarLista
            End If
        End If
    End If
End Sub

Public Sub UserForm_Initialize()
    ActualizarComboBox
    ActualizarHoja
    ActualizarLista
End Sub

Public Sub ActualizarComboBox()
    strSQL = "SELECT ID_CAMPO, NOMBRE_CAMPO FROM CAMPO WHERE ID_CAMPO NOT IN (SELECT ID_CAMPO_FK FROM PARAM_CAMPO WHERE ID_PARAM_CUENTA_FK = " & frmEditarCuenta.lbIdParamCuenta.Caption & ") ORDER BY ID_CAMPO"
    OpenDB
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarComboBox", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    cmbCampo.Clear
    Do While Not rs.EOF
        cmbCampo.AddItem rs.Fields("NOMBRE_CAMPO")
        cmbCampo.List(cmbCampo.ListCount - 1, 1) = rs.Fields("ID_CAMPO")
        rs.MoveNext
    Loop
    
    If cmbCampo.ListCount > 0 Then
        cmbCampo.ListIndex = 0
    End If
    
    closeRS
End Sub

Public Sub ActualizarHoja()
    strSQL = "SELECT ID_PARAM_CAMPO, NOMBRE_CAMPO, COLUMNA FROM PARAM_CAMPO LEFT JOIN CAMPO ON PARAM_CAMPO.ID_CAMPO_FK = CAMPO.ID_CAMPO WHERE ID_PARAM_CUENTA_FK = " & frmEditarCuenta.lbIdParamCuenta.Caption & " ORDER BY ID_CAMPO_FK"
    
    'Limpiar Hoja
    ThisWorkbook.Sheets("TEMP3").Range(ThisWorkbook.Sheets("TEMP3").Range("dataSetTemp3"), ThisWorkbook.Sheets("TEMP3").Range("dataSetTemp3").End(xlDown)).ClearContents
    
    OpenDB
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    If rs.RecordCount > 0 Then
        ThisWorkbook.Sheets("TEMP3").Range("dataSetTemp3").Cells(1, 1).CopyFromRecordset rs
    End If
    
    closeRS
End Sub

Public Sub ActualizarLista()
    With ThisWorkbook.Sheets("TEMP3")
        ListBox1.ColumnWidths = "30;100;50"
        ListBox1.ColumnCount = 3
        ListBox1.ColumnHeads = True
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!" & Left(.Range("dataSetTemp3").Address, Len(.Range("dataSetTemp3").Address) - 1) & .Range("A2").End(xlDown).Row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!" & .Range("dataSetTemp3").Address
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
    End With
End Sub
