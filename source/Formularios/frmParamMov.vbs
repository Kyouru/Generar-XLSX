
Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btEditar_Click()
    If ListBox1.ListIndex <> -1 Then
        editarParamMov.Show
    End If
End Sub

Private Sub btNuevo_Click()
    nuevoParamMov.Show
End Sub

Public Sub UserForm_Initialize()
    ActualizarHoja
    ActualizarLista
End Sub

Public Sub ActualizarHoja()

    strSQL = "SELECT ID_PARAM_MOVIMIENTO, ID_TIPO_MOVIMIENTO, ID_COINCIDENCIA, NOMBRE_BANCO, NOMBRE_TIPO_MOVIMIENTO, DESCRIPCION, NOMBRE_COINCIDENCIA, COMENTARIO FROM ((PARAM_MOVIMIENTO LEFT JOIN BANCO ON BANCO.ID_BANCO = PARAM_MOVIMIENTO.ID_BANCO_FK) LEFT JOIN TIPO_MOVIMIENTO ON TIPO_MOVIMIENTO.ID_TIPO_MOVIMIENTO = PARAM_MOVIMIENTO.ID_TIPO_MOVIMIENTO_FK) LEFT JOIN COINCIDENCIA ON COINCIDENCIA.ID_COINCIDENCIA = PARAM_MOVIMIENTO.ID_COINCIDENCIA_FK ORDER BY ID_BANCO, ID_TIPO_MOVIMIENTO, ID_COINCIDENCIA"
    
    'Limpiar Hoja
    ThisWorkbook.Sheets("TEMP4").Range(ThisWorkbook.Sheets("TEMP4").Range("dataSetTemp4"), ThisWorkbook.Sheets("TEMP4").Range("dataSetTemp4").End(xlDown)).ClearContents
    
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
        ThisWorkbook.Sheets("TEMP4").Range("dataSetTemp4").Cells(1, 1).CopyFromRecordset rs
    End If
    
    closeRS
End Sub

Public Sub ActualizarLista()
    With ThisWorkbook.Sheets("TEMP4")
        ListBox1.ColumnWidths = "0;0;0;130;60;100;50;80"
        ListBox1.ColumnCount = 8
        ListBox1.ColumnHeads = True
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!" & Left(.Range("dataSetTemp4").Address, Len(.Range("dataSetTemp4").Address) - 1) & .Range("A2").End(xlDown).Row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!" & .Range("dataSetTemp4").Address
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
    End With
End Sub

