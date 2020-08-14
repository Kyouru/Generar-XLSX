
Public Sub UserForm_Initialize()
    strSQL = "SELECT ID_PARAM_CUENTA, NOMBRE_CORTO + ' ' + NUMERO_CORTO, RANGO_CUENTA_COL, RANGO_CUENTA_ROW, INICIO_COL, INICIO_ROW, NOMBRE_ARCHIVO, EXTENSION, IDENTIFICADOR_CUENTA FROM PARAM_CUENTA ORDER BY NOMBRE_CORTO + ' ' + NUMERO_CORTO"
    
    'Limpiar Hoja
    ThisWorkbook.Sheets("TEMP2").Range(ThisWorkbook.Sheets("TEMP2").Range("dataSetTemp2"), ThisWorkbook.Sheets("TEMP2").Range("dataSetTemp2").End(xlDown)).ClearContents
    
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
        ThisWorkbook.Sheets("TEMP2").Range("dataSetTemp2").Cells(1, 1).CopyFromRecordset rs
    End If
    closeRS
    
    ActualizarLista
End Sub

Public Sub ActualizarLista()
    With ThisWorkbook.Sheets("TEMP2")
        ListBox1.ColumnWidths = "0;60;60;60;60;60;90;30;150;"
        ListBox1.ColumnCount = 9
        ListBox1.ColumnHeads = True
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!" & Left(.Range("dataSetTemp2").Address, Len(.Range("dataSetTemp2").Address) - 1) & .Range("A2").End(xlDown).Row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!" & .Range("dataSetTemp2").Address
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
    End With
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If ListBox1.ListIndex <> -1 Then
        frmEditarCuenta.Show
    End If
End Sub
