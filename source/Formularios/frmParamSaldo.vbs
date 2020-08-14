Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btNuevo_Click()

End Sub

Public Sub UserForm_Initialize()
    ActualizarHoja
    ActualizarLista
End Sub

Public Sub ActualizarHoja()
    
    strSQL = "SELECT ID_PARAM_SALDO, ID_BANCO_FK, NOMBRE_BANCO, SALDO_COL, SALDO_ROW, TRATAMIENTO, COMENTARIO, ANULADO FROM PARAM_SALDO LEFT JOIN BANCO ON BANCO.ID_BANCO = PARAM_SALDO.ID_BANCO_FK"
    
    'Limpiar Hoja
    ThisWorkbook.Sheets("TEMP5").Range(ThisWorkbook.Sheets("TEMP5").Range("dataSetTemp5"), ThisWorkbook.Sheets("TEMP5").Range("dataSetTemp5").End(xlDown)).ClearContents
    
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
        ThisWorkbook.Sheets("TEMP5").Range("dataSetTemp5").Cells(1, 1).CopyFromRecordset rs
    End If
    
    closeRS
End Sub

Public Sub ActualizarLista()
    With ThisWorkbook.Sheets("TEMP5")
        ListBox1.ColumnWidths = "0;0;80;40;40;40;40"
        ListBox1.ColumnCount = 7
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


