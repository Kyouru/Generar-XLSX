Private Sub btAceptar_Click()
    If ListBox1.ListIndex <> -1 Then
        Dim strDir As String
        Dim tmpXLSX As String
        Dim tempDate As Date
        
        tempDate = ListBox1.List(ListBox1.ListIndex, 3)
        
        tmpXLSX = ListBox1.List(ListBox1.ListIndex, 1) & " " & ListBox1.List(ListBox1.ListIndex, 2)
        
        strDir = ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(tempDate, "YYYY") & "\" & Format(tempDate, "MM") & "\" & Format(tempDate, "YYYY.MM.DD") & "\" & tmpXLSX & ".xlsx"
        
        If Dir(strDir) <> "" Then
            Unload Me
            Workbooks.Open (strDir)
        Else
            MsgBox "Error: No se encontr・el XLSX" & vbCrLf & ">> Ruta: " & strDir
        End If
    End If
End Sub

Private Sub btBuscar_Click()
        strSQL = "SELECT ID_CARGA_XLSX, NOMBRE_CORTO + ' ' + NUMERO_CORTO, IIF(CORRELATIVO <= 9, '0' + CSTR(CORRELATIVO), CSTR(CORRELATIVO)), FORMAT(FECHA, 'DD/MM/YYYY'), FORMAT(FECHA, 'hh:mm:ss'), FORMAT(FECHA_GEN, 'DD/MM/YYYY'), USUARIO FROM (CARGA_XLSX LEFT JOIN PARAM_CUENTA ON PARAM_CUENTA.ID_PARAM_CUENTA = CARGA_XLSX.ID_PARAM_CUENTA_FK) WHERE 1=1"
        
        If Sheets("CARGAS").Range("TMP_FECHA") <> "" Then
            If IsDate(Sheets("CARGAS").Range("TMP_FECHA")) Then
                strSQL = strSQL & " AND FORMAT(FECHA, 'YYYY-MM-DD') = '" & Format(Sheets("CARGAS").Range("TMP_FECHA"), "YYYY-MM-DD") & "'"
            End If
        End If
        
        If tbCuenta.Text <> "" Then
            strSQL = strSQL & " AND NOMBRE_CORTO + ' ' + NUMERO_CORTO LIKE '%" & tbCuenta.Text & "%'"
        End If
        
        If obFechaHora Then
            strSQL = strSQL & " ORDER BY FECHA DESC, FECHA_GEN DESC, NOMBRE_CORTO + ' ' + NUMERO_CORTO, CORRELATIVO"
        Else
            If obCuenta Then
                strSQL = strSQL & " ORDER BY NOMBRE_CORTO + ' ' + NUMERO_CORTO, CORRELATIVO, FECHA DESC, FECHA_GEN DESC"
            End If
        End If
        
        'Limpiar Hoja
        ThisWorkbook.Sheets("TEMP").Range(ThisWorkbook.Sheets("TEMP").Range("dataSetTemp"), ThisWorkbook.Sheets("TEMP").Range("dataSetTemp").End(xlDown)).ClearContents
        
        OpenDB
        
        On Error Resume Next
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        
        'Log del Error
        If cnn.Errors.count > 0 Then
            Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btBuscar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
            cnn.Errors.Clear
            closeRS
            Exit Sub
        End If
        
        If rs.RecordCount > 0 Then
            ThisWorkbook.Sheets("TEMP").Range("dataSetTemp").Cells(1, 1).CopyFromRecordset rs
        End If
        closeRS
        
        ActualizarLista
End Sub

Private Sub btCalendario_Click()
    Dim strFecha As String
    Dim tmpFecha As Date
    Sheets("CARGAS").Range("TMP_FECHA") = ""
    frmCalendario.Show
    If [TMP_FECHA] <> "" Then
        If IsDate([TMP_FECHA]) Then
            tmpFecha = fechaDateStr([TMP_FECHA])
            lbFecha.Caption = Format(tmpFecha, "DD/MM/YYYY")
            btBuscar_Click
        End If
    End If
End Sub

Private Sub btCancelar_Click()
    Unload Me
End Sub

'Agrega la Hoja Temporal a la ListBox
Public Sub ActualizarLista()
    With ThisWorkbook.Sheets("TEMP")
        ListBox1.ColumnWidths = "0;80;20;50;50;50;150;"
        ListBox1.ColumnCount = 7
        ListBox1.ColumnHeads = True
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!" & Left(.Range("dataSetTemp").Address, Len(.Range("dataSetTemp").Address) - 1) & .Range("A2").End(xlDown).Row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!" & .Range("dataSetTemp").Address
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
    End With
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btAceptar_Click
End Sub

Private Sub obFechaHora_Change()
    btBuscar_Click
End Sub

Private Sub tbCuenta_Change()
    btBuscar_Click
End Sub

Private Sub UserForm_Initialize()
    lbFecha.Caption = Format([TMP_FECHA], "DD/MM/YYYY")
    btBuscar_Click
End Sub
