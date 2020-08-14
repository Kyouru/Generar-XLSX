
Public proxActualizacion As Date
Public strUltimaCarga As String
Public cargar As Boolean

'Abrir el Explorador de Windows en la ruta configurada
Private Sub btAbrirRuta_Click()
    If [DIRECTORIO] <> False Then
        If dirExists([DIRECTORIO]) Then
            Shell "C:\WINDOWS\explorer.exe """ & [DIRECTORIO] & "", vbNormalFocus
        End If
    End If
End Sub

Public Sub btActualizar_Click()
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
        
    actualizarTodo
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

Public Sub btCarga_Click()
    
    Me.Select
    
    
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
        
    Dim wb As Workbook
    Dim rng As Range
    Dim guardado As Boolean         'Variable utilizada para saber si el libro se guardo, para luego eliminarlo
    Dim file As Variant
    Dim fileDate As Date
    Dim tmpSplit As Variant
    Dim tmpGlosa As String
    Dim nOperacion As String
    Dim i As Integer
    
    'Variables de la cuenta identificada
    Dim id_param_cuenta As Integer
    Dim id_banco_param_cuenta As Integer
    Dim id_cuenta As Integer
    Dim nombre_archivo As String
    Dim extension As String
    Dim rango_cuenta_col As Integer
    Dim rango_cuenta_row As Integer
    Dim inicio_col As Integer
    Dim inicio_row As Integer
    Dim identificador_cuenta As String
    Dim nombre_corto As String
    Dim numero_corto As String
    
    'Variable de los campos de la cuenta identificada
    Dim fecha_movimiento As Integer
    Dim monto As Integer
    Dim monto2 As Integer
    Dim glosa As Integer
    Dim numero_operacion As Integer
    Dim hora_movimiento As Integer
    Dim itf As Integer
    Dim pre_num_ope As Integer
    
    Dim Data(), r As Long
    Dim count As Integer
    
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    'Descomprimir Zip que empiezen por Movimientos_
    If [DESCOMPRIMIR_ZIP] = "SI" Then
        Dim oApp As Object
        Set oApp = CreateObject("Shell.Application")
        
        file = Dir([DIRECTORIO])
        
        While (file <> "")
            If Right(file, 4) = ".zip" And Left(file, 12) = "Movimientos_" Then
                oApp.Namespace(Me.Range("DIRECTORIO").Value).CopyHere oApp.Namespace(Me.Range("DIRECTORIO").Value & file).Items
                Kill Me.Range("DIRECTORIO").Value & file
            End If
            file = Dir
        Wend
        
        'Exit Sub
    End If
    'Valores Iniciales
    file = Dir([DIRECTORIO])
    
    'Crear las carpetas para almacenar los extractos
    If Not dirExists(ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(Now(), "YYYY") & "\") Then
        MkDir ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(Now(), "YYYY") & "\"
    End If
    If Not dirExists(ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(Now(), "YYYY") & "\" & Format(Now(), "MM")) Then
        MkDir ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(Now(), "YYYY") & "\" & Format(Now(), "MM")
    End If
    If Not dirExists(ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(Now(), "YYYY") & "\" & Format(Now(), "MM") & "\" & Format(Now(), "YYYY.MM.DD")) Then
        MkDir ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(Now(), "YYYY") & "\" & Format(Now(), "MM") & "\" & Format(Now(), "YYYY.MM.DD")
    End If
    If Not dirExists(ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(Now(), "YYYY") & "\" & Format(Now(), "MM") & "\" & Format(Now(), "YYYY.MM.DD") & "\RAW") Then
        MkDir ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(Now(), "YYYY") & "\" & Format(Now(), "MM") & "\" & Format(Now(), "YYYY.MM.DD") & "\RAW"
    End If
    
    OpenDB
    
    While (file <> "")
        If file <> "tmpCarga" Then
            fileDate = FileLastModified([DIRECTORIO] & file)
            cargar = False
            
            If [BACKUP_EXTRACTO] = "SI" Then
                
                On Error Resume Next
                Kill [DIRECTORIO] & "tmpCarga"
                On Error GoTo 0
                
                tmpSplit = Split(file, ".")
                Call fso.CopyFile([DIRECTORIO] & file, [DIRECTORIO] & "tmpCarga", True)
            End If
            
            If [RENOMBRAR_EEXPORT] = "SI" Then
                If Left(file, 7) = "eExport" And UBound(Split(file, ".")) = 0 Then
                    Name [DIRECTORIO] & file As [DIRECTORIO] & file & ".csv"
                    file = file & ".csv"
                End If
            End If
            
            If [CONVERTIR_BIFTXT_XLSX] = "SI" Then
                If InStr(1, file, "001000005126") > 0 Or InStr(1, file, "001000005134") > 0 Or InStr(1, file, "001000006228") > 0 Or InStr(1, file, "001000006236") > 0 Then
                    Workbooks.Open ([DIRECTORIO] & file)
                    Set wb = Workbooks(file)
                    If wb.Sheets(1).Cells(1, 1) = "DESCRIPCION                    FECHA    N. REFER. DEBITO           CREDITO         " Then
                        wb.Sheets(1).Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
                            FieldInfo:=Array(Array(0, 2), Array(31, 2), Array(40, 2), Array(50, 2), Array(67, 2)) _
                            , TrailingMinusNumbers:=True
                        wb.Sheets(1).Cells(1, 3) = "N. REFERENCIA"
                        
                        Dim fileAnterior As String
                        fileAnterior = file
                        
                        If UBound(Split(file, " (")) = 2 Then
                            file = Split(file, " (")(0) & "1" & correlativoArchivo([DIRECTORIO], CStr(Split(file, " (")(0) & "1.csv")) & ".csv"
                        Else
                            file = Split(file, ".")(0) & "1" & correlativoArchivo([DIRECTORIO], CStr(Split(file, ".")(0) & "1.csv")) & ".csv"
                        End If
                        wb.SaveAs [DIRECTORIO] & file, FileFormat:=56
                        Kill [DIRECTORIO] & fileAnterior
                        cerrarWb wb
                    End If
                End If
            End If
            
            tmpSplit = Split(file, ".")
            
            If UBound(tmpSplit) >= 1 Then
                'Identificar Extracto
                strSQL = "SELECT * FROM PARAM_CUENTA WHERE INSTR('" & tmpSplit(0) & "', NOMBRE_ARCHIVO) AND EXTENSION = '" & tmpSplit(UBound(tmpSplit)) & "'"
                
                Set rs = Nothing
                On Error Resume Next
                rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
                On Error GoTo 0
                
                'Log del Error
                If cnn.Errors.count > 0 Then
                    Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btCarga_Click (2)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
                    cnn.Errors.Clear
                    closeRS
                    Exit Sub
                End If
                If rs.RecordCount > 0 Then
                    Workbooks.Open ([DIRECTORIO] & file)
                    Set wb = Workbooks(file)
                    If rs.RecordCount > 1 Then
                        Do While Not rs.EOF
                            If wb.Sheets(1).Cells(rs.Fields("RANGO_CUENTA_ROW"), rs.Fields("RANGO_CUENTA_COL")) = rs.Fields("IDENTIFICADOR_CUENTA") Then
                    
                                id_param_cuenta = rs.Fields("ID_PARAM_CUENTA")
                                id_banco_param_cuenta = rs.Fields("ID_BANCO_FK")
                                nombre_archivo = rs.Fields("NOMBRE_ARCHIVO")
                                rango_cuenta_col = rs.Fields("RANGO_CUENTA_COL")
                                rango_cuenta_row = rs.Fields("RANGO_CUENTA_ROW")
                                identificador_cuenta = rs.Fields("IDENTIFICADOR_CUENTA")
                                inicio_col = rs.Fields("INICIO_COL")
                                inicio_row = rs.Fields("INICIO_ROW")
                                nombre_corto = rs.Fields("NOMBRE_CORTO")
                                numero_corto = rs.Fields("NUMERO_CORTO")
                                extension = rs.Fields("EXTENSION")
                                id_cuenta = rs.Fields("ID_CUENTA_FK")
                                
                                cargar = True
                                
                            End If
                            rs.MoveNext
                        Loop
                    ElseIf rs.RecordCount = 1 Then
                        id_param_cuenta = rs.Fields("ID_PARAM_CUENTA")
                        id_banco_param_cuenta = rs.Fields("ID_BANCO_FK")
                        nombre_archivo = rs.Fields("NOMBRE_ARCHIVO")
                        rango_cuenta_col = rs.Fields("RANGO_CUENTA_COL")
                        rango_cuenta_row = rs.Fields("RANGO_CUENTA_ROW")
                        identificador_cuenta = rs.Fields("IDENTIFICADOR_CUENTA")
                        inicio_col = rs.Fields("INICIO_COL")
                        inicio_row = rs.Fields("INICIO_ROW")
                        nombre_corto = rs.Fields("NOMBRE_CORTO")
                        numero_corto = rs.Fields("NUMERO_CORTO")
                        extension = rs.Fields("EXTENSION")
                        id_cuenta = rs.Fields("ID_CUENTA_FK")
                        
                        cargar = True
                        
                    End If
                    
                    If cargar And wb.Sheets(1).Cells(inicio_row, inicio_col) <> "" Then
                        
                        strSQL = "SELECT * FROM PARAM_CAMPO WHERE ID_PARAM_CUENTA_FK = " & id_param_cuenta
                        
                        Set rs = Nothing
                        
                        On Error Resume Next
                        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
                        On Error GoTo 0
                        
                        'Log del Error
                        If cnn.Errors.count > 0 Then
                            Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btCarga_Click (3)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
                            cnn.Errors.Clear
                            closeRS
                            Exit Sub
                        End If
                        
                        fecha_movimiento = 0    'ID_CAMPO: 1
                        monto = 0               'ID_CAMPO: 2
                        monto2 = 0              'ID_CAMPO: 3
                        glosa = 0               'ID_CAMPO: 4
                        numero_operacion = 0    'ID_CAMPO: 5
                        hora_movimiento = 0     'ID_CAMPO: 6
                        itf = 0                 'ID_CAMPO: 7
                        pre_num_ope = 0         'ID_CAMPO: 8
                        
                        Do While Not rs.EOF
                            Select Case rs.Fields("ID_CAMPO_FK")
                            Case 1
                                fecha_movimiento = rs.Fields("COLUMNA")
                            Case 2
                                monto = rs.Fields("COLUMNA")
                            Case 3
                                monto2 = rs.Fields("COLUMNA")
                            Case 4
                                glosa = rs.Fields("COLUMNA")
                            Case 5
                                numero_operacion = rs.Fields("COLUMNA")
                            Case 6
                                hora_movimiento = rs.Fields("COLUMNA")
                            Case 7
                                itf = rs.Fields("COLUMNA")
                            Case 8
                                pre_num_ope = rs.Fields("COLUMNA")
                            End Select
                            
                            rs.MoveNext
                        Loop
                        
                        'Exceso de fechas en el extracto
                        Dim unique As New Scripting.Dictionary
                        
                        Data = wb.Sheets(1).Range(wb.Sheets(1).Cells(inicio_row, fecha_movimiento), wb.Sheets(1).Cells(inicio_row + 1000, fecha_movimiento)).Value
                        
                        Set unique = CreateObject("Scripting.Dictionary")
                        For r = 1 To UBound(Data)
                            unique(Data(r, 1)) = Empty
                        Next r
                        If unique.count > 9 Then
                            closeRS
                            MsgBox ("Se encontro muchas fechas (+7)")
                            cerrarWb wb
                            If Not dirExists([DIRECTORIO] & "\ERROR_FECHA") Then
                                MkDir [DIRECTORIO] & "\ERROR_FECHA"
                            End If
                            Name [DIRECTORIO] & file As [DIRECTORIO] & "ERROR_FECHA\" & Split(file, ".")(0) & correlativoArchivo([DIRECTORIO] & "ERROR_FECHA\", CStr(file)) & "." & Split(file, ".")(1)
                            On Error Resume Next
                            Kill [DIRECTORIO] & "tmpCarga"
                            On Error GoTo 0
                            Exit Sub
                        End If
                        
                        registrarCarga id_param_cuenta, nombre_corto, numero_corto, fileDate
                        
                        'Casos Especiales
                        Set rng = wb.Sheets(1).Range(wb.Sheets(1).Cells(inicio_row, inicio_col), wb.Sheets(1).Cells(inicio_row, inicio_col + 50))
                        If wb.Sheets(1).Cells(inicio_row, inicio_col).Offset(1, 0) <> "" Then
                            i = wb.Sheets(1).Cells(inicio_row, inicio_col).End(xlDown).Row - wb.Sheets(1).Cells(inicio_row, inicio_col).Row + 1
                        Else
                            i = 1
                        End If
                        
                        'Nuevo formato IBK 2020-08-07
                        If [FORMATO_IBK] = "SI" Then
                            If wb.Sheets(1).Cells(13, 3) = "Fecha de operaci" And _
                                wb.Sheets(1).Cells(13, 6) = "Fecha de proceso" And _
                                wb.Sheets(1).Cells(13, 11) = "Nro. de operaci" And _
                                wb.Sheets(1).Cells(13, 14) = "Movimiento" And _
                                wb.Sheets(1).Cells(13, 18) = "Descripci" And _
                                wb.Sheets(1).Cells(13, 23) = "Canal" And _
                                wb.Sheets(1).Cells(13, 26) = "Cargo" And _
                                wb.Sheets(1).Cells(13, 32) = "Abono" Then
                                Dim premonto As String
                                i = inicio_row
                                Workbooks.Open (ThisWorkbook.Path & Application.PathSeparator & "Template IBK.xlsx")
                                Set wb2 = Workbooks("Template IBK.xlsx")
                                If wb.Sheets(1).Cells(8, 7) = "Corriente Soles 041-3000106330" Then
                                    'MsgBox "Formato IBK OK"
                                    
                                    'wb.Sheets(1).Cells(14, 2) = wb.Sheets(1).Cells(13, 3)
                                    'wb.Sheets(1).Cells(14, 3) = wb.Sheets(1).Cells(13, 3)
                                    'wb.Sheets(1).Cells(14, 4) = wb.Sheets(1).Cells(13, 3)
                                    
                                    'wb.Sheets(1).Cells(14, 5) = wb.Sheets(1).Cells(13, 6)
                                    'wb.Sheets(1).Cells(14, 6) = wb.Sheets(1).Cells(13, 6)
                                    'wb.Sheets(1).Cells(14, 7) = wb.Sheets(1).Cells(13, 6)
                                    'wb.Sheets(1).Cells(14, 8) = wb.Sheets(1).Cells(13, 6)
                                    'wb.Sheets(1).Cells(14, 9) = wb.Sheets(1).Cells(13, 6)
                                    
                                    'wb.Sheets(1).Cells(14, 10) = wb.Sheets(1).Cells(13, 11)
                                    'wb.Sheets(1).Cells(14, 11) = wb.Sheets(1).Cells(13, 11)
                                    'wb.Sheets(1).Cells(14, 12) = wb.Sheets(1).Cells(13, 11)
                                    
                                    'wb.Sheets(1).Cells(14, 13) = wb.Sheets(1).Cells(13, 14)
                                    'wb.Sheets(1).Cells(14, 14) = wb.Sheets(1).Cells(13, 14)
                                    'wb.Sheets(1).Cells(14, 15) = wb.Sheets(1).Cells(13, 14)
                                    'wb.Sheets(1).Cells(14, 16) = wb.Sheets(1).Cells(13, 14)
                                    
                                    'wb.Sheets(1).Cells(14, 17) = wb.Sheets(1).Cells(13, 18)
                                    'wb.Sheets(1).Cells(14, 18) = wb.Sheets(1).Cells(13, 18)
                                    'wb.Sheets(1).Cells(14, 19) = wb.Sheets(1).Cells(13, 18)
                                    'wb.Sheets(1).Cells(14, 20) = wb.Sheets(1).Cells(13, 18)
                                    'wb.Sheets(1).Cells(14, 21) = wb.Sheets(1).Cells(13, 18)
                                    
                                    'wb.Sheets(1).Cells(14, 22) = wb.Sheets(1).Cells(13, 23)
                                    'wb.Sheets(1).Cells(14, 23) = wb.Sheets(1).Cells(13, 23)
                                    'wb.Sheets(1).Cells(14, 24) = wb.Sheets(1).Cells(13, 23)
                                    
                                    'wb.Sheets(1).Cells(14, 25) = wb.Sheets(1).Cells(13, 26)
                                    'wb.Sheets(1).Cells(14, 26) = wb.Sheets(1).Cells(13, 26)
                                    'wb.Sheets(1).Cells(14, 27) = wb.Sheets(1).Cells(13, 26)
                                    'wb.Sheets(1).Cells(14, 28) = wb.Sheets(1).Cells(13, 26)
                                    'wb.Sheets(1).Cells(14, 29) = wb.Sheets(1).Cells(13, 26)
                                    'wb.Sheets(1).Cells(14, 30) = wb.Sheets(1).Cells(13, 26)
                                    
                                    'wb.Sheets(1).Cells(14, 31) = wb.Sheets(1).Cells(13, 32)
                                    'wb.Sheets(1).Cells(14, 32) = wb.Sheets(1).Cells(13, 32)
                                    'wb.Sheets(1).Cells(14, 33) = wb.Sheets(1).Cells(13, 32)
                                    'wb.Sheets(1).Cells(14, 34) = wb.Sheets(1).Cells(13, 32)
                                    'wb.Sheets(1).Cells(14, 35) = wb.Sheets(1).Cells(13, 32)
                                    
                                    'With wb.Sheets(1).Range(wb.Sheets(1).Range("B15:AI15"), wb.Sheets(1).Range("B15:AI15").End(xlDown))
                                    '    .HorizontalAlignment = xlGeneral
                                    '    .VerticalAlignment = xlCenter
                                    '    .WrapText = True
                                    '    .Orientation = 0
                                    '    .AddIndent = False
                                    '    .IndentLevel = 0
                                    '    .ShrinkToFit = False
                                    '    .ReadingOrder = xlContext
                                    '    .MergeCells = False
                                    'End With
                                    
                                    wb2.Sheets(1).Cells(10, 1) = "Cuenta:  Cuenta Corriente Soles 041-3000106330"
                                    premonto = "S/"
                                    
                                ElseIf wb.Sheets(1).Cells(8, 7) = "Corriente Dares 041-3000106347" Then
                                    wb2.Sheets(1).Cells(10, 1) = "Cuenta:  Cuenta Corriente Soles 041-3000106347"
                                    premonto = "US$"
                                End If
                                
                                While wb.Sheets(1).Cells(i, 2) <> ""
                                    'Fecha Op.
                                    wb2.Sheets(1).Cells(i + 3, 1) = wb.Sheets(1).Cells(i, 2)
                                    
                                    'Fecha Proc.
                                    wb2.Sheets(1).Cells(i + 3, 2) = wb.Sheets(1).Cells(i, 5)
                                    
                                    'Movimiento
                                    wb2.Sheets(1).Cells(i + 3, 3) = wb.Sheets(1).Cells(i, 13)
                                    
                                    'Detalle
                                    wb2.Sheets(1).Cells(i + 3, 4) = wb.Sheets(1).Cells(i, 17)
                                    
                                    'Nro. Operacion
                                    If wb.Sheets(1).Cells(i, 10) <> "" Then
                                        wb2.Sheets(1).Cells(i + 3, 5) = wb.Sheets(1).Cells(i, 10)
                                    Else
                                        wb2.Sheets(1).Cells(i + 3, 5) = "-"
                                    End If
                                    
                                    'Canal
                                    wb2.Sheets(1).Cells(i + 3, 6) = wb.Sheets(1).Cells(i, 22)
                                    
                                    'Cargo/Abono
                                    If wb.Sheets(1).Cells(i, 25) <> "" Then
                                        wb2.Sheets(1).Cells(i + 3, 8) = premonto & Format(wb.Sheets(1).Cells(i, 25), "#,##0.00")
                                    Else
                                        wb2.Sheets(1).Cells(i + 3, 9) = premonto & Format(wb.Sheets(1).Cells(i, 31), "#,##0.00")
                                    End If
                                    
                                    i = i + 1
                                Wend
                                
                                
                                'wb.Sheets(1).Cells(15, 2)
                                
                                wb.Close
                                Kill [DIRECTORIO] & file
                                file = "exportar" & "_" & Format(Now(), "hhmmss") & ".xlsx"
                                
                                Application.DisplayAlerts = False
                                wb2.SaveAs ([DIRECTORIO] & file)
                                Set wb = wb2
                                Set wb2 = Nothing
                                Application.DisplayAlerts = True
                            End If
                        End If
                        
                        'Ultimos 20 Movimientos BCP / Modificar la fecha de los movimientos despues del IT
                        If [FECHA_20BCP] = "SI" And nombre_corto = Me.Range("FECHA_20BCP").Offset(0, 1) Then
                            Dim bcpitf As Boolean
                            Dim bcpitf_fecha As Date
                            bcpitf = False
                            
                            While IsDate(rng(i, 1)) And i > 0
                                If bcpitf = True Then
                                    If rng(i, CInt(fecha_movimiento)) = Format(bcpitf_fecha, "DD/MM/YYYY") Then
                                        rng(i, CInt(fecha_movimiento)) = fechaDateStr(DateSerial(Year(bcpitf_fecha), Month(bcpitf_fecha), Day(bcpitf_fecha) + 1))
                                        rng(i, CInt(fecha_movimiento)).NumberFormat = "DD/MM/YYYY"
                                        rng(i, CInt(fecha_movimiento)).Interior.Color = 65535
                                    Else
                                        bcpitf = False
                                    End If
                                End If
                                If rng(i, CInt(glosa)) = "IMPUESTO ITF" Then
                                    bcpitf = True
                                    bcpitf_fecha = rng(i, CInt(fecha_movimiento))
                                    If rng(i, CInt(numero_operacion)) = "0000000" Then
                                        rng(i, CInt(numero_operacion)) = "0      0"
                                    End If
                                End If
                                i = i - 1
                            Wend
                            wb.Save
                        End If
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        ''''''''''''''''''''''''''''''''''''' Comisiones BBVA
                        If [OPE_COM_BBVA] = "SI" And nombre_corto = Me.Range("OPE_COM_BBVA").Offset(0, 1) Then
                            
                            Dim ultimaFila As Boolean
                            ultimaFila = True
                            ' Iguala el numero de operacion de la comision y el de la operacion correspondiente
                            '  Recorer de abajo a arriba los movimientos del BBVA
                        
                            'Glosa de comisiones por Transferencias a Terceros
                            'COMIS. TRASPASO O/P BANCA AUTOMATIC
                            
                            'Glosa de comisiones por Depositos
                            'COMISION DEPOSITO O/P
                            
                            While IsDate(rng(i, 1)) And i > 0
                                'Si la ultima Fila es una comision, su contraparte no se encuentra en el extracto procesado
                                If Not ultimaFila Then
                                    If rng(i, CInt(glosa)) = "COMIS. TRASPASO O/P BANCA AUTOMATIC" Or _
                                    rng(i, CInt(glosa)) = "COMISION DEPOSITO O/P" Then
                                        rng(i, CInt(numero_operacion)) = rng(i + 1, CInt(numero_operacion))
                                    End If
                                Else
                                    ultimaFila = False
                                    If rng(i, CInt(glosa)) = "COMIS. TRASPASO O/P BANCA AUTOMATIC" Or _
                                    rng(i, CInt(glosa)) = "COMISION DEPOSITO O/P" Then
                                        rng(i).EntireRow.Delete
                                    End If
                                End If
                                i = i - 1
                            Wend
                            Application.DisplayAlerts = False
                            wb.SaveAs [DIRECTORIO] & wb.Name, FileFormat:=xlWorkbookNormal
                            Application.DisplayAlerts = True
                        End If
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        '''Cambia el nro de cuenta del extracto de AELU DOLARES
                        If [CTA_SISGO_AELU] = "SI" And nombre_corto & " " & numero_corto = Me.Range("CTA_SISGO_AELU").Offset(0, 1) Then
                            wb.Sheets(1).Cells(rango_cuenta_row, rango_cuenta_col) = Me.Range("CTA_SISGO_AELU").Offset(0, 2)
                            wb.Save
                        End If
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                            
                        '''''''''' Remover ":" en la columna de Hora del Banco Pichincha para utilizarlo como Numero de Operacion
                        If [HORA_PICHINCHA] = "SI" And nombre_corto = Me.Range("HORA_PICHINCHA").Offset(0, 1) Then
                            i = 1
                            While rng(i, 1) <> ""
                                rng(i, CInt(hora_movimiento)) = Replace(rng(i, CInt(hora_movimiento)), ":", "")
                                i = i + 1
                            Wend
                            wb.Save
                        End If
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        '''''''''' Remover ultima fila del extracto de santander si es igual a 
                        If [ULT_FILA_SANTA] = "SI" And nombre_corto = Me.Range("ULT_FILA_SANTA").Offset(0, 1) Then
                            If wb.Sheets(1).Cells(inicio_row, inicio_col).End(xlDown) = "" Then
                                wb.Sheets(1).Cells(inicio_row, inicio_col).End(xlDown).EntireRow.Delete
                            End If
                        End If
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        'Compartir el xlsx
                        If [COMPARTIR_XLSX] = "SI" Then
                            wb.SaveAs Filename:=ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(Now(), "YYYY") & "\" & Format(Now(), "MM") & "\" & Format(Now(), "YYYY.MM.DD") & "\" & strUltimaCarga, FileFormat:=xlOpenXMLWorkbook
                        End If
                        
                        'Guarda el backup original en la ruta compartida
                        If [BACKUP_EXTRACTO] = "SI" Then
                            tmpSplit = Split(file, ".")
                            If UBound(tmpSplit) > 0 Then
                                FileCopy [DIRECTORIO] & "tmpCarga", ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(Now(), "YYYY") & "\" & Format(Now(), "MM") & "\" & Format(Now(), "YYYY.MM.DD") & "\Raw\" & tmpSplit(0) & "_" & Format(Now(), "HHMMSS") & "." & tmpSplit(1)
                            Else
                                FileCopy [DIRECTORIO] & "tmpCarga", ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format(Now(), "YYYY") & "\" & Format(Now(), "MM") & "\" & Format(Now(), "YYYY.MM.DD") & "\Raw\" & tmpSplit(0) & "_" & Format(Now(), "HHMMSS")
                            End If
                        End If
                        
                        'Copiar a Carpeta Input
                        If [A_INPUT] = "SI" Then
                            Call fso.CopyFile([DIRECTORIO] & file, ThisWorkbook.Sheets("L").Range("INPUT_PATH") & Split(file, ".")(0) & "_" & lastID & "." & Split(file, ".")(1), True)
                        End If
                        
                        'Guardar el xlsx para Carga Manual
                        If [CARGA_MANUAL] = "SI" Then
                            If Not dirExists([DIRECTORIO] & "MANUAL") Then
                                MkDir [DIRECTORIO] & "MANUAL"
                            End If
                            
                            '''''''''' Cambiar el formato de fecha del extrato del BANBIF; DD/MM/YY -> DD/MM/YYYY (Solamente en el caso de carga manual)
                            If [FORMATO_FECHA_BIF] = "SI" And nombre_corto = Me.Range("FORMATO_FECHA_BIF").Offset(0, 1) Then
                                Dim splitfecha As Variant
                                i = 1
                                Do While rng(i, CInt(fecha_movimiento)) <> ""
                                    splitfecha = Split(rng(i, CInt(fecha_movimiento)), "/")
                                    If UBound(splitfecha) <> 2 Then
                                        MsgBox "Error en Fecha"
                                        closeRS
                                        Exit Sub
                                    End If
                                    rng(i, CInt(fecha_movimiento)) = Format(DateSerial(splitfecha(2), splitfecha(1), splitfecha(0)), "YYYY-MM-DD")
                                    i = i + 1
                                Loop
                                rng(i, CInt(fecha_movimiento)).EntireColumn.NumberFormat = "DD/MM/YYYY"
                            End If
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            
                            wb.SaveAs Filename:=[DIRECTORIO] & "MANUAL\" & strUltimaCarga, FileFormat:=xlOpenXMLWorkbook
                        End If
                        
                        'Eliminar del Disco Local
                        If [BORRAR_LOCAL] = "SI" Then
                            cerrarWb wb
                            Kill [DIRECTORIO] & file
                        End If
                    ElseIf wb.Sheets(1).Cells(inicio_row, inicio_col) = "" Then
                        MsgBox "El extracto no contiene movimientos"
                        cerrarWb wb
                        If Not dirExists([DIRECTORIO] & "\NO_SE_CARGO") Then
                            MkDir [DIRECTORIO] & "\NO_SE_CARGO"
                        End If
                        Name [DIRECTORIO] & file As [DIRECTORIO] & "NO_SE_CARGO\" & Split(file, ".")(0) & correlativoArchivo([DIRECTORIO] & "NO_SE_CARGO\", CStr(file)) & "." & Split(file, ".")(1)
                    End If
                    
                    cerrarWb wb
                    
                Else
                    cerrarWb wb
                    If Not dirExists([DIRECTORIO] & "\NO_SE_CARGO") Then
                        MkDir [DIRECTORIO] & "\NO_SE_CARGO"
                    End If
                    Name [DIRECTORIO] & file As [DIRECTORIO] & "NO_SE_CARGO\" & Split(file, ".")(0) & correlativoArchivo([DIRECTORIO] & "NO_SE_CARGO\", CStr(file)) & "." & Split(file, ".")(1)
                End If
                
            End If
        End If
        file = Dir
    Wend
    
    On Error Resume Next
    Kill [DIRECTORIO] & "tmpCarga"
    On Error GoTo 0
    
    actualizarTodo
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub cargarMovimientos(wb As Workbook, _
                        id_banco As Integer, _
                        id_cuenta As Integer, _
                        inicio_col As Integer, _
                        inicio_row As Integer, _
                        fecha_movimiento As Integer, _
                        hora_movimiento As Integer, _
                        monto As Integer, _
                        monto2 As Integer, _
                        glosa As Integer, _
                        numero_operacion As Integer, _
                        itf As Integer, _
                        pre_num_ope As Integer)
                        
    Dim importe As Double
    Dim tmpGlosa As String
    Dim nroOpe As String
    Dim tipoOpe As Integer
    
    Dim duplicado As Boolean
    
    Dim j As Integer
    j = 0
    
    While wb.Sheets(1).Cells(inicio_row + j, inicio_col) <> ""
        
        'Evitar Duplicado
        duplicado = True
        While duplicado
            If pre_num_ope > 0 Then
                tmpGlosa = wb.Sheets(1).Cells(inicio_row + j, inicio_col + pre_num_ope - 1) & " - " & wb.Sheets(1).Cells(inicio_row + j, inicio_col + glosa - 1)
            Else
                tmpGlosa = wb.Sheets(1).Cells(inicio_row + j, inicio_col + glosa - 1)
            End If
            
            nroOpe = wb.Sheets(1).Cells(inicio_row + j, inicio_col + numero_operacion - 1)
            If nroOpe = "" Then
                nroOpe = "NULL"
            Else
                nroOpe = Replace("'" & nroOpe & "'", "", "")
            End If
        
            tmpGlosa = Replace(tmpGlosa, "'", "''")
            If tmpGlosa = "" Then
                tmpGlosa = "NULL"
            Else
                tmpGlosa = "'" & tmpGlosa & "'"
            End If
            
            If monto2 <> 0 Then
                If wb.Sheets(1).Cells(inicio_row + j, inicio_col + monto - 1) <> 0 And IsNumeric(Replace(Replace(Replace(wb.Sheets(1).Cells(inicio_row + j, inicio_col + monto - 1), "S/", ""), "US$", ""), "", "")) Then
                    importe = Replace(Replace(Replace(wb.Sheets(1).Cells(inicio_row + j, inicio_col + monto - 1), "S/", ""), "US$", ""), "", "") * (-1)
                Else
                    importe = Replace(Replace(Replace(wb.Sheets(1).Cells(inicio_row + j, inicio_col + monto2 - 1), "S/", ""), "US$", ""), "", "")
                End If
            Else
                importe = Replace(wb.Sheets(1).Cells(inicio_row + j, inicio_col + monto - 1), "", "")
            End If
            
            strSQL = "SELECT * FROM MOVIMIENTO WHERE ID_CUENTA_FK = " & id_cuenta & " AND GLOSA = " & tmpGlosa & " AND NUMERO_OPERACION = " & nroOpe & " AND FECHA_MOVIMIENTO = #" & Format(wb.Sheets(1).Cells(inicio_row + j, inicio_col + fecha_movimiento - 1), "YYYY-MM-DD") & "# AND MONTO = " & importe
            
            Set rs = Nothing
            On Error Resume Next
            rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
            On Error GoTo 0
            
            'Log del Error
            If cnn.Errors.count > 0 Then
                Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - cargarMovimientos", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
                cnn.Errors.Clear
                closeRS
                Exit Sub
            End If
            
            If rs.RecordCount = 0 Then
                duplicado = False
            Else
                j = j + 1
                If wb.Sheets(1).Cells(inicio_row + j, inicio_col + fecha_movimiento - 1) = "" Then
                    Exit Sub
                End If
            End If
        Wend
        
        ''''''Determinar tipo de movimientos // Default = GENERAL
        strSQL = "SELECT * FROM PARAM_MOVIMIENTO LEFT JOIN COINCIDENCIA ON PARAM_MOVIMIENTO.ID_COINCIDENCIA_FK = COINCIDENCIA.ID_COINCIDENCIA WHERE ID_BANCO_FK = " & id_banco & " ORDER BY ID_COINCIDENCIA"
        
        Set rs = Nothing
        On Error Resume Next
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        
        'Log del Error
        If cnn.Errors.count > 0 Then
            Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - cargarMovimientos", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
            cnn.Errors.Clear
            closeRS
            Exit Sub
        End If
        
        tipoOpe = 1
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                Select Case rs.Fields("NOMBRE_COINCIDENCIA")
                Case "EXACTA"
                    If wb.Sheets(1).Cells(inicio_row + j, inicio_col + glosa - 1) = rs.Fields("DESCRIPCION") Then
                        tipoOpe = rs.Fields("ID_TIPO_MOVIMIENTO_FK")
                        rs.MoveLast
                    End If
                Case "INICIAL"
                    If Left(tmpGlosa, Len(rs.Fields("DESCRIPCION"))) = rs.Fields("DESCRIPCION") Then
                        tipoOpe = rs.Fields("ID_TIPO_MOVIMIENTO_FK")
                        rs.MoveLast
                    End If
                Case "FINAL"
                    If Right(tmpGlosa, Len(rs.Fields("DESCRIPCION"))) = rs.Fields("DESCRIPCION") Then
                        tipoOpe = rs.Fields("ID_TIPO_MOVIMIENTO_FK")
                        rs.MoveLast
                    End If
                Case "CONTENIDA"
                    If InStr(1, tmpGlosa, rs.Fields("DESCRIPCION"), vbBinaryCompare) > 0 Then
                        tipoOpe = rs.Fields("ID_TIPO_MOVIMIENTO_FK")
                        rs.MoveLast
                    End If
                End Select
                rs.MoveNext
            Wend
        End If
        
        'Set rs = Nothing
        ''''''''''''''''''''''''''''''''''''''''''
        
        strSQL = "INSERT INTO MOVIMIENTO " & _
                "(ID_CUENTA_FK, MONTO, ID_TIPO_MOVIMIENTO_FK, GLOSA, NUMERO_OPERACION, FECHA_MOVIMIENTO, HORA_MOVIMIENTO, ANULADO, FECHA_GENERADO, USUARIO) VALUES " & _
                "(" & id_cuenta & ", " & _
                importe & ", " & _
                tipoOpe & ", " & _
                tmpGlosa & ", " & _
                nroOpe & ", " & _
                "#" & Format(wb.Sheets(1).Cells(inicio_row + j, inicio_col + fecha_movimiento - 1), "YYYY-MM-DD") & "#, "
        Dim tmpHora As String
        If hora_movimiento <> 0 Then
            tmpHora = wb.Sheets(1).Cells(inicio_row + j, inicio_col + hora_movimiento - 1)
            On Error Resume Next
            tmpHora = Format(wb.Sheets(1).Cells(inicio_row + j, inicio_col + hora_movimiento - 1), "HH:MM:SS")
            On Error GoTo 0
            strSQL = strSQL & "'" & tmpHora & "', "
        Else
            strSQL = strSQL & "NULL, "
        End If
        
        strSQL = strSQL & "FALSE, " & _
                "#" & Format(Now(), "YYYY-MM-DD HH:MM:SS") & "#, " & _
                "'" & Application.UserName & "')"
        
        On Error Resume Next
        cnn.Execute strSQL
        On Error GoTo 0
        
        'Log del Error
        If cnn.Errors.count > 0 Then
            Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - cargarMovimientos (1)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
            cnn.Errors.Clear
            closeRS
            Exit Sub
        End If
        
        If itf > 0 Then
            If Replace(wb.Sheets(1).Cells(inicio_row + j, inicio_col + itf - 1), "", "") <> 0 Then
                 strSQL = "INSERT INTO MOVIMIENTO " & _
                    "(ID_CUENTA_FK, MONTO, ID_TIPO_MOVIMIENTO_FK, GLOSA, NUMERO_OPERACION, FECHA_MOVIMIENTO, HORA_MOVIMIENTO, ANULADO, FECHA_GENERADO, USUARIO) VALUES " & _
                    "(" & id_cuenta & ", " & _
                    Replace(wb.Sheets(1).Cells(inicio_row + j, inicio_col + itf - 1), "", "") & ", " & _
                    "2, " & _
                    tmpGlosa & ", " & _
                    nroOpe & ", " & _
                    "#" & Format(wb.Sheets(1).Cells(inicio_row + j, inicio_col + fecha_movimiento - 1), "YYYY-MM-DD") & "#, "
                    
                If hora_movimiento <> 0 Then
                    strSQL = strSQL & "'" & wb.Sheets(1).Cells(inicio_row + j, inicio_col + hora_movimiento - 1) & "', "
                Else
                    strSQL = strSQL & "NULL, "
                End If
                
                strSQL = strSQL & "FALSE, " & _
                        "#" & Format(Now(), "YYYY-MM-DD HH:MM:SS") & "#, " & _
                        "'" & Application.UserName & "')"
                
                On Error Resume Next
                cnn.Execute strSQL
                On Error GoTo 0
                
                'Log del Error
                If cnn.Errors.count > 0 Then
                    Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - cargarMovimientos (1)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
                    cnn.Errors.Clear
                    closeRS
                    Exit Sub
                End If
            End If
        End If
        j = j + 1
    Wend
End Sub

Function correlativoArchivo(ruta As String, fichero As String) As String
    correlativoArchivo = ""
    Dim j As Integer: j = 0
    
    While IsFile(ruta & Split(fichero, ".")(0) & correlativoArchivo & "." & Split(fichero, ".")(1))
        j = j + 1
        correlativoArchivo = " (" & j & ")"
    Wend
End Function

Sub cerrarWb(wb As Workbook)
    If Not wb Is Nothing Then
        Application.DisplayAlerts = False
        On Error Resume Next
        wb.Close
        On Error GoTo 0
        Set wb = Nothing
        Application.DisplayAlerts = True
    End If
End Sub

Function IsFile(ByVal fName As String) As Boolean
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function

Private Sub registrarCarga( _
    id_param_cuenta As Integer, _
    nombre_corto As String, _
    numero_corto As String, _
    fechaMod As Date)
    
    Dim correlativo As String
    
    Dim rng As Range
    Dim i As Integer
    
    strSQL = "SELECT MAX(CORRELATIVO) FROM CARGA_XLSX WHERE ID_PARAM_CUENTA_FK = " & id_param_cuenta & " AND FORMAT(FECHA,'YYYY-MM-DD') = '" & Format(fechaMod, "YYYY-MM-DD") & "' GROUP BY ID_PARAM_CUENTA_FK"
    
    Set rs2 = Nothing
    On Error Resume Next
    rs2.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Hoja1.cargar = False
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - registrarCarga", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    If rs2.RecordCount > 0 Then
    
        correlativo = rs2.Fields(0) + 1
        If Not IsNumeric(correlativo) Then
            correlativo = 1
        End If
    Else
        correlativo = 1
    End If
    
    If Len(correlativo) = 1 Then
        correlativo = "0" & correlativo
    End If
    
    strSQL = "INSERT INTO CARGA_XLSX (ID_PARAM_CUENTA_FK, CORRELATIVO, FECHA, FECHA_GEN, USUARIO, EN_INPUT, FECHA_SALIDA) VALUES (" & id_param_cuenta & ", " & correlativo & ", #" & Format(fechaMod, "YYYY-MM-DD HH:MM:SS") & "#, #" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "#, '" & Application.UserName & "', TRUE, #" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "#)"
    
    On Error Resume Next
    cnn.Execute strSQL
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Hoja1.cargar = False
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - registrarCarga (1)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    Hoja1.cargar = True
    strUltimaCarga = nombre_corto & " " & numero_corto & " " & correlativo & ".xlsx"
    
    strSQL = "SELECT @@IDENTITY"
    
    Set rs2 = Nothing
    On Error Resume Next
    rs2.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    If cnn.Errors.count > 0 Then
        Hoja1.cargar = False
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - registrarCarga (2)", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
        cnn.Errors.Clear
        closeRS
        Exit Sub
    End If
    
    If rs2.RecordCount > 0 Then
        lastID = rs2.Fields(0)
    End If
    
End Sub

Private Sub btHoyRAW_Click()
    Dim strPath As String
    
    strPath = ThisWorkbook.Sheets("L").Range("XLSX_PATH")
    
    If dirExists(strPath & Format(Now(), "YYYY") & "\" & Format(Now(), "MM") & "\" & Format(Now(), "YYYY.MM.DD") & "\RAW\") Then
        Shell "C:\WINDOWS\explorer.exe """ & strPath & Format(Now(), "YYYY") & "\" & Format(Now(), "MM") & "\" & Format(Now(), "YYYY.MM.DD") & "\RAW\", vbNormalFocus
    End If
    
    closeRS
End Sub

Private Sub btHoyOutput_Click()
    Dim outputPath As String
    
    outputPath = ThisWorkbook.Sheets("L").Range("OUTPUT_PATH")
    
    If dirExists(outputPath & Format(Now(), "DD-MM-YYYY") & "\") Then
        Shell "C:\WINDOWS\explorer.exe """ & outputPath & Format(Now(), "DD-MM-YYYY") & "\", vbNormalFocus
    End If
    
    closeRS
End Sub

Private Sub btHoyXLSX_Click()
    Dim strPath As String
    
    strPath = ThisWorkbook.Sheets("L").Range("XLSX_PATH")
    
    If dirExists(strPath & Format(Now(), "YYYY") & "\" & Format(Now(), "MM") & "\" & Format(Now(), "YYYY.MM.DD")) Then
        Shell "C:\WINDOWS\explorer.exe """ & strPath & Format(Now(), "YYYY") & "\" & Format(Now(), "MM") & "\" & Format(Now(), "YYYY.MM.DD") & "", vbNormalFocus
    End If
    
    closeRS
End Sub


Private Sub btMovimientos_Click()
    [TMP_FECHA] = ""
    [TMP_FECHA2] = ""
    
    frmCalendario2.Show
    
    If [TMP_FECHA] <> "" And [TMP_FECHA2] <> "" Then
        If IsDate([TMP_FECHA]) And IsDate([TMP_FECHA2]) Then
            Hoja4.Activate
            
            Hoja4.Range("4:4").AutoFilter Field:=1, Criteria1:="<>Cualquier Cosa"
            Hoja4.Range("4:4").AutoFilter
            Hoja4.Range("4:4").AutoFilter
            
            [FECHA_INICIO] = [TMP_FECHA]
            [FECHA_FIN] = [TMP_FECHA2]
            
            strSQL = "SELECT NOMBRE_BANCO, NUMERO_CUENTA_SISGO, NUMERO_OPERACION, IIF(MONTO < 0, MONTO * (-1), 0), IIF(MONTO > 0, MONTO, 0), NOMBRE_TIPO_MOVIMIENTO, GLOSA, FECHA_MOVIMIENTO, USUARIO, FECHA_GENERADO " & _
            " FROM (((MOVIMIENTO LEFT JOIN CUENTA ON CUENTA.ID_CUENTA = MOVIMIENTO.ID_CUENTA_FK) " & _
                "LEFT JOIN BANCO ON BANCO.ID_BANCO = CUENTA.ID_BANCO_FK) " & _
                "LEFT JOIN TIPO_MOVIMIENTO ON TIPO_MOVIMIENTO.ID_TIPO_MOVIMIENTO = MOVIMIENTO.ID_TIPO_MOVIMIENTO_FK) " & _
                    "WHERE FECHA_MOVIMIENTO >= #" & Format([FECHA_INICIO], "YYYY-MM-DD") & "# AND FECHA_MOVIMIENTO <= #" & Format([FECHA_FIN], "YYYY-MM-DD") & "#"
            
            OpenDB
            
            On Error Resume Next
            rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
            On Error GoTo 0
            
            'Log del Error
            If cnn.Errors.count > 0 Then
                Hoja1.cargar = False
                Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btMovimientos_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
                cnn.Errors.Clear
                closeRS
                Exit Sub
            End If
            Hoja4.Range(Hoja4.Range("dataSetMovimientos"), Hoja4.Range("dataSetMovimientos").End(xlDown)).ClearContents
            
            If rs.RecordCount > 0 Then
                Hoja4.Range("dataSetMovimientos").CopyFromRecordset rs
            End If
            
            closeRS
            
        End If
    End If
End Sub

Private Sub btParametrizarCta_Click()
    frmParamCuenta.Show
End Sub

Private Sub btParametrizarMov_Click()
    frmParamMov.Show
End Sub

Private Sub btParametrizarSaldo_Click()
    frmParamSaldo.Show
End Sub

Public Sub btPichinchaAr_Click()

    Dim X As Variant
    Dim Path As String
    Path = ActiveWorkbook.Path & Application.PathSeparator & "Pichincha Arturo\Tarjeta.exe"
    Application.WindowState = xlMinimized
    
    On Error Resume Next
    X = Shell(Path, vbNormalFocus)
End Sub

Function GetFolder(strPath As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Public Function BrowseForFolder(Optional OpenAt As Variant) As Variant
    Dim ShellApp As Object
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Por favor seleccione la carpeta", NO_OPTIONS, OpenAt)
    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0

    Set ShellApp = Nothing
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select
    BrowseForFolder = BrowseForFolder & "\"
    Exit Function

Invalid:
    BrowseForFolder = False
End Function

Public Sub verificarVersion()
    Dim fso As FileSystemObject
    Dim txtStream As TextStream
    
    Dim ultimaVersion As String
    
    Dim verPath As String
    
    verPath = ThisWorkbook.Sheets("L").Range("XLSX_VERSION")
        
    Set fso = New FileSystemObject
    Set txtStream = fso.OpenTextFile(verPath, ForReading, False)
    
    ultimaVersion = txtStream.ReadLine
    
    If [VERSIONXLSX] <> ultimaVersion Then
        Range("VERSIONXLSX").Interior.Pattern = xlSolid
        Range("VERSIONXLSX").Interior.Color = 49407
        Range("VERSIONXLSX").Offset(0, 1).Interior.Pattern = xlSolid
        Range("VERSIONXLSX").Offset(0, 1).Interior.Color = 49407
        Range("VERSIONXLSX").Offset(0, 2).Interior.Pattern = xlSolid
        Range("VERSIONXLSX").Offset(0, 2).Interior.Color = 49407
        If Left(Right(ultimaVersion, 3), 1) < Left(Right([VERSIONXLSX], 3), 1) Then
            Range("VERSIONXLSX").Offset(0, 1).Value = "Version Desarrollo"
        Else
            If Right(ultimaVersion, 1) < Right([VERSIONXLSX], 1) And Left(Right(ultimaVersion, 3), 1) = Left(Right([VERSIONXLSX], 3), 1) Then
                Range("VERSIONXLSX").Offset(0, 1).Value = "Version Desarrollo"
            Else
                Range("VERSIONXLSX").Offset(0, 1).Value = "No es la ultima Version"
            End If
        End If
    Else
        Range("VERSIONXLSX").Interior.Pattern = xlNone
        Range("VERSIONXLSX").Offset(0, 1).Interior.Pattern = xlNone
        Range("VERSIONXLSX").Offset(0, 2).Interior.Pattern = xlNone
        Range("VERSIONXLSX").Offset(0, 1).Value = ""
    End If
End Sub

Public Function dirExists(s_directory As String) As Boolean

    Set OFSO = CreateObject("Scripting.FileSystemObject")
    dirExists = OFSO.FolderExists(s_directory)
    Set OFSO = Nothing
    
End Function

Private Sub btRegistrarHistorico_Click()
    Dim i As Integer
    Dim id_cuenta As Integer
    Dim id_banco As Integer
    Dim inicio_col As Integer
    Dim inicio_row As Integer
    Dim strDir As String
    Dim wb As Workbook
    Dim fecha_movimiento As Integer
    Dim monto As Integer
    Dim monto2 As Integer
    Dim glosa As Integer
    Dim numero_operacion As Integer
    Dim hora_movimiento As Integer
    Dim itf As Integer
    Dim pre_num_ope As Integer
    
    Dim ctas_limpiar As String
    
    i = 1
        
    OpenDB
    While Me.Range("dataSet")(i, 1) <> ""
        If Me.Range("dataSet")(i, 1).Interior.Pattern = xlSolid Then
            strDir = ThisWorkbook.Sheets("L").Range("XLSX_PATH") & Format([FECHA_HISTORICO], "YYYY") & "\" & Format([FECHA_HISTORICO], "MM") & "\" & Format([FECHA_HISTORICO], "YYYY.MM.DD") & "\" & Me.Range("dataSet")(i, 1).Value & ".xlsx"
            If Dir(strDir) <> "" Then
                Workbooks.Open (strDir)
                Set wb = Workbooks(Me.Range("dataSet")(i, 1).Value & ".xlsx")
            Else
                MsgBox "Error: No se encontr・el XLSX" & vbCrLf & " >> Ruta: " & strDir
            End If
            
            strSQL = "SELECT * FROM PARAM_CUENTA LEFT JOIN PARAM_CAMPO ON PARAM_CUENTA.ID_PARAM_CUENTA = PARAM_CAMPO.ID_PARAM_CUENTA_FK WHERE INSTR('" & Me.Range("dataSet")(i, 1).Value & "', NOMBRE_CORTO + ' ' + NUMERO_CORTO)"
            Set rs = Nothing
            On Error Resume Next
            rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
            On Error GoTo 0
            
            'Log del Error
            If cnn.Errors.count > 0 Then
                Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btRegistrarHistorico_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
                cnn.Errors.Clear
                closeRS
                Exit Sub
            End If
            
            If rs.RecordCount > 0 Then
                id_cuenta = rs.Fields("ID_CUENTA_FK")
                id_banco = rs.Fields("ID_BANCO_FK")
                inicio_col = rs.Fields("INICIO_COL")
                inicio_row = rs.Fields("INICIO_ROW")
                
                fecha_movimiento = 0    'ID_CAMPO: 1
                monto = 0               'ID_CAMPO: 2
                monto2 = 0              'ID_CAMPO: 3
                glosa = 0               'ID_CAMPO: 4
                numero_operacion = 0    'ID_CAMPO: 5
                hora_movimiento = 0     'ID_CAMPO: 6
                itf = 0                 'ID_CAMPO: 7
                pre_num_ope = 0         'ID_CAMPO: 8
                
                Do While Not rs.EOF
                    Select Case rs.Fields("ID_CAMPO_FK")
                    Case 1
                        fecha_movimiento = rs.Fields("COLUMNA")
                    Case 2
                        monto = rs.Fields("COLUMNA")
                    Case 3
                        monto2 = rs.Fields("COLUMNA")
                    Case 4
                        glosa = rs.Fields("COLUMNA")
                    Case 5
                        numero_operacion = rs.Fields("COLUMNA")
                    Case 6
                        hora_movimiento = rs.Fields("COLUMNA")
                    Case 7
                        itf = rs.Fields("COLUMNA")
                    Case 8
                        pre_num_ope = rs.Fields("COLUMNA")
                    End Select
                    
                    rs.MoveNext
                Loop
                
                cargarMovimientos wb, id_banco, id_cuenta, inicio_col, inicio_row, fecha_movimiento, hora_movimiento, monto, monto2, glosa, numero_operacion, itf, pre_num_ope
                
            End If
            
            Me.Range("dataSet")(i, 1).Interior.Pattern = xlNone
            cerrarWb wb
        End If
        i = i + 1
    Wend
    closeRS
End Sub

Public Sub btReporteRetraso_Click()
    [TMP_FECHA] = ""
    [TMP_FECHA2] = ""
    
    frmCalendario2.Show
    
    If [TMP_FECHA] <> "" And [TMP_FECHA2] <> "" Then
        Dim Lastrow As Integer
        strSQL = "SELECT ID_CARGA_XLSX, NOMBRE_CORTO + ' ' + NUMERO_CORTO + ' ' + CSTR(CORRELATIVO), FECHA_GEN, FECHA_SALIDA, FECHA_SALIDA - FECHA_GEN FROM CARGA_XLSX LEFT JOIN PARAM_CUENTA ON PARAM_CUENTA.ID_PARAM_CUENTA = CARGA_XLSX.ID_PARAM_CUENTA_FK WHERE FORMAT(FECHA,'YYYY-MM-DD') >= '" & fechaDateStr([TMP_FECHA]) & "' AND FORMAT(FECHA,'YYYY-MM-DD') <= '" & fechaDateStr([TMP_FECHA2]) & "' ORDER BY FECHA_SALIDA ASC, ID_CARGA_XLSX ASC"
        OpenDB
        
        On Error Resume Next
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        
        'Log del Error
        If cnn.Errors.count > 0 Then
            Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btRegistrarHistorico_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
            cnn.Errors.Clear
            closeRS
            Exit Sub
        End If
        
        Hoja9.Range(Hoja9.Range("dataSetReporte"), Hoja9.Range("dataSetReporte").End(xlDown)).ClearContents
        
        If rs.RecordCount > 0 Then
            Hoja9.Activate
            Hoja9.Range("dataSetReporte").CopyFromRecordset rs
            'Hoja9.Range("dataSetReporte").Offset(0, 5)(1, 1).FormulaR1C1 = "=RC[-2]-R[-1]C[-2]"
            'Lastrow = Hoja9.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            'Hoja9.Range("F2").AutoFill Hoja9.Range("F2:F" & Lastrow)
            
            Hoja9.Range("dataSetReporte").Offset(0, 2).EntireColumn.NumberFormat = "hh:mm:ss am/pm"
            Hoja9.Range("dataSetReporte").Offset(0, 3).EntireColumn.NumberFormat = "hh:mm:ss am/pm"
            Hoja9.Range("dataSetReporte").Offset(0, 4).EntireColumn.NumberFormat = "hh:mm:ss"
            Hoja9.Range("dataSetReporte").Offset(0, 5).EntireColumn.NumberFormat = "hh:mm:ss"
        End If
    End If
    
    closeRS
End Sub

Private Sub btSetDirectorio_Click()
    Dim Ret

    Ret = BrowseForFolder()
    
    If Ret <> False Then
        [DIRECTORIO] = Ret
        strSQL = "UPDATE ALIAS SET RUTA_DESCARGA = '" & [DIRECTORIO] & "' WHERE NOMBRE = '" & Application.UserName & "'"
        OpenDB
        On Error Resume Next
        cnn.Execute strSQL
        On Error GoTo 0
        
        'Log del Error
        If cnn.Errors.count > 0 Then
            Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btSetDirectorio_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
            cnn.Errors.Clear
            closeRS
            Exit Sub
        End If
        
        closeRS
    End If
    
End Sub

Public Sub btXLSXgenerados_Click()
    Sheets("CARGAS").Range("TMP_FECHA") = Format(Now(), "YYYY-MM-DD")
    busqXLSX.Show
End Sub

Function FileLastModified(strFullFileName As String)
    Dim fs As Object, f As Object, s As String
   
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(strFullFileName)
   
    s = f.DateLastModified
    FileLastModified = s
    
    Set fs = Nothing: Set f = Nothing
    
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim sw As Boolean
    sw = False
    For Each estr In arr
        If InStr(1, stringToBeFound, estr, vbTextCompare) Then
            sw = True
            Exit For
        End If
    Next estr
    IsInArray = sw
End Function

Private Sub cbActualizar_Click()
    If Hoja1.cbActualizar Then
        proxActualizacion = Now + TimeValue("00:" & [INTERVALO] & ":00")
        Application.OnTime proxActualizacion, Procedure:="actualizarTodo"
        [PROXIMA_ACTUALIZACION] = proxActualizacion
    Else
        [PROXIMA_ACTUALIZACION] = ""
        On Error Resume Next
        Application.OnTime proxActualizacion, Procedure:="actualizarTodo", Schedule:=False
        On Error GoTo 0
    End If
End Sub

