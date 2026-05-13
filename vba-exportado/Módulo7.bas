Attribute VB_Name = "Módulo7"
Sub CrearTablaDeVenta()

    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim dictPeriodos As Object
    Dim UltFila As Long, i As Long
    Dim p As String, it As String
    Dim cantidad As Double
    Dim filaOut As Long, colOut As Long
    Dim clavePeriodo As Variant
    Dim peso As Double
    Dim itemsOrdenados As Variant
    Dim item As Variant
    Dim filaInicio As Long: filaInicio = 2
    Dim colTotal As Long

    Set dictPeriodos = CreateObject("Scripting.Dictionary")
    Set wsOrigen = ThisWorkbook.Sheets("venta de pollo")
    UltFila = wsOrigen.Cells(wsOrigen.Rows.count, 1).End(xlUp).Row

    itemsOrdenados = Array("Alitas", "Filete", "Ruster", "POP", "P. Crispy", "P. Original", "Strips")

    ' Crear o limpiar hoja destino
    On Error Resume Next
    Set wsDestino = Sheets("Tabla de venta")
    If wsDestino Is Nothing Then
        wsDestino.Name = "Tabla de venta"
    Else
        wsDestino.Cells.Clear
    End If
    On Error GoTo 0

    ' Encabezado
    wsDestino.Cells(1, 1).Value = fechaBase
    wsDestino.Cells(1, 2).Value = "% Ajuste"

For i = 2 To UltFila
    p = Trim(wsOrigen.Cells(i, 1).Value)
    If Not dictPeriodos.Exists(p) Then
        dictPeriodos.Add p, dictPeriodos.count + 3
        ' >>> aquí:
        If InStr(p, " -") > 0 Then
            wsDestino.Cells(1, dictPeriodos(p)).Value = Trim(Split(p, "-")(0))
        Else
            wsDestino.Cells(1, dictPeriodos(p)).Value = p
        End If
    End If
Next i


    colTotal = dictPeriodos.count + 3
    wsDestino.Cells(1, colTotal).Value = "Total"

    ' Escribir los ítems
    For i = LBound(itemsOrdenados) To UBound(itemsOrdenados)
        wsDestino.Cells(i + filaInicio, 1).Value = itemsOrdenados(i)
        wsDestino.Cells(i + filaInicio, 2).Value = 0
        wsDestino.Cells(i + filaInicio, 2).NumberFormat = "0.00%"
    Next i

    ' Llenar valores
    For i = 2 To UltFila
        p = Trim(wsOrigen.Cells(i, 1).Value)
        it = Trim(wsOrigen.Cells(i, 2).Value)
        cantidad = wsOrigen.Cells(i, 3).Value / 4

        ' Renombrar ítems
        Select Case UCase(it)
            Case "NUEVO POPCORN MEDIANO TERMINADO"
                it = "POP": peso = Round(cantidad * 0.18)
            Case "NUEVO POPCORN GRANDE TERMINADO"
                it = "POP": peso = Round(cantidad * 0.2)
            Case "RUSTER TERMINADO"
                it = "Ruster": peso = Round(cantidad)
            Case "ALITA TERMINADO"
                it = "Alitas": peso = Round(cantidad)
            Case "STRIP TERMINADO"
                it = "Strips": peso = Round(cantidad)
            Case "PIEZA CRISPY TERMINADA"
                it = "P. Crispy": peso = Round(cantidad)
            Case "PIEZA ORIGINAL TERMINADA"
                it = "P. Original": peso = Round(cantidad)
            Case Else
                If Left(UCase(it), 3) = "KCS" Then
                    it = "Filete"
                End If
                peso = Round(cantidad)
        End Select

        ' Buscar fila correcta
        For filaOut = filaInicio To filaInicio + UBound(itemsOrdenados)
            If wsDestino.Cells(filaOut, 1).Value = it Then
                colOut = dictPeriodos(p)
                With wsDestino.Cells(filaOut, colOut)
                    If IsNumeric(.Value) Then
                        .Value = .Value + peso
                    Else
                        .Value = peso
                    End If
                End With
                Exit For
            End If
        Next filaOut
    Next i

    ' Ajuste de porcentaje
    Call AplicarAjustePorcentaje(wsDestino, filaInicio, filaInicio + UBound(itemsOrdenados), 3, colTotal - 1)

    ' Calcular total por fila
    Dim suma As Double
    Dim valorCelda As Variant
    Dim partes() As String
    Dim divisor As Double
    Dim cajones As Double

    Call CalcularTotalesConCajones(wsDestino, filaInicio, filaInicio + UBound(itemsOrdenados), 3, colTotal)
    'Call DividirCeldasYAgregarTexto(wsDestino, filaInicio, filaInicio + UBound(itemsOrdenados), 3, colTotal - 1)

    ' Formato elegante y KFC
    With wsDestino.Range("A1", wsDestino.Cells(filaInicio + UBound(itemsOrdenados), colTotal))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium ' Bordes más gruesos
    End With

    ' Encabezado solo hasta el final de la tabla
    With wsDestino.Range(wsDestino.Cells(1, 1), wsDestino.Cells(1, colTotal))
        .Interior.Color = RGB(207, 18, 39) ' Rojo KFC
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .RowHeight = 30
    End With

    ' Fondo blanco para los datos
    With wsDestino.Range(wsDestino.Cells(filaInicio, 1), wsDestino.Cells(filaInicio + UBound(itemsOrdenados), colTotal))
        .Interior.Color = RGB(255, 255, 255)
    End With

    ' Fondo gris claro para "Total"
    With wsDestino.Range(wsDestino.Cells(filaInicio, colTotal), wsDestino.Cells(filaInicio + UBound(itemsOrdenados), colTotal))
        .Interior.Color = RGB(245, 245, 245)
        .Font.Bold = True
    End With

    ' Darle estilo especial a los nombres de los ítems
    With wsDestino.Range(wsDestino.Cells(filaInicio, 1), wsDestino.Cells(filaInicio + UBound(itemsOrdenados), 1))
        .Interior.Color = RGB(230, 230, 230) ' Gris clarito para distinguir
        .Font.Bold = True
    End With

    ' Aumentar altura de todas las filas
    For i = 1 To filaInicio + UBound(itemsOrdenados)
        wsDestino.Rows(i).RowHeight = 24
    Next i

    Application.EnableEvents = True
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case "Tabla de venta", "Combos pollo", "Combos no encontrados"
                ' Asegurarse de que estas hojas estén visibles
                ws.Visible = xlSheetVisible
            Case Else
                ' Ocultar todas las demás
                ws.Visible = xlSheetHidden
        End Select
    Next ws
    
    ' Activar la hoja principal de resultados para el usuario.
    ThisWorkbook.Sheets("Tabla de venta").Activate
End Sub


Sub AplicarAjustePorcentaje(ws As Worksheet, filaInicio As Long, filaFin As Long, colInicioDatos As Long, colFinDatos As Long)
    Dim fila As Long, col As Long
    Dim ajuste As Double
    For fila = filaInicio To filaFin
        ajuste = ws.Cells(fila, 2).Value
        If IsNumeric(ajuste) Then
            For col = colInicioDatos To colFinDatos
                If IsNumeric(ws.Cells(fila, col).Value) Then
                    ws.Cells(fila, col).Value = ws.Cells(fila, col).Value * (1 + ajuste)
                End If
            Next col
        End If
    Next fila
End Sub

Sub DividirCeldasYAgregarTexto(ws As Worksheet, filaInicio As Long, filaFin As Long, colInicio As Long, colFin As Long)
    Dim fila As Long, col As Long
    Dim valor As Double, divisor As Double
    Dim resultado As Long, texto As String
    For fila = filaInicio To filaFin
        divisor = ObtenerDivisor(ws.Cells(fila, 1).Value)
        For col = colInicio To colFin
            If IsNumeric(ws.Cells(fila, col).Value) Then
                valor = ws.Cells(fila, col).Value
                resultado = Round(valor / divisor, 0)
                texto = ObtenerTexto(resultado)
                If texto <> "" Then
                    ws.Cells(fila, col).Value = valor & " (" & texto & ")"
                End If
            End If
        Next col
    Next fila
End Sub

Function ObtenerDivisor(item As String) As Double
    Select Case UCase(item)
        Case "STRIPS": ObtenerDivisor = 18
        Case "P. ORIGINAL", "P. CRISPY": ObtenerDivisor = 18
        Case "RUSTER": ObtenerDivisor = 15
        Case "ALITAS": ObtenerDivisor = 24
        Case "POP", "FILETE": ObtenerDivisor = 18
        Case Else: ObtenerDivisor = 1
    End Select
End Function

Function ObtenerTexto(resultado As Long) As String
    Select Case resultado
        Case 1: ObtenerTexto = "R"
        Case 2: ObtenerTexto = "B"
        Case 3: ObtenerTexto = "B+R"
        Case 4, 5: ObtenerTexto = "T+R"
        Case 6: ObtenerTexto = "T+B"
        Case 7: ObtenerTexto = "T+3R"
        Case 8: ObtenerTexto = "2T"
        Case 9: ObtenerTexto = "2T+R"
        Case 10: ObtenerTexto = "2T+B"
        Case 11: ObtenerTexto = "T+3R"
        Case Else: ObtenerTexto = ""
    End Select
End Function

Sub AgregarTextoEnTotal(ws As Worksheet, filaInicio As Long, filaFin As Long, colTotal As Long)
    Dim fila As Long, valor As Double, divisor As Double, texto As String
    For fila = filaInicio To filaFin
        valor = ws.Cells(fila, colTotal).Value
        Select Case ws.Cells(fila, 1).Value
            Case "Alitas": divisor = 336
            Case "Filete": divisor = 168
            Case "Ruster": divisor = 300
            Case "P. Crispy", "P. Original": divisor = 144
            Case "POP": divisor = 22
            Case "Strips": divisor = 360
            Case Else: divisor = 1
        End Select
        If divisor <> 0 Then
            texto = Format(valor / divisor, "0.0") & " cajones"
            ws.Cells(fila, colTotal).Value = valor & " (" & texto & ")"
        End If
    Next fila
End Sub

Sub FormatearTablaKFC(ws As Worksheet, filaInicio As Long, filaFin As Long, colTotal As Long)
    Dim rng As Range, i As Long
    
    ' Encabezados
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, colTotal))
        .Interior.Color = RGB(178, 34, 34)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Toda la tabla
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(filaFin, colTotal))
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = vbBlack

    ' Alternar filas
    For i = filaInicio To filaFin
        If (i Mod 2) = 0 Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, colTotal)).Interior.Color = RGB(255, 255, 255)
        Else
            ws.Range(ws.Cells(i, 1), ws.Cells(i, colTotal)).Interior.Color = RGB(242, 242, 242)
        End If
    Next i

    ' Resaltar la columna Total
    With ws.Range(ws.Cells(filaInicio, colTotal), ws.Cells(filaFin, colTotal))
        .Font.Bold = True
        .Interior.Color = RGB(245, 245, 245)
    End With

    ' Nombres de productos
    With ws.Range(ws.Cells(filaInicio, 1), ws.Cells(filaFin, 1))
        .Font.Bold = True
    End With

    ' Fecha grande y roja
    With ws.Cells(1, 1)
        .Font.Size = 12
        .Font.Color = RGB(178, 34, 34)
        .Font.Bold = True
    End With

    ws.Columns.AutoFit
End Sub

' Ubicación: Módulo7
' Reemplazar la subrutina original con esta.

Sub CalcularTotalesConCajones(ws As Worksheet, filaInicio As Long, filaFin As Long, colInicioDatos As Long, colTotal As Long)
    Dim i As Long, colOut As Long
    Dim suma As Double

    For i = filaInicio To filaFin
        ' Simplemente calcula la suma de la fila.
        suma = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(i, colInicioDatos), ws.Cells(i, colTotal - 1)))
        
        ' Y la coloca como un número puro en la columna Total.
        ws.Cells(i, colTotal).Value = suma
    Next i
End Sub

Function ObtenerDivisorTotal(item As String) As Double
    Select Case UCase(item)
        Case "ALITAS": ObtenerDivisorTotal = 336
        Case "FILETE": ObtenerDivisorTotal = 168
        Case "RUSTER": ObtenerDivisorTotal = 300
        Case "P. CRISPY", "P. ORIGINAL": ObtenerDivisorTotal = 144
        Case "POP": ObtenerDivisorTotal = 22
        Case "STRIPS": ObtenerDivisorTotal = 360
        Case Else: ObtenerDivisorTotal = 1
    End Select
End Function





