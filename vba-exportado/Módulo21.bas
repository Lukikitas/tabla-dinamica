Attribute VB_Name = "Módulo21"
Option Explicit

' Módulo 21: Procesamiento y Creación de Tabla
' Depuración: registra únicamente el cálculo de:
' - "PIEZA CRISPY TERMINADA"
' - "PIEZA ORIGINAL TERMINADA"

Sub ProcesarVentaDiariaYCrearTabla()

    ' Declaración de variables
    Dim dicPollo As Object, dicVentaPollo As Object, dicErrores As Object
    Dim wsVenta As Worksheet, wsCombo As Worksheet, wsDatos As Worksheet, wsErrores As Worksheet
    Dim wsDebug As Worksheet
    Dim hora As Variant, producto As String, cantidadVenta As Long
    Dim comboNombre As String, nombreItemPollo As Variant, cantidadCombo As Long
    Dim clave As Variant, peso As Double, filaOut As Long, colOut As Long
    Dim i As Long, j As Long, fila As Long, ultFilaDatos As Long, ultFilaCombo As Long, filaError As Long
    Dim itemsOrdenados As Variant, colTotal As Long
    Dim dictPeriodos As Object
    Dim ignoreList As Object, itemToIgnore As Variant
    Dim filaDebug As Long
    Dim itemUpper As String

    ' --- PASO 1: ESTABLECER HOJAS DE TRABAJO ---
    On Error Resume Next
    Set wsDatos = ThisWorkbook.Sheets("Venta por hora")
    Set wsCombo = ThisWorkbook.Sheets("Combos pollo")
    On Error GoTo 0
    
    If wsDatos Is Nothing Or wsCombo Is Nothing Then
        MsgBox "Faltan las hojas necesarias ('Venta por hora' o 'Combos pollo') para continuar.", vbCritical
        Exit Sub
    End If

    ' --- Preparar hojas de error y debug ---
    On Error Resume Next
    Set wsErrores = ThisWorkbook.Sheets("Combos no encontrados")
    If wsErrores Is Nothing Then
        Set wsErrores = ThisWorkbook.Sheets.Add
        wsErrores.Name = "Combos no encontrados"
    End If
    wsErrores.Cells.Clear
    wsErrores.Range("A1").Value = "Productos vendidos no encontrados en 'Combos pollo'"
    filaError = 2
    Set dicErrores = CreateObject("Scripting.Dictionary")
    
    Set wsDebug = ThisWorkbook.Sheets("Debug_Calculo_Piezas")
    If wsDebug Is Nothing Then
        Set wsDebug = ThisWorkbook.Sheets.Add
        wsDebug.Name = "Debug_Calculo_Piezas"
    End If
    wsDebug.Cells.Clear
    wsDebug.Range("A1:F1").Value = Array("Hora", "Combo Vendido", "Cant. Vendida", "Item de Pollo", "Cant. por Combo (Definición)", "Total Ítems Calculado")
    filaDebug = 2
    On Error GoTo 0
    
    Set ignoreList = CreateObject("Scripting.Dictionary")
    For Each itemToIgnore In Array("COCA SL", "PAPA CHICA", "COCA Z SL", "SPRITE Z SL", "FANTA Z SL", "PAPA MEDIANA", _
                                   "SALSA BARBACOA BUCKET", "SALSA PICANTE BUCKET", "SALSA MOSTAZA - MIEL BUCKET", _
                                   "AGUA 500 ML", "COCA BOTELLA GLV", "PAPA MEDIANA DOM", "PAPA CHICA DOM", _
                                   "ORIGINAL", "CRISPY", "SALSA BARBACOAP", "AGUA MINERAL", "COCA COLA GRANDE", _
                                   "CAFE C LECHE 12", "CHOCOLATE", "COCA CHICA")
        ignoreList(UCase$(CStr(itemToIgnore))) = True
    Next itemToIgnore

    ' --- PASO 2: CARGAR DICCIONARIO DE DEFINICIÓN DE COMBOS ---
    Set dicPollo = CreateObject("Scripting.Dictionary")
    ultFilaCombo = wsCombo.Cells(wsCombo.Rows.count, "A").End(xlUp).Row

    For i = 2 To ultFilaCombo
        comboNombre = Trim$(CStr(wsCombo.Cells(i, 2).Value))
        If comboNombre <> "" Then
            If Not dicPollo.Exists(comboNombre) Then
                Set dicPollo(comboNombre) = CreateObject("Scripting.Dictionary")
                For j = i To ultFilaCombo
                    nombreItemPollo = Trim$(CStr(wsCombo.Cells(j, 3).Value))
                    cantidadCombo = Val(wsCombo.Cells(j, 4).Value)
                    If nombreItemPollo <> "" And cantidadCombo > 0 Then
                        dicPollo(comboNombre)(nombreItemPollo) = Val(dicPollo(comboNombre)(nombreItemPollo)) + cantidadCombo
                    End If
                    If Trim$(CStr(wsCombo.Cells(j + 1, 2).Value)) <> "" Then Exit For
                Next j
            End If
        End If
    Next i

    ' --- PASO 3: PROCESAR DATOS ---
    Set dicVentaPollo = CreateObject("Scripting.Dictionary")
    ultFilaDatos = wsDatos.Cells(wsDatos.Rows.count, "A").End(xlUp).Row

    For i = 2 To ultFilaDatos
        hora = wsDatos.Cells(i, 1).Value
        producto = CStr(wsDatos.Cells(i, 2).Value)
        cantidadVenta = Val(wsDatos.Cells(i, 3).Value)

        If dicPollo.Exists(producto) Then
            For Each nombreItemPollo In dicPollo(producto).keys
                If Trim$(CStr(hora)) <> "" And Trim$(CStr(nombreItemPollo)) <> "" Then
                    clave = CStr(hora) & "|" & CStr(nombreItemPollo)
                    cantidadCombo = Val(dicPollo(producto)(nombreItemPollo))
                    dicVentaPollo(clave) = Val(dicVentaPollo(clave)) + (cantidadVenta * cantidadCombo)

                    ' DEBUG SOLO PARA PIEZAS CRISPY/ORIGINAL
                    itemUpper = UCase$(CStr(nombreItemPollo))
                    If itemUpper = "PIEZA CRISPY TERMINADA" Or itemUpper = "PIEZA ORIGINAL TERMINADA" Then
                        wsDebug.Cells(filaDebug, 1).Value = hora
                        wsDebug.Cells(filaDebug, 2).Value = producto
                        wsDebug.Cells(filaDebug, 3).Value = cantidadVenta
                        wsDebug.Cells(filaDebug, 4).Value = nombreItemPollo
                        wsDebug.Cells(filaDebug, 5).Value = cantidadCombo
                        wsDebug.Cells(filaDebug, 6).Value = cantidadVenta * cantidadCombo
                        filaDebug = filaDebug + 1
                    End If
                End If
            Next nombreItemPollo
        Else
            If Not ignoreList.Exists(UCase$(producto)) Then
                If Not dicErrores.Exists(producto) Then
                    dicErrores.Add producto, True
                    wsErrores.Cells(filaError, 1).Value = producto
                    filaError = filaError + 1
                End If
            End If
        End If
    Next i
    
    wsDebug.Columns.AutoFit
    
    If dicVentaPollo.count = 0 Then
        MsgBox "No se encontraron coincidencias de combos de pollo en las ventas del día.", vbExclamation
        wsErrores.Columns("A").AutoFit
        Exit Sub
    End If

    ' --- PASO 4 Y 5: ARMAR "Tabla de venta" ---
    On Error Resume Next
    Set wsVenta = ThisWorkbook.Sheets("Tabla de venta")
    If wsVenta Is Nothing Then
        Set wsVenta = ThisWorkbook.Sheets.Add
        wsVenta.Name = "Tabla de venta"
    End If
    wsVenta.Cells.Clear
    On Error GoTo 0

    itemsOrdenados = Array("Alitas", "Filete", "Ruster", "POP", "P. Crispy", "P. Original", "Strips")
    Set dictPeriodos = CreateObject("Scripting.Dictionary")

    wsVenta.Cells(1, 1).Value = "PRODUCTO / HORA"
    wsVenta.Cells(1, 2).Value = "% Ajuste"

    For Each clave In dicVentaPollo.keys
        If InStr(1, CStr(clave), "|", vbTextCompare) > 0 Then
            hora = Split(CStr(clave), "|")(0)
            If Not dictPeriodos.Exists(hora) Then dictPeriodos.Add hora, 0
        End If
    Next clave

    i = 3
    For Each hora In dictPeriodos.keys
        dictPeriodos(hora) = i
        wsVenta.Cells(1, i).Value = Trim$(Split(CStr(hora), "-")(0))
        i = i + 1
    Next hora

    colTotal = dictPeriodos.count + 3
    wsVenta.Cells(1, colTotal).Value = "Total"

    For i = LBound(itemsOrdenados) To UBound(itemsOrdenados)
        wsVenta.Cells(i + 2, 1).Value = itemsOrdenados(i)
        wsVenta.Cells(i + 2, 2).Value = 0
        wsVenta.Cells(i + 2, 2).NumberFormat = "0.00%"
    Next i

    For Each clave In dicVentaPollo.keys
        Dim partes() As String
        partes = Split(CStr(clave), "|")
        hora = partes(0)
        nombreItemPollo = partes(1)
        cantidadVenta = Val(dicVentaPollo(clave))

        Select Case UCase$(CStr(nombreItemPollo))
            Case "NUEVO POPCORN MEDIANO TERMINADO"
                producto = "POP": peso = Round(cantidadVenta * 0.18)
            Case "NUEVO POPCORN GRANDE TERMINADO", "POPCORN GRANDE TERMINADO (Sin envase)"
                producto = "POP": peso = Round(cantidadVenta * 0.2)
            Case "RUSTER TERMINADO"
                producto = "Ruster": peso = Round(cantidadVenta)
            Case "ALITA TERMINADO"
                producto = "Alitas": peso = Round(cantidadVenta)
            Case "STRIP TERMINADO"
                producto = "Strips": peso = Round(cantidadVenta)
            Case "PIEZA CRISPY TERMINADA"
                producto = "P. Crispy": peso = Round(cantidadVenta)
            Case "PIEZA ORIGINAL TERMINADA"
                producto = "P. Original": peso = Round(cantidadVenta)
            Case Else
                If Left$(UCase$(CStr(nombreItemPollo)), 3) = "KCS" Then
                    producto = "Filete": peso = Round(cantidadVenta)
                Else
                    producto = ""
                End If
        End Select

        If producto <> "" And dictPeriodos.Exists(hora) Then
            For filaOut = 2 To UBound(itemsOrdenados) + 2
                If wsVenta.Cells(filaOut, 1).Value = producto Then
                    colOut = dictPeriodos(hora)
                    wsVenta.Cells(filaOut, colOut).Value = Val(wsVenta.Cells(filaOut, colOut).Value) + peso
                    Exit For
                End If
            Next filaOut
        End If
    Next clave

    For fila = 2 To UBound(itemsOrdenados) + 2
        If colTotal > 3 Then
            wsVenta.Cells(fila, colTotal).Formula = "=SUM(" & wsVenta.Range(wsVenta.Cells(fila, 3), wsVenta.Cells(fila, colTotal - 1)).Address & ")"
        Else
            wsVenta.Cells(fila, colTotal).Value = 0
        End If
    Next fila

    Dim RangoTabla As Range
    Set RangoTabla = wsVenta.Range("A1", wsVenta.Cells(UBound(itemsOrdenados) + 2, colTotal))

    With RangoTabla
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With wsVenta.Range(wsVenta.Cells(1, 1), wsVenta.Cells(1, colTotal))
        .Interior.Color = RGB(207, 18, 39)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .RowHeight = 30
    End With

    With wsVenta.Range(wsVenta.Cells(2, 1), wsVenta.Cells(UBound(itemsOrdenados) + 2, 1))
        .Interior.Color = RGB(230, 230, 230)
        .Font.Bold = True
    End With

    With wsVenta.Range(wsVenta.Cells(2, colTotal), wsVenta.Cells(UBound(itemsOrdenados) + 2, colTotal))
        .Interior.Color = RGB(245, 245, 245)
        .Font.Bold = True
    End With

    wsVenta.Columns.AutoFit
    wsErrores.Columns("A").AutoFit

End Sub


