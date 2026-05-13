Attribute VB_Name = "Módulo5"
Sub GenerarPromedioPolloPorHoraConDebug()

    Dim wsCombo As Worksheet, wsSemana As Worksheet, wsSalida As Worksheet, wsDetalle As Worksheet
    Dim i As Long, f As Long, fila As Long
    Dim combo As String, nombreItemPollo As Variant
    Dim cantidadCombo As Long, cantidadVenta As Long
    Dim hora As String, clave As String, comboNombre As String
    Dim filaCombo As Long, ultFilaCombo As Long
    Dim dicPollo As Object, dicSemana As Object
    Dim claves As Variant
    Dim filaDetalle As Long

    ' Obtener hoja "Combos pollo"
    On Error Resume Next
    Set wsCombo = Sheets("Combos pollo")
    If wsCombo Is Nothing Then
        MsgBox "No existe la hoja 'Combos pollo'", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Crear diccionario de combos ? {Combo -> {Ítem -> Cantidad}}
    Set dicPollo = CreateObject("Scripting.Dictionary")
    ultFilaCombo = wsCombo.Cells(wsCombo.Rows.count, 1).End(xlUp).Row

    For filaCombo = 2 To ultFilaCombo
        combo = Trim(wsCombo.Cells(filaCombo, 2).Value)
        nombreItemPollo = Trim(wsCombo.Cells(filaCombo, 3).Value)
        cantidadCombo = wsCombo.Cells(filaCombo, 4).Value

        If combo <> "" And nombreItemPollo <> "" And cantidadCombo > 0 Then
            If Not dicPollo.Exists(combo) Then
                Set dicPollo(combo) = CreateObject("Scripting.Dictionary")
            End If
            ' No sobrescribir si ya existe
            If dicPollo(combo).Exists(nombreItemPollo) Then
                dicPollo(combo)(nombreItemPollo) = dicPollo(combo)(nombreItemPollo) + cantidadCombo
            Else
                dicPollo(combo).Add nombreItemPollo, cantidadCombo
            End If
        End If
    Next filaCombo

    ' Crear diccionario para acumulación por hora|item
    Set dicSemana = CreateObject("Scripting.Dictionary")

    ' Preparar hoja de detalle
    On Error Resume Next
    Set wsDetalle = Sheets("Detalle de cálculo pollo")
    If wsDetalle Is Nothing Then
        Set wsDetalle = Sheets.Add(After:=Sheets(Sheets.count))
        wsDetalle.Name = "Detalle de cálculo pollo"
    Else
        wsDetalle.Cells.Clear
    End If
    On Error GoTo 0

    wsDetalle.Range("A1:G1").Value = Array("Semana", "Hora", "Combo", "Ítem de Pollo", "Cant Combos", "Cant por Combo", "Total")
    filaDetalle = 2

    ' Procesar Semana 1 a Semana 4
    For f = 1 To 4
        Set wsSemana = Nothing
        On Error Resume Next
        Set wsSemana = Sheets("Semana " & f)
        On Error GoTo 0

        If wsSemana Is Nothing Then
            MsgBox "Falta la hoja Semana " & f, vbExclamation
            Exit Sub
        End If

        fila = 2
        Do While wsSemana.Cells(fila, 1).Value <> ""
            hora = Trim(wsSemana.Cells(fila, 1).Value)
            comboNombre = Trim(wsSemana.Cells(fila, 2).Value)
            cantidadVenta = Val(wsSemana.Cells(fila, 3).Value)

            If dicPollo.Exists(comboNombre) Then
                For Each nombreItemPollo In dicPollo(comboNombre).keys
                    cantidadCombo = dicPollo(comboNombre)(nombreItemPollo)
                    clave = hora & "|" & nombreItemPollo

                    ' Acumular total por ítem
                    If dicSemana.Exists(clave) Then
                        dicSemana(clave) = dicSemana(clave) + (cantidadVenta * cantidadCombo)
                    Else
                        dicSemana.Add clave, (cantidadVenta * cantidadCombo)
                    End If

                    ' Escribir en hoja de detalle
                    wsDetalle.Cells(filaDetalle, 1).Value = f
                    wsDetalle.Cells(filaDetalle, 2).Value = hora
                    wsDetalle.Cells(filaDetalle, 3).Value = comboNombre
                    wsDetalle.Cells(filaDetalle, 4).Value = nombreItemPollo
                    wsDetalle.Cells(filaDetalle, 5).Value = cantidadVenta
                    wsDetalle.Cells(filaDetalle, 6).Value = cantidadCombo
                    wsDetalle.Cells(filaDetalle, 7).Value = cantidadVenta * cantidadCombo
                    filaDetalle = filaDetalle + 1
                Next nombreItemPollo
            End If

            fila = fila + 1
        Loop
    Next f

    ' Crear hoja de promedio
    On Error Resume Next
    Set wsSalida = Sheets("Promedio pollo por hora")
    If wsSalida Is Nothing Then
        Set wsSalida = Sheets.Add(After:=Sheets(Sheets.count))
        wsSalida.Name = "Promedio pollo por hora"
    Else
        wsSalida.Cells.Clear
    End If
    On Error GoTo 0

    wsSalida.Range("A1:C1").Value = Array("Hora", "Ítem de Pollo", "Promedio")
    fila = 2
    claves = dicSemana.keys

    For i = 0 To UBound(claves)
        clave = claves(i)
        hora = Split(clave, "|")(0)
        nombreItemPollo = Split(clave, "|")(1)
        cantidadCombo = dicSemana(clave)

        wsSalida.Cells(fila, 1).Value = hora
        wsSalida.Cells(fila, 2).Value = nombreItemPollo
        wsSalida.Cells(fila, 3).Value = Round(cantidadCombo / 4, 2)
        fila = fila + 1
    Next i

    wsSalida.Columns("A:C").AutoFit
    wsDetalle.Columns("A:G").AutoFit

    MsgBox "Promedio generado correctamente con detalle de cálculo.", vbInformation

End Sub

