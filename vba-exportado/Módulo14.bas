Attribute VB_Name = "Módulo14"
Sub GenerarPromedioPolloPorHora()

    Dim wsCombo As Worksheet, wsSemana As Worksheet, wsSalida As Worksheet
    Dim i As Long, f As Long, fila As Long
    Dim combo As String, nombreItemPollo As Variant, cantidad As Long
    Dim hora As String, clave As String
    Dim comboNombre As String
    Dim cantidadVenta As Long, cantidadCombo As Long
    Dim dicPollo As Object, dicSemana As Object
    Dim claves As Variant
    Dim filaCombo As Long, ultFilaCombo As Long

    ' Obtener hoja "Combos pollo"
    On Error Resume Next
    Set wsCombo = Sheets("Combos pollo")
    If wsCombo Is Nothing Then
        MsgBox "No existe la hoja 'Combos pollo'", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Crear diccionario de combos de pollo ? {combo -> {item -> cantidad}}
    Set dicPollo = CreateObject("Scripting.Dictionary")

    ultFilaCombo = wsCombo.Cells(wsCombo.Rows.count, 1).End(xlUp).Row

    For filaCombo = 2 To ultFilaCombo
        combo = Trim(wsCombo.Cells(filaCombo, 2).Value)
        nombreItemPollo = Trim(wsCombo.Cells(filaCombo, 3).Value)
        cantidad = wsCombo.Cells(filaCombo, 4).Value

        If combo <> "" And nombreItemPollo <> "" And cantidad > 0 Then
            If Not dicPollo.Exists(combo) Then
                Set dicPollo(combo) = CreateObject("Scripting.Dictionary")
            End If

            If dicPollo(combo).Exists(nombreItemPollo) Then
                dicPollo(combo)(nombreItemPollo) = dicPollo(combo)(nombreItemPollo) + cantidad
            Else
                dicPollo(combo).Add nombreItemPollo, cantidad
            End If
        End If
    Next filaCombo

    ' Crear diccionario general ? clave = hora|item, valor = suma
    Set dicSemana = CreateObject("Scripting.Dictionary")

    ' Leer Semana 1 a 4
    For f = 1 To 4
        Set wsSemana = Nothing
        On Error Resume Next
        Set wsSemana = Sheets("Semana " & f)
        On Error GoTo 0

        If wsSemana Is Nothing Then
            MsgBox "No se encontró la hoja 'Semana " & f & "'", vbExclamation
            Exit Sub
        End If

        fila = 2
        Do While wsSemana.Cells(fila, 1).Value <> ""
            hora = Trim(wsSemana.Cells(fila, 1).Value)
            comboNombre = Trim(wsSemana.Cells(fila, 2).Value)
            cantidadVenta = wsSemana.Cells(fila, 3).Value

            If dicPollo.Exists(comboNombre) Then
                For Each nombreItemPollo In dicPollo(comboNombre).keys
                    cantidadCombo = dicPollo(comboNombre)(nombreItemPollo)
                    clave = hora & "|" & nombreItemPollo

                    If dicSemana.Exists(clave) Then
                        dicSemana(clave) = dicSemana(clave) + (cantidadVenta * cantidadCombo)
                    Else
                        dicSemana.Add clave, (cantidadVenta * cantidadCombo)
                    End If
                Next nombreItemPollo
            End If

            fila = fila + 1
        Loop
    Next f

    ' Crear hoja de salida
    On Error Resume Next
    Set wsSalida = Sheets("Promedio pollo por hora")
    If wsSalida Is Nothing Then
        Set wsSalida = Sheets.Add(After:=Sheets(Sheets.count))
        wsSalida.Name = "Promedio pollo por hora"
    Else
        wsSalida.Cells.Clear
    End If
    On Error GoTo 0

    ' Escribir resultados
    wsSalida.Range("A1:C1").Value = Array("Hora", "Ítem de Pollo", "Promedio")
    fila = 2
    claves = dicSemana.keys

    For i = 0 To UBound(claves)
        clave = claves(i)
        hora = Split(clave, "|")(0)
        nombreItemPollo = Split(clave, "|")(1)
        cantidad = dicSemana(clave)

        wsSalida.Cells(fila, 1).Value = hora
        wsSalida.Cells(fila, 2).Value = nombreItemPollo
        wsSalida.Cells(fila, 3).Value = Round(cantidad / 4, 2)
        fila = fila + 1
    Next i

    wsSalida.Columns("A:C").AutoFit
    MsgBox "Promedio de pollo por hora generado correctamente.", vbInformation

End Sub

