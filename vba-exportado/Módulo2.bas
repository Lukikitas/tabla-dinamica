Attribute VB_Name = "Módulo2"
Sub ExtraerCombosPollo()

    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim ultimaFila As Long, i As Long, filaDestino As Long
    Dim plu As String, nombreCombo As String
    Dim leyendoItems As Boolean
    Dim primeraFilaCombo As Boolean
    Dim itemNombre As String
    Dim tieneItemPollo As Boolean
    
    Dim listaPollo As Variant
    listaPollo = Array( _
        "STRIP TERMINADO ", _
        "PIEZA CRISPY TERMINADA ", _
        "PIEZA ORIGINAL TERMINADA ", _
        "ALITA TERMINADO ", _
        "RUSTER TERMINADO ", _
        "RUSTER QUESO EN FETA TERMINADO ", _
        "KCS DELUXE TERMINADO ", _
        "KCS ORIGINAL TERMINADO ", _
        "KCS BBQ BACON TERMINADO ", _
        "KCS MELT TERMINADO ", _
        "POPCORN GRANDE TERMINADO (Sin envase) ", _
        "NUEVO POPCORN MEDIANO TERMINADO ", _
        "NUEVO POPCORN GRANDE TERMINADO " _
    )
    
    Set wsOrigen = ThisWorkbook.Sheets(1) ' Ajustar si es necesario
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Combos pollo")
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Sheets.Add(After:=wsOrigen)
        wsDestino.Name = "Combos pollo"
    Else
        wsDestino.Cells.ClearContents
    End If
    On Error GoTo 0

    wsDestino.Range("A1:D1").Value = Array("PLU", "Nombre del PLU", "Nombre del Item", "Cantidad")

    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.count, "A").End(xlUp).Row
    filaDestino = 2
    leyendoItems = False

    For i = 1 To ultimaFila
        If wsOrigen.Cells(i, 1).Value = "# PLU" Then
            plu = wsOrigen.Cells(i + 1, 1).Value
            nombreCombo = wsOrigen.Cells(i + 1, 2).Value
            leyendoItems = False
        ElseIf wsOrigen.Cells(i, 1).Value = "# ITEM" Then
            leyendoItems = True
            primeraFilaCombo = True
            i = i + 1
            
            Dim filaInicio As Long
            filaInicio = i
            Dim filaFin As Long
            ' Buscar fin de ítems
            Do While wsOrigen.Cells(i, 1).Value <> "" Or wsOrigen.Cells(i, 6).Value <> ""
                i = i + 1
            Loop
            filaFin = i - 1
            
            ' Verificar si contiene al menos un ítem de pollo
            tieneItemPollo = False
            For j = filaInicio To filaFin
                itemNombre = wsOrigen.Cells(j, 6).Value
                If EsItemPollo(itemNombre, listaPollo) Then
                    tieneItemPollo = True
                    Exit For
                End If
            Next j
            
            ' Si contiene, copiar solo los ítems de pollo
            If tieneItemPollo Then
                For j = filaInicio To filaFin
                    itemNombre = wsOrigen.Cells(j, 6).Value
                    If EsItemPollo(itemNombre, listaPollo) Then
                        If primeraFilaCombo Then
                            wsDestino.Cells(filaDestino, 1).Value = plu
                            wsDestino.Cells(filaDestino, 2).Value = nombreCombo
                            primeraFilaCombo = False
                        End If
                        wsDestino.Cells(filaDestino, 3).Value = itemNombre
                        wsDestino.Cells(filaDestino, 4).Value = wsOrigen.Cells(j, 8).Value ' Cantidad
                        filaDestino = filaDestino + 1
                    End If
                Next j
            End If
            
            i = filaFin
        End If
    Next i

    MsgBox "¡Extracción de combos con pollo completada!", vbInformation

End Sub

Function EsItemPollo(nombre As String, lista As Variant) As Boolean
    Dim k As Long
    For k = LBound(lista) To UBound(lista)
        If Trim(nombre) = Trim(lista(k)) Then
            EsItemPollo = True
            Exit Function
        End If
    Next k
    EsItemPollo = False
End Function



