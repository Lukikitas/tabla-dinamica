Attribute VB_Name = "Módulo1"
Sub ExtraerCombos()

    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim ultimaFila As Long, i As Long, filaDestino As Long
    Dim plu As String, nombreCombo As String
    Dim leyendoItems As Boolean
    Dim primeraFilaCombo As Boolean
    
    Set wsOrigen = ThisWorkbook.Sheets(1) ' Ajustar si no es la primera hoja
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Combos")
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Sheets.Add(After:=wsOrigen)
        wsDestino.Name = "Combos"
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
            i = i + 1 ' Salta a la primera fila de ítems
            Do While wsOrigen.Cells(i, 1).Value <> "" Or wsOrigen.Cells(i, 6).Value <> ""
                If Trim(wsOrigen.Cells(i, 6).Value) <> "" Then
                    ' Solo escribe PLU y nombre del combo en la primera fila del grupo
                    If primeraFilaCombo Then
                        wsDestino.Cells(filaDestino, 1).Value = plu
                        wsDestino.Cells(filaDestino, 2).Value = nombreCombo
                        primeraFilaCombo = False
                    End If
                    wsDestino.Cells(filaDestino, 3).Value = wsOrigen.Cells(i, 6).Value ' Item
                    wsDestino.Cells(filaDestino, 4).Value = wsOrigen.Cells(i, 8).Value ' Cantidad
                    filaDestino = filaDestino + 1
                End If
                i = i + 1
            Loop
            i = i - 1 ' Para que el bucle principal continúe correctamente
        End If
    Next i

    MsgBox "¡Extracción de combos completada!", vbInformation

End Sub


