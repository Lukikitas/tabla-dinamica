Attribute VB_Name = "Módulo8"
Sub VerificarCoincidenciasDeCombos()

    Dim wsVenta As Worksheet, wsCombos As Worksheet, wsErrores As Worksheet
    Dim UltFilaVenta As Long, UltFilaCombos As Long
    Dim i As Long, j As Long
    Dim ComboVenta As String, ComboCombo As String
    Dim Encontrado As Boolean
    Dim filaError As Long

    Set wsVenta = Sheets("Resumen de 4 semanas")
    Set wsCombos = Sheets("Combos pollo")

    ' Crear hoja para combos no encontrados
    On Error Resume Next
    Set wsErrores = Sheets("Combos no encontrados")
    If wsErrores Is Nothing Then
        Set wsErrores = Sheets.Add(After:=Sheets(Sheets.count))
        wsErrores.Name = "Combos no encontrados"
    Else
        wsErrores.Cells.Clear
    End If
    On Error GoTo 0

    wsErrores.Cells(1, 1).Value = "Combo en venta por hora"
    wsErrores.Cells(1, 2).Value = "Encontrado en Combos pollo?"

    UltFilaVenta = wsVenta.Cells(wsVenta.Rows.count, 2).End(xlUp).Row
    UltFilaCombos = wsCombos.Cells(wsCombos.Rows.count, 2).End(xlUp).Row

    filaError = 2

    ' Comparar cada combo de "venta por hora" contra los combos en "Combos pollo"
    For i = 2 To UltFilaVenta
        ComboVenta = NormalizarTexto(wsVenta.Cells(i, 2).Value)
        Encontrado = False

        For j = 2 To UltFilaCombos
            ComboCombo = NormalizarTexto(wsCombos.Cells(j, 2).Value)
            If ComboVenta = ComboCombo Then
                Encontrado = True
                Exit For
            End If
        Next j

        If Not Encontrado Then
            wsErrores.Cells(filaError, 1).Value = wsVenta.Cells(i, 2).Value
            wsErrores.Cells(filaError, 2).Value = "NO"
            filaError = filaError + 1
        End If
    Next i

    ' Eliminar duplicados en la hoja de errores
    With wsErrores
        .Range("A1:B" & filaError - 1).RemoveDuplicates Columns:=1, Header:=xlYes
    End With


End Sub

Function NormalizarTexto(txt As String) As String
    NormalizarTexto = UCase(Trim(Replace(Replace(txt, Chr(160), ""), vbTab, "")))
End Function


