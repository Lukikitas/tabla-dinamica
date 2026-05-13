Attribute VB_Name = "Módulo4"
Sub GenerarVentaDePolloPorHora()

    Dim wsVenta As Worksheet, wsCombos As Worksheet, wsSalida As Worksheet
    Dim Diccionario As Object, clave As String
    Dim UltFilaVenta As Long, ultFilaCombo As Long
    Dim fila As Long, i As Long
    Dim periodo As String, combo As String, CantCombo As Long
    Dim filaCombo As Long
    Dim item As String, CantidadItem As Double
    Dim claves() As Variant
    Dim ultimoCombo As String
    Dim comboFilaCombo As String
    Dim ComboYaProcesado As Boolean

    Set wsVenta = Sheets("Resumen de 4 semanas")
    Set wsCombos = Sheets("Combos pollo")

    ' Crear hoja de salida o limpiar si ya existe
    On Error Resume Next
    Set wsSalida = Sheets("Venta de pollo")
    If wsSalida Is Nothing Then
        Set wsSalida = Sheets.Add(After:=Sheets(Sheets.count))
        wsSalida.Name = "Venta de pollo"
    Else
        wsSalida.Cells.Clear
    End If
    On Error GoTo 0

    wsSalida.Range("A1:C1").Value = Array("Hora", "Nombre de ítem", "Cantidad")
    Set Diccionario = CreateObject("Scripting.Dictionary")

    UltFilaVenta = wsVenta.Cells(wsVenta.Rows.count, 1).End(xlUp).Row
    ultFilaCombo = wsCombos.Cells(wsCombos.Rows.count, 1).End(xlUp).Row

    For fila = 2 To UltFilaVenta
        periodo = Trim(wsVenta.Cells(fila, 1).Value)
        combo = NormalizarTexto(wsVenta.Cells(fila, 2).Value)
        CantCombo = wsVenta.Cells(fila, 3).Value

        If periodo <> "" And combo <> "" And CantCombo > 0 Then
            ComboYaProcesado = False
            ultimoCombo = ""

            For filaCombo = 2 To ultFilaCombo
                comboFilaCombo = Trim(wsCombos.Cells(filaCombo, 2).Value)

                ' Si encontramos un nuevo combo, lo registramos
                If comboFilaCombo <> "" Then
                    ultimoCombo = NormalizarTexto(comboFilaCombo)

                    ' Si ya procesamos este combo antes, salteamos el resto del archivo
                    If ultimoCombo = combo Then
                        If ComboYaProcesado Then Exit For
                        ComboYaProcesado = True
                    ElseIf ComboYaProcesado Then
                        Exit For ' Terminamos el bloque del combo procesado
                    End If
                End If

                ' Procesar solo si estamos en el bloque correcto
                If ultimoCombo = combo And ComboYaProcesado Then
                    item = Trim(wsCombos.Cells(filaCombo, 3).Value)
                    CantidadItem = wsCombos.Cells(filaCombo, 4).Value

                    If item <> "" And CantidadItem > 0 Then
                        clave = periodo & "|" & item
                        If Diccionario.Exists(clave) Then
                            Diccionario(clave) = Diccionario(clave) + CantCombo * CantidadItem
                        Else
                            Diccionario.Add clave, CantCombo * CantidadItem
                        End If
                    End If
                End If
            Next filaCombo
        End If
    Next fila

    If Diccionario.count = 0 Then
        MsgBox "No se encontraron coincidencias para generar la venta de pollo.", vbExclamation
        Exit Sub
    End If

    claves = Diccionario.keys
    i = 2
    For fila = 0 To UBound(claves)
        clave = claves(fila)
        wsSalida.Cells(i, 1).Value = Split(clave, "|")(0)
        wsSalida.Cells(i, 2).Value = Split(clave, "|")(1)
        wsSalida.Cells(i, 3).Value = Diccionario(clave)
        i = i + 1
    Next fila

End Sub

Function NormalizarTexto(txt As String) As String
    NormalizarTexto = UCase(Trim(Replace(Replace(txt, Chr(160), ""), vbTab, "")))
End Function

