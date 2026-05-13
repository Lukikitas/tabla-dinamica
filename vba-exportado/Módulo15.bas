Attribute VB_Name = "Módulo15"
Sub VerificarContenidoDeDiccionarioPollo()

    Dim wsCombo As Worksheet, wsSalida As Worksheet
    Dim combo As Variant, nombreItemPollo As Variant
    Dim cantidadCombo As Long
    Dim filaCombo As Long, ultFilaCombo As Long, fila As Long
    Dim dicPollo As Object

    ' Obtener hoja Combos pollo
    Set wsCombo = Sheets("Combos pollo")
    ultFilaCombo = wsCombo.Cells(wsCombo.Rows.count, 1).End(xlUp).Row

    ' Crear diccionario
    Set dicPollo = CreateObject("Scripting.Dictionary")

    For filaCombo = 2 To ultFilaCombo
        combo = Trim(wsCombo.Cells(filaCombo, 2).Value)
        nombreItemPollo = Trim(wsCombo.Cells(filaCombo, 3).Value)
        cantidadCombo = wsCombo.Cells(filaCombo, 4).Value

        If combo <> "" And nombreItemPollo <> "" And cantidadCombo > 0 Then
            If Not dicPollo.Exists(combo) Then
                Set dicPollo(combo) = CreateObject("Scripting.Dictionary")
            End If

            ' ?? CORRECTO: Solo agregamos si NO existe
            If dicPollo(combo).Exists(nombreItemPollo) Then
                dicPollo(combo)(nombreItemPollo) = dicPollo(combo)(nombreItemPollo) + cantidadCombo
            Else
                dicPollo(combo).Add nombreItemPollo, cantidadCombo
            End If
        End If
    Next filaCombo

    ' Crear hoja de resultado
    On Error Resume Next
    Set wsSalida = Sheets("Diccionario Pollo")
    If wsSalida Is Nothing Then
        Set wsSalida = Sheets.Add(After:=Sheets(Sheets.count))
        wsSalida.Name = "Diccionario Pollo"
    Else
        wsSalida.Cells.Clear
    End If
    On Error GoTo 0

    ' Escribir encabezados
    wsSalida.Range("A1:C1").Value = Array("Combo", "Ítem de Pollo", "Cantidad")
    fila = 2

    ' Volcar contenido
    For Each combo In dicPollo.keys
        For Each nombreItemPollo In dicPollo(combo).keys
            wsSalida.Cells(fila, 1).Value = combo
            wsSalida.Cells(fila, 2).Value = nombreItemPollo
            wsSalida.Cells(fila, 3).Value = dicPollo(combo)(nombreItemPollo)
            fila = fila + 1
        Next nombreItemPollo
    Next combo

    wsSalida.Columns("A:C").AutoFit
    MsgBox "Verificación completa. Revisa la hoja 'Diccionario Pollo'.", vbInformation

End Sub

