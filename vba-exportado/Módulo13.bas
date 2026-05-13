Attribute VB_Name = "Módulo13"
Sub ResumirVentas4Semanas()

    Dim resumen As Object
    Dim wsResumen As Worksheet
    Dim hoja As Worksheet
    Dim clave As Variant
    Dim hora As String, producto As String
    Dim cantidad As Long
    Dim fila As Long
    Dim i As Long

    ' Crear diccionario para acumular datos
    Set resumen = CreateObject("Scripting.Dictionary")

    ' Recorrer hojas Semana 1 a Semana 4
    For i = 1 To 4
        Set hoja = ThisWorkbook.Sheets("Semana " & i)

        For fila = 2 To hoja.Cells(hoja.Rows.count, "A").End(xlUp).Row
            hora = hoja.Cells(fila, 1).Value
            producto = hoja.Cells(fila, 2).Value
            cantidad = hoja.Cells(fila, 3).Value

            clave = hora & "|" & producto

            If resumen.Exists(clave) Then
                resumen(clave) = resumen(clave) + cantidad
            Else
                resumen.Add clave, cantidad
            End If
        Next fila
    Next i

    ' Crear hoja de resumen
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Resumen de 4 semanas").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsResumen = Sheets.Add
    wsResumen.Name = "Resumen de 4 semanas"

    wsResumen.Range("A1:C1").Value = Array("Hora", "Producto", "Cantidad")

    ' Volcar los datos
    fila = 2
    For Each clave In resumen.keys
        hora = Split(clave, "|")(0)
        producto = Split(clave, "|")(1)
        cantidad = resumen(clave)

        wsResumen.Cells(fila, 1).Value = hora
        wsResumen.Cells(fila, 2).Value = producto
        wsResumen.Cells(fila, 3).Value = cantidad
        fila = fila + 1
    Next clave

    ' Ordenar por Hora (columna A)
    With wsResumen.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsResumen.Range("A2:A" & fila - 1), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange wsResumen.Range("A1:C" & fila - 1)
        .Header = xlYes
        .Apply
    End With

    wsResumen.Columns("A:C").AutoFit


End Sub


