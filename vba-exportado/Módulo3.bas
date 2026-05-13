Attribute VB_Name = "Módulo3"
Sub ImportarVentasPorHora()

    Dim RutaArchivo As String
    Dim texto As String
    Dim linea As String
    Dim horaActual As String
    Dim producto As String
    Dim cantidad As String
    Dim i As Long
    Dim fila As Long
    Dim regex As Object
    Dim matches As Object
    
    ' Seleccionar archivo de texto
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Selecciona el archivo .txt"
        .Filters.Add "Archivos de texto", "*.txt", 1
        If .Show <> -1 Then Exit Sub
        RutaArchivo = .SelectedItems(1)
    End With
    
    ' Leer contenido completo
    Open RutaArchivo For Input As #1
    texto = Input$(LOF(1), 1)
    Close #1
    
    ' Limpiar hoja "venta por hora"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets("venta por hora")
    If ws Is Nothing Then
        Set ws = Sheets.Add
        ws.Name = "venta por hora"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ws.Cells(1, 1).Value = "Hora"
    ws.Cells(1, 2).Value = "Producto"
    ws.Cells(1, 3).Value = "Cantidad"
    
    ' Preparar expresión regular
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = """[^""]+"":\s*""([^""]+)"""
    
    ' Separar líneas
    Dim lineas() As String
    lineas = Split(texto, vbLf)
    fila = 2
    
    For i = 0 To UBound(lineas)
        linea = Trim(lineas(i))
        
        ' Buscar la hora
        If InStr(linea, """hora"":") > 0 Then
            Set matches = regex.Execute(linea)
            If matches.count > 0 Then horaActual = matches(0).SubMatches(0)
        
        ' Buscar el producto
        ElseIf InStr(linea, """producto"":") > 0 Then
            Set matches = regex.Execute(linea)
            If matches.count > 0 Then producto = matches(0).SubMatches(0)
        
        ' Buscar la cantidad
        ElseIf InStr(linea, """cantidad"":") > 0 Then
            cantidad = ObtenerCantidad(linea)
            
            ' Solo escribir si tenemos todos los datos
            If horaActual <> "" And producto <> "" And cantidad <> "" Then
                ws.Cells(fila, 1).Value = horaActual
                ws.Cells(fila, 2).Value = Replace(producto, "*", "") ' sin asteriscos
                ws.Cells(fila, 3).Value = cantidad
                fila = fila + 1
                producto = ""
                cantidad = ""
            End If
        End If
    Next i

    MsgBox "Importación completa. Total filas: " & fila - 2

End Sub

Function ObtenerCantidad(linea As String) As String
    Dim partes() As String
    partes = Split(linea, ":")
    If UBound(partes) >= 1 Then
        ObtenerCantidad = Trim(Replace(partes(1), ",", ""))
    Else
        ObtenerCantidad = ""
    End If
End Function

