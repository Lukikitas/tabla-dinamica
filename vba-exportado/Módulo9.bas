Attribute VB_Name = "Módulo9"
Sub ImportarPromedioDeVentasPorHora()

    Dim fechaInput As String, fechaBase As Date
    Dim fechas(1 To 4) As String
    Dim i As Long, f As Long
    Dim url As String, http As Object
    Dim respuesta As String
    Dim regex As Object, matches As Object
    Dim bloque As String, hora As String, producto As String, cantidad As Long
    Dim regexProd As Object, matchesProd As Object
    Dim subMatch As Object
    Dim clave As Variant
    Dim acumulador As Object
    Dim debugWS As Worksheet, debugFila As Long

    ' Pedir fecha al usuario
    fechaInput = InputBox("Ingrese la fecha (dd/mm/yyyy):", "Fecha base")

    If Not IsDate(fechaInput) Then
        MsgBox "Formato de fecha incorrecto.", vbExclamation
        Exit Sub
    End If

    fechaBase = CDate(fechaInput)

    ' Calcular las 4 fechas anteriores (mismo día, sin incluir la ingresada)
    For i = 1 To 4
        fechas(i) = Format(fechaBase - 7 * i, "yyyy/mm/dd")
    Next i

    ' Crear acumulador
    Set acumulador = CreateObject("Scripting.Dictionary")

    ' Preparar hoja de debug
    On Error Resume Next
    Set debugWS = Sheets("Debug ventas semanales")
    If debugWS Is Nothing Then
        Set debugWS = Sheets.Add
        debugWS.Name = "Debug ventas semanales"
    Else
        debugWS.Cells.Clear
    End If
    On Error GoTo 0

    debugWS.Range("A1:D1").Value = Array("Fecha", "Hora", "Producto", "Cantidad")
    debugFila = 2

    ' Preparar expresiones regulares
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = """hora"":\s*""([^""]+)""[\s\S]+?""cuerpo"":\s*\[(.*?)\](?=},|\}\])"
    End With

    Set regexProd = CreateObject("VBScript.RegExp")
    With regexProd
        .Global = True
        .IgnoreCase = True
        .Pattern = """producto"":\s*""([^""]+)"".*?""cantidad"":\s*([0-9]+)"
    End With

    ' Recorrer las 4 fechas
    For f = 1 To 4
        url = "http://10.211.10.250:3000/reporte/Ventas/ProductoHora?restaurante=19&fecha_inicio=" & fechas(f) & "&fecha_fin=" & fechas(f) & "&estado=&canal=undefined&usuario=2C91C9BA-B49A-EE11-8925-6045BDBB68D1&cadena=1&cajero=undefined"

        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", url, False
        http.Send

        If http.Status <> 200 Then
            MsgBox "Error al obtener datos del " & fechas(f), vbExclamation
            Exit Sub
        End If

        respuesta = http.responseText
        Set matches = regex.Execute(respuesta)

        For Each match In matches
            hora = match.SubMatches(0)
            bloque = match.SubMatches(1)

            Set matchesProd = regexProd.Execute(bloque)
            For Each subMatch In matchesProd
                producto = Replace(subMatch.SubMatches(0), "*", "")
                cantidad = CLng(subMatch.SubMatches(1))
                clave = hora & "|" & producto

                ' Guardar en hoja de debug
                debugWS.Cells(debugFila, 1).Value = fechas(f)
                debugWS.Cells(debugFila, 2).Value = hora
                debugWS.Cells(debugFila, 3).Value = producto
                debugWS.Cells(debugFila, 4).Value = cantidad
                debugFila = debugFila + 1

                ' Acumular para promedio
                If acumulador.Exists(clave) Then
                    acumulador(clave) = acumulador(clave) + cantidad
                Else
                    acumulador.Add clave, cantidad
                End If
            Next subMatch
        Next match
    Next f

    ' Crear hoja de salida
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets("Venta por hora promedio")
    If ws Is Nothing Then
        Set ws = Sheets.Add
        ws.Name = "Venta por hora promedio"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ws.Range("A1:C1").Value = Array("Hora", "Producto", "Promedio")
    i = 2
    For Each clave In acumulador.keys
        hora = Split(clave, "|")(0)
        producto = Split(clave, "|")(1)
        cantidad = acumulador(clave)

        ws.Cells(i, 1).Value = hora
        ws.Cells(i, 2).Value = producto
        ws.Cells(i, 3).Value = Round(cantidad / 4, 2)
        i = i + 1
    Next clave

    ws.Columns("A:C").AutoFit
    debugWS.Columns("A:D").AutoFit

    MsgBox "Promedio generado y debug guardado en hoja 'Debug ventas semanales'.", vbInformation

End Sub

