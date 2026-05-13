Attribute VB_Name = "Módulo10"
Option Explicit

' Módulo10 (Versión Final con Consolidación por Hora+Producto)
' - Importa el JSON desde el endpoint
' - Extrae (Hora, Producto, Cantidad)
' - CONSOLIDA duplicados en la misma hora sumando cantidades

Sub ImportarVentasPorHoraFormateada()

    Dim url As String, http As Object, respuesta As String
    Dim ws As Worksheet
    Dim matches As Object, regex As Object
    Dim fila As Long
    Dim bloque As String, horaActual As String
    Dim producto As String, productoKey As String
    Dim cantidad As Long
    Dim fechaSeleccionada As String, fechaFormateada As String
    Dim regexProd As Object, matchesProd As Object
    
    Dim dicConsolidado As Object          ' key: hora|PRODUCTO_NORMALIZADO  value: sumaCantidad
    Dim dicDisplay As Object              ' key: hora|PRODUCTO_NORMALIZADO  value: producto para mostrar
    Dim claveHP As String, k As Variant
    Dim partes() As String
    
    Dim match As Object, subMatch As Object

    ' Solicitar al usuario la fecha
    fechaSeleccionada = InputBox("Ingrese la fecha en formato dd/mm/aaaa (ejemplo: " & Format(Date, "dd/mm/yy") & "):", "Seleccionar Fecha")
    If Not IsDate(fechaSeleccionada) Then
        MsgBox "Fecha no válida.", vbExclamation
        Exit Sub
    End If
    fechaFormateada = Format(CDate(fechaSeleccionada), "yyyy/mm/dd")

    ' Construir URL
    url = "http://10.211.10.250:3000/reporte/Ventas/ProductoHora?restaurante=19&fecha_inicio=" & fechaFormateada & _
          "&fecha_fin=" & fechaFormateada & _
          "&estado=&canal=undefined&usuario=2C91C9BA-B49A-EE11-8925-6045BDBB68D1&cadena=1&cajero=undefined"

    ' Obtener contenido del link
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then
        MsgBox "Error al acceder al link. Código HTTP: " & http.Status, vbExclamation
        Exit Sub
    End If
    
    respuesta = http.responseText
    Set http = Nothing
    
    If Trim$(respuesta) = "" Then
        MsgBox "La respuesta del servidor está vacía. No se pueden importar datos.", vbExclamation
        Exit Sub
    End If

    ' Preparar hoja "Venta por hora"
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Venta por hora")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = "Venta por hora"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ws.Range("A1:C1").Value = Array("Hora", "Producto", "Cantidad")
    fila = 2

    ' Diccionarios para consolidar
    Set dicConsolidado = CreateObject("Scripting.Dictionary")
    Set dicDisplay = CreateObject("Scripting.Dictionary")

    ' Regex principal para buscar bloques por "cabecera" y "cuerpo"
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = """cabecera"":""([^""]+)"".*?""cuerpo"":\[(.*?)\]"
    End With

    ' Regex para buscar productos y cantidades dentro de un bloque "cuerpo"
    Set regexProd = CreateObject("VBScript.RegExp")
    With regexProd
        .Global = True
        .IgnoreCase = True
        .Pattern = """producto"":""([^""]+)"".*?""cantidad"":([0-9]+)"
    End With

    Set matches = regex.Execute(respuesta)

    If matches.count = 0 Then
        MsgBox "Error Crítico: No se encontraron bloques de hora con el patrón. Verifique la respuesta del servidor.", vbCritical
        Exit Sub
    End If

    ' Parse + consolidación
    For Each match In matches
        If match.SubMatches.count > 1 Then
            horaActual = match.SubMatches(0)
            bloque = match.SubMatches(1)

            Set matchesProd = regexProd.Execute(bloque)

            For Each subMatch In matchesProd
                If subMatch.SubMatches.count > 1 Then
                    producto = Replace(subMatch.SubMatches(0), "*", "")
                    producto = Trim$(producto)
                    
                    ' Normalización para consolidar (misma hora + mismo producto)
                    productoKey = UCase$(producto) ' clave estable
                    claveHP = CStr(horaActual) & "|" & productoKey
                    
                    cantidad = CLng(subMatch.SubMatches(1))
                    
                    If Not dicConsolidado.Exists(claveHP) Then
                        dicConsolidado.Add claveHP, cantidad
                        dicDisplay.Add claveHP, producto ' guardo cómo mostrarlo
                    Else
                        dicConsolidado(claveHP) = CLng(dicConsolidado(claveHP)) + cantidad
                    End If
                End If
            Next subMatch
        End If
    Next match

    ' Volcar a hoja ya consolidado
    For Each k In dicConsolidado.keys
        partes = Split(CStr(k), "|")
        ws.Cells(fila, 1).Value = partes(0)                ' Hora
        ws.Cells(fila, 2).Value = dicDisplay(k)            ' Producto
        ws.Cells(fila, 3).Value = dicConsolidado(k)        ' Cantidad sumada
        fila = fila + 1
    Next k

    ws.Columns("A:C").AutoFit

End Sub


