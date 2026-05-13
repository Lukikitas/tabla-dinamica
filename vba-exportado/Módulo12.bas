Attribute VB_Name = "Módulo12"
Option Explicit

Private Function LeerEntre(ByVal txt As String, ByVal a As Long, ByVal ch As String) As String
    Dim i1 As Long, i2 As Long
    i1 = InStr(a, txt, ch): If i1 = 0 Then Exit Function
    i2 = InStr(i1 + Len(ch), txt, ch): If i2 = 0 Then Exit Function
    LeerEntre = Mid$(txt, i1 + Len(ch), i2 - (i1 + Len(ch)))
End Function

Private Function LeerNumeroDespuesDe(ByVal txt As String, ByVal clave As String, Optional ByVal desde As Long = 1) As Long
    Dim p As Long, i As Long, j As Long, c As String
    p = InStr(desde, txt, clave): If p = 0 Then Exit Function
    i = p + Len(clave)
    Do While i <= Len(txt)
        c = Mid$(txt, i, 1)
        If c Like "[0-9-]" Then Exit Do
        i = i + 1
    Loop
    j = i
    Do While j <= Len(txt) And Mid$(txt, j, 1) Like "[0-9]"
        j = j + 1
    Loop
    LeerNumeroDespuesDe = CLng(Val(Mid$(txt, i, j - i)))
End Function

Sub ImportarSemanasPorSeparado()
    ' --- Estado UI/Excel ---
    Dim prevScreen As Boolean, prevEvents As Boolean, prevStatus As Boolean, prevStatusText As Variant
    Dim prevCalc As XlCalculation

    ' --- Variables de proceso ---
    Dim fechaInput As String, fechaBase As Date
    Dim fechas(1 To 4) As String, nombresHojas(1 To 4) As String
    Dim f As Long, url As String, http As Object, respuesta As String
    Dim ws As Worksheet, fila As Long
    Dim partesHora() As String, bloque As String
    Dim s As String, i As Long, j As Long
    Dim hora As String, producto As String, cantidad As Long
    Dim pCuerpo As Long, pIni As Long, pFin As Long
    Dim items() As String, trozo As String

    On Error GoTo FIN_SEGURO

    ' -------- pedir fecha --------
    fechaInput = InputBox("Ingrese la fecha (dd/mm/yyyy):", "Fecha base")
    If Not IsDate(fechaInput) Then
        MsgBox "Formato de fecha incorrecto.", vbExclamation
        Exit Sub
    End If
    fechaBase = CDate(fechaInput)

    ' -------- congelar pantalla y mostrar “Cargando datos…” --------
    prevScreen = Application.ScreenUpdating:        Application.ScreenUpdating = False
    prevEvents = Application.EnableEvents:         Application.EnableEvents = False
    prevCalc = Application.Calculation:            Application.Calculation = xlCalculationManual
    prevStatus = Application.DisplayStatusBar:     Application.DisplayStatusBar = True
    prevStatusText = Application.StatusBar:        Application.StatusBar = "Cargando datos..."
    Application.Cursor = xlWait

    ' Fechas y nombres de hojas
    For f = 1 To 4
        fechas(f) = Format$(fechaBase - 7 * f, "yyyy/mm/dd")
        nombresHojas(f) = "Semana " & f
    Next f

    ' Recrear hojas (sin mensajes intermedios)
    Application.DisplayAlerts = False
    For f = 1 To 4
        On Error Resume Next
        ThisWorkbook.Sheets(nombresHojas(f)).Delete
        On Error GoTo 0
        ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.count)).Name = nombresHojas(f)
    Next f
    Application.DisplayAlerts = True

    ' -------- procesar 4 semanas (parser sin RegExp) --------
    For f = 1 To 4
        Set ws = ThisWorkbook.Sheets(nombresHojas(f))
        ws.Cells.Clear
        ws.Range("A1:C1").Value = Array("Hora", "Producto", "Cantidad")
        fila = 2

        url = "http://10.211.10.250:3000/reporte/Ventas/ProductoHora?restaurante=19&fecha_inicio=" & _
              fechas(f) & "&fecha_fin=" & fechas(f) & _
              "&estado=&canal=undefined&usuario=2C91C9BA-B49A-EE11-8925-6045BDBB68D1&cadena=1&cajero=undefined"

        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", url, False
        http.Send
        If http.Status <> 200 Then GoTo SiguienteF ' falló descarga: seguimos sin interrumpir

        respuesta = http.responseText

        ' Cortar por cada "hora":
        partesHora = Split(respuesta, """hora"":")
        For i = 1 To UBound(partesHora)
            s = partesHora(i)

            ' hora (primer string del trozo)
            hora = LeerEntre(s, 1, """")
            If Len(hora) = 0 Then GoTo SiguienteHora

            ' "cuerpo":[ ... ]
            pCuerpo = InStr(1, s, """cuerpo"":")
            If pCuerpo = 0 Then GoTo SiguienteHora

            pIni = InStr(pCuerpo, s, "[")
            pFin = InStr(pIni + 1, s, "]")
            If pIni = 0 Or pFin = 0 Or pFin <= pIni Then GoTo SiguienteHora

            bloque = Mid$(s, pIni + 1, pFin - pIni - 1)
            If Len(bloque) = 0 Then GoTo SiguienteHora

            ' items dentro del bloque
            items = Split(bloque, """producto"":")
            For j = 1 To UBound(items)
                trozo = items(j)
                producto = LeerEntre(trozo, 1, """")
                If Len(producto) = 0 Then GoTo ProximoItem
                producto = Replace(producto, "*", "")

                cantidad = LeerNumeroDespuesDe(trozo, """cantidad"":")
                If cantidad = 0 And InStr(1, trozo, """cantidad"":") = 0 Then GoTo ProximoItem

                ws.Cells(fila, 1).Value = hora
                ws.Cells(fila, 2).Value = producto
                ws.Cells(fila, 3).Value = cantidad
                fila = fila + 1
ProximoItem:
            Next j
SiguienteHora:
        Next i

        ws.Columns("A:C").AutoFit
SiguienteF:
    Next f
ThisWorkbook.Sheets("Tabla de venta").Range("A1").Value = "" & Format$(fechaBase, "dd/mm/yyyy")

    ' -------- pasos posteriores (silenciosos) --------
    On Error Resume Next
    Call ResumirVentas4Semanas
    Call GenerarVentaDePolloPorHora
    Call CrearTablaDeVenta
    Call VerificarCoincidenciasDeCombos
    On Error GoTo 0
ThisWorkbook.Sheets("Tabla de venta").Range("A1").Value = "" & Format$(fechaBase, "dd/mm/yyyy")
FIN_SEGURO:
    ' -------- restaurar estado UI/Excel --------
    Application.StatusBar = False
    Application.DisplayStatusBar = prevStatus
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Application.Cursor = xlDefault

    ' -------- único mensaje final --------
    If IsDate(fechaInput) Then
        MsgBox "Tabla generada correctamente para la fecha " & Format$(CDate(fechaInput), "dd/mm/yyyy") & ".", vbInformation
    End If

End Sub


