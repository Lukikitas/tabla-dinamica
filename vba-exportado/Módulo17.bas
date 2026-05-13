Attribute VB_Name = "Módulo17"
' =====================================================================================
' MÓDULO DE FUNCIONES INDEPENDIENTES PARA "TABLA DE VENTA"
' =====================================================================================
' Las macros en este módulo se pueden ejecutar directamente (Alt + F8)
' sobre una hoja "Tabla de venta" ya generada.

Option Explicit

' Constantes para los nombres de las hojas, para evitar errores de tipeo.
Private Const NOMBRE_HOJA As String = "Tabla de venta"

' -------------------------------------------------------------------------------------
' 1) MACRO: APLICAR AJUSTE DE PORCENTAJE
' -------------------------------------------------------------------------------------
Public Sub Aplicar_Ajuste_Porcentaje()
    Dim ws As Worksheet
    Dim filaInicio As Long, filaFin As Long
    Dim colInicioDatos As Long, colFinDatos As Long
    Dim celTotal As Range
    Dim colTotal As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(NOMBRE_HOJA)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja '" & NOMBRE_HOJA & "'.", vbCritical
        Exit Sub
    End If

    filaInicio = 2
    filaFin = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' Detectar la columna "Total" por encabezado en fila 1 (robusto)
    Set celTotal = ws.Rows(1).Find(What:="Total", LookIn:=xlValues, LookAt:=xlWhole, _
                                   SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If Not celTotal Is Nothing Then
        colTotal = celTotal.Column
    Else
        colTotal = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    End If

    colInicioDatos = 3                    ' Columna C
    colFinDatos = colTotal - 1            ' Excluir "Total"

    If colFinDatos < colInicioDatos Then
        MsgBox "No hay columnas de datos para ajustar.", vbExclamation
        Exit Sub
    End If

    Call Logic_AplicarAjuste(ws, filaInicio, filaFin, colInicioDatos, colFinDatos)

    ' Ocultar columnas CV:DL al final
    OcultarColumnasCV_DL ws

    MsgBox "Ajuste de porcentaje aplicado correctamente.", vbInformation
End Sub

Private Sub Logic_AplicarAjuste(ws As Worksheet, filaInicio As Long, filaFin As Long, colInicioDatos As Long, colFinDatos As Long)
    Dim fila As Long, col As Long
    Dim ajuste As Double
    For fila = filaInicio To filaFin
        ajuste = 0
        If IsNumeric(ws.Cells(fila, 2).Value) Then
            ajuste = CDbl(ws.Cells(fila, 2).Value)
        End If
        If ajuste <> 0 Then
            For col = colInicioDatos To colFinDatos
                If IsNumeric(ws.Cells(fila, col).Value) Then
                    ws.Cells(fila, col).Value = CDbl(ws.Cells(fila, col).Value) * (1 + ajuste)
                End If
            Next col
        End If
    Next fila
End Sub

' -------------------------------------------------------------------------------------
' 2) MACRO: AGREGAR TEXTO DE CAJONES EN CADA CELDA Y EN "Total"
' -------------------------------------------------------------------------------------
Public Sub Agregar_Texto_Cajones()
    Dim ws As Worksheet
    Dim filaInicio As Long, filaFin As Long
    Dim colInicio As Long, colFin As Long
    Dim celTotal As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(NOMBRE_HOJA)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja '" & NOMBRE_HOJA & "'.", vbCritical
        Exit Sub
    End If

    filaInicio = 2
    filaFin = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    colInicio = 3 ' Columna C

    ' Detectar la columna "Total" por encabezado en fila 1 (robusto)
    Set celTotal = ws.Rows(1).Find(What:="Total", LookIn:=xlValues, LookAt:=xlWhole, _
                                   SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If Not celTotal Is Nothing Then
        colFin = celTotal.Column           ' Incluye la columna Total
    Else
        colFin = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    End If

    If colFin < colInicio Then
        MsgBox "No hay columnas de datos para procesar.", vbExclamation
        Exit Sub
    End If

    Call Logic_DividirYAgregarTexto(ws, filaInicio, filaFin, colInicio, colFin)
    Call Logic_AgregarCajonesAlTotal(ws, filaInicio, filaFin, colFin)

    ws.Columns.AutoFit

    ' Ocultar columnas CV:DL al final
    OcultarColumnasCV_DL ws
End Sub

Private Sub Logic_DividirYAgregarTexto(ws As Worksheet, filaInicio As Long, filaFin As Long, colInicio As Long, colFin As Long)
    Dim fila As Long, col As Long
    Dim valor As Double, divisor As Double
    Dim resultado As Long, texto As String
    Dim valorOriginal As String

    For fila = filaInicio To filaFin
        divisor = Logic_ObtenerDivisor(CStr(ws.Cells(fila, 1).Value))
        ' Recorrer hasta la penúltima columna (antes del Total)
        For col = colInicio To colFin - 1
            valorOriginal = CStr(ws.Cells(fila, col).Value)

            ' Extraer solo el número si ya tiene texto tipo "123 (R)"
            If InStr(valorOriginal, "(") > 0 Then
                valor = Val(Trim(Left$(valorOriginal, InStr(valorOriginal, "(") - 1)))
            ElseIf IsNumeric(ws.Cells(fila, col).Value) Then
                valor = CDbl(ws.Cells(fila, col).Value)
            Else
                valor = 0
            End If

            If valor > 0 And divisor > 0 Then
                resultado = Round(valor / divisor, 0)
                texto = Logic_ObtenerTexto(resultado)
                If texto <> "" Then
                    ws.Cells(fila, col).Value = CStr(valor) & " (" & texto & ")"
                Else
                    ws.Cells(fila, col).Value = valor
                End If
            End If
        Next col
    Next fila
End Sub

Private Sub Logic_AgregarCajonesAlTotal(ws As Worksheet, filaInicio As Long, filaFin As Long, colTotal As Long)
    Dim fila As Long, valor As Double, divisor As Double, texto As String
    Dim valorOriginal As String

    For fila = filaInicio To filaFin
        valorOriginal = CStr(ws.Cells(fila, colTotal).Value)
        If InStr(valorOriginal, "(") > 0 Then
            ' Si ya tiene texto, extraemos el número
            valor = Val(Trim(Left$(valorOriginal, InStr(valorOriginal, "(") - 1)))
        ElseIf IsNumeric(ws.Cells(fila, colTotal).Value) Then
            valor = CDbl(ws.Cells(fila, colTotal).Value)
        Else
            valor = 0
        End If

        If valor > 0 Then
            divisor = Logic_ObtenerDivisorCajones(CStr(ws.Cells(fila, 1).Value))
            If divisor > 1 Then
                texto = Format$(valor / divisor, "0.0") & " cajones"
                ws.Cells(fila, colTotal).Value = CStr(valor) & " (" & texto & ")"
            Else
                ws.Cells(fila, colTotal).Value = valor
            End If
        End If
    Next fila
End Sub

' -------------------------------------------------------------------------------------
' 3) MACRO: FORMATEAR LA TABLA CON ESTILO "KFC"
' -------------------------------------------------------------------------------------
Public Sub Formatear_Tabla_Como_KFC()
    Dim ws As Worksheet
    Dim filaInicio As Long, filaFin As Long
    Dim colTotal As Long
    Dim celTotal As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(NOMBRE_HOJA)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja '" & NOMBRE_HOJA & "'.", vbCritical
        Exit Sub
    End If

    filaInicio = 2
    filaFin = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' Detectar "Total" en fila 1
    Set celTotal = ws.Rows(1).Find(What:="Total", LookIn:=xlValues, LookAt:=xlWhole, _
                                   SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If Not celTotal Is Nothing Then
        colTotal = celTotal.Column
    Else
        colTotal = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    End If

    Call Logic_FormatearTabla(ws, filaInicio, filaFin, colTotal)

    ' Ocultar columnas CV:DL al final
    OcultarColumnasCV_DL ws

    MsgBox "Formato de tabla aplicado.", vbInformation
End Sub

Private Sub Logic_FormatearTabla(ws As Worksheet, filaInicio As Long, filaFin As Long, colTotal As Long)
    ' Formato general
    With ws.Range("A1", ws.Cells(filaFin, colTotal))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' Encabezado (fila 1)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, colTotal))
        .Interior.Color = RGB(207, 18, 39) ' Rojo KFC
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .RowHeight = 30
    End With

    ' Columna de ítems (A2:A...)
    With ws.Range(ws.Cells(filaInicio, 1), ws.Cells(filaFin, 1))
        .Interior.Color = RGB(230, 230, 230)
        .Font.Bold = True
    End With

    ' Columna "Total"
    With ws.Range(ws.Cells(filaInicio, colTotal), ws.Cells(filaFin, colTotal))
        .Interior.Color = RGB(245, 245, 245)
        .Font.Bold = True
    End With

    ' Datos (B..colTotal-1)
    If colTotal > 2 Then
        With ws.Range(ws.Cells(filaInicio, 2), ws.Cells(filaFin, colTotal - 1))
            .Interior.Color = RGB(255, 255, 255)
            .Font.Bold = False
        End With
    End If

    ws.Columns.AutoFit
End Sub

' -------------------------------------------------------------------------------------
' FUNCIONES AUXILIARES (USADAS POR LAS MACROS ANTERIORES)
' -------------------------------------------------------------------------------------
Private Function Logic_ObtenerDivisor(item As String) As Double
    Select Case UCase$(Trim$(item))
        Case "STRIPS", "P. ORIGINAL", "P. CRISPY", "POP"
            Logic_ObtenerDivisor = 18
        Case "FILETE"
            Logic_ObtenerDivisor = 7
        Case "RUSTER"
            Logic_ObtenerDivisor = 15
        Case "ALITAS"
            Logic_ObtenerDivisor = 24
        Case Else
            Logic_ObtenerDivisor = 1
    End Select
End Function

Private Function Logic_ObtenerTexto(resultado As Long) As String
    Select Case resultado
        Case 1:  Logic_ObtenerTexto = "R"
        Case 2:  Logic_ObtenerTexto = "B"
        Case 3:  Logic_ObtenerTexto = "B+R"
        Case 4:  Logic_ObtenerTexto = "T"
        Case 5:  Logic_ObtenerTexto = "T+R"
        Case 6:  Logic_ObtenerTexto = "T+B"
        Case 7:  Logic_ObtenerTexto = "T+3R"
        Case 8:  Logic_ObtenerTexto = "2T"
        Case 9:  Logic_ObtenerTexto = "2T+R"
        Case 10: Logic_ObtenerTexto = "2T+B"
        Case 11: Logic_ObtenerTexto = "2T+3R"
        Case 12: Logic_ObtenerTexto = "3T"
        Case 13: Logic_ObtenerTexto = "3T+R"
        Case 14: Logic_ObtenerTexto = "3T+B"
        Case 15: Logic_ObtenerTexto = "3T+3B"
        Case Else: Logic_ObtenerTexto = ""
    End Select
End Function

Private Function Logic_ObtenerDivisorCajones(item As String) As Double
    Select Case UCase$(Trim$(item))
        Case "ALITAS":                   Logic_ObtenerDivisorCajones = 336
        Case "FILETE":                   Logic_ObtenerDivisorCajones = 168
        Case "RUSTER":                   Logic_ObtenerDivisorCajones = 300
        Case "P. CRISPY", "P. ORIGINAL": Logic_ObtenerDivisorCajones = 144
        Case "POP":                      Logic_ObtenerDivisorCajones = 22
        Case "STRIPS":                   Logic_ObtenerDivisorCajones = 360
        Case Else:                       Logic_ObtenerDivisorCajones = 1
    End Select
End Function

' -------------------------------------------------------------------------------------
' UTILIDAD: OCULTAR COLUMNAS CV:DL
' -------------------------------------------------------------------------------------
Private Sub OcultarColumnasCV_DL(ws As Worksheet)
    On Error Resume Next
    ws.Range("CV:DL").EntireColumn.Hidden = True
    On Error GoTo 0
End Sub


