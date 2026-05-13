Attribute VB_Name = "Módulo6"
Option Explicit

Private Const SH_RESUMEN As String = "Resumen de 4 semanas"
Private Const SH_TABLA As String = "Tabla de venta"
Private Const SH_COMBOS_POLLO As String = "Combos pollo"

' === Publico: llamá esto al final de tu proceso de promedio/semanas ===
Public Sub ImprimirTop5Combos_DesdeResumen4Semanas_Horas()
    If Not SheetExists(SH_RESUMEN) Then
        MsgBox "No existe la hoja '" & SH_RESUMEN & "'.", vbExclamation
        Exit Sub
    End If
    If Not SheetExists(SH_TABLA) Then
        MsgBox "No existe la hoja '" & SH_TABLA & "'.", vbExclamation
        Exit Sub
    End If
    If Not SheetExists(SH_COMBOS_POLLO) Then
        MsgBox "No existe la hoja '" & SH_COMBOS_POLLO & "'. Generá/extraé combos pollo primero.", vbExclamation
        Exit Sub
    End If

    Dim dCombos As Object: Set dCombos = GetDicCombosPollo() ' key norm -> nombre original
    If dCombos.count = 0 Then
        MsgBox "La hoja 'Combos pollo' no tiene combos cargados.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_RESUMEN)

    ' Detectar columnas por encabezado (fila 1)
    Dim colHora As Long, colProd As Long, colCant As Long
    colHora = DetectCol(ws, Array("HORA"))
    colProd = DetectCol(ws, Array("PRODUCTO", "PLU", "ARTICULO", "ARTÍCULO", "NOMBRE"))
    colCant = DetectCol(ws, Array("CANTIDAD", "QTY", "CANT"))

    ' Fallback típico A=Hora B=Producto C=Cantidad
    If colProd = 0 Then colProd = 2
    If colCant = 0 Then colCant = 3
    If colHora = 0 Then colHora = 1

    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.count, colProd).End(xlUp).Row
    If lastR < 2 Then
        MsgBox "'" & SH_RESUMEN & "' no tiene datos.", vbExclamation
        Exit Sub
    End If

    ' Acumular sumatoria por combo a lo largo de las horas
    Dim dSum As Object: Set dSum = CreateObject("Scripting.Dictionary")

    Dim r As Long, prod As String, k As String, qty As Double
    For r = 2 To lastR
        prod = CStr(ws.Cells(r, colProd).Value)
        k = NormalizeKey(prod)
        If Len(k) > 0 And dCombos.Exists(k) Then
            qty = Val(ws.Cells(r, colCant).Value)
            dSum(k) = Val(dSum(k)) + qty
        End If
    Next r

    If dSum.count = 0 Then
        MsgBox "No se encontraron combos (de 'Combos pollo') dentro de '" & SH_RESUMEN & "'.", vbInformation
        Exit Sub
    End If

    ' Obtener top 5 ordenado
    Dim topKeys As Variant
    topKeys = GetTopNKeysByValue(dSum, 5)

    ' Imprimir debajo de Tabla de venta
    PrintTopUnderTable topKeys, dSum, dCombos, "TOP 5 COMBOS (Resumen 4 semanas) — sumatoria por hora"

    MsgBox "Top 5 combos impreso debajo de '" & SH_TABLA & "'.", vbInformation
End Sub

' =========================
' Impresión
' =========================
Private Sub PrintTopUnderTable(ByVal topKeys As Variant, ByVal dSum As Object, ByVal dCombos As Object, ByVal title As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_TABLA)

    ' Buscar última fila usada (por columna A)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastR < 1 Then lastR = 1

    Dim startR As Long: startR = lastR + 2

    ' Título
    ws.Cells(startR, "A").Value = title
    ws.Cells(startR, "A").Font.Bold = True

    ' Headers
    ws.Cells(startR + 1, "A").Value = "Rank"
    ws.Cells(startR + 1, "B").Value = "Combo"
    ws.Cells(startR + 1, "C").Value = "Apariciones"
    ws.Range(ws.Cells(startR + 1, "A"), ws.Cells(startR + 1, "C")).Font.Bold = True

    Dim i As Long, outR As Long
    outR = startR + 2

    If IsEmpty(topKeys) Then Exit Sub

    For i = LBound(topKeys) To UBound(topKeys)
        Dim k As String: k = CStr(topKeys(i))
        ws.Cells(outR, "A").Value = (i - LBound(topKeys) + 1)
        ws.Cells(outR, "B").Value = dCombos(k)     ' nombre original
        ws.Cells(outR, "C").Value = dSum(k)        ' sumatoria
        outR = outR + 1
    Next i

    ws.Columns("A:C").AutoFit
End Sub

' =========================
' Combos pollo dictionary
' =========================
Private Function GetDicCombosPollo() As Object
    ' Lee "Combos pollo": col B = combo original
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_COMBOS_POLLO)
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, "B").End(xlUp).Row

    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, combo As String, k As String

    For r = 2 To lastR
        combo = CStr(ws.Cells(r, "B").Value)
        k = NormalizeKey(combo)
        If Len(k) > 0 Then
            If Not d.Exists(k) Then d.Add k, combo
        End If
    Next r

    Set GetDicCombosPollo = d
End Function

' =========================
' Utilidades
' =========================
Private Function NormalizeKey(ByVal s As String) As String
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbTab, " ")
    s = Trim$(s)
    NormalizeKey = UCase$(s)
End Function

Private Function SheetExists(ByVal shName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(shName) Is Nothing
    On Error GoTo 0
End Function

Private Function DetectCol(ByVal ws As Worksheet, ByVal synonyms As Variant) As Long
    ' Busca en fila 1 un header que matchee alguno de los sinónimos
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim c As Long, i As Long
    Dim h As String, target As String

    For c = 1 To lastCol
        h = NormalizeKey(CStr(ws.Cells(1, c).Value))
        If Len(h) > 0 Then
            For i = LBound(synonyms) To UBound(synonyms)
                target = NormalizeKey(CStr(synonyms(i)))
                If h = target Then
                    DetectCol = c
                    Exit Function
                End If
            Next i
        End If
    Next c

    DetectCol = 0
End Function

Private Function GetTopNKeysByValue(ByVal d As Object, ByVal n As Long) As Variant
    If d Is Nothing Or d.count = 0 Then
        GetTopNKeysByValue = Empty
        Exit Function
    End If

    Dim keys As Variant: keys = d.keys
    Dim i As Long, j As Long, tmp As Variant

    ' Ordenar desc
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If Val(d(keys(j))) > Val(d(keys(i))) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i

    ' Cortar top N
    Dim count As Long: count = (UBound(keys) - LBound(keys) + 1)
    If count > n Then
        ReDim Preserve keys(LBound(keys) To LBound(keys) + n - 1)
    End If

    GetTopNKeysByValue = keys
End Function

