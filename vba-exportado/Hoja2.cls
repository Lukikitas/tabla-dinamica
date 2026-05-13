VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Módulo: Hoja2 (Versión Final y Robusta)
' Corrige el error de desbordamiento y el error en el cálculo del rango.

Private Sub Worksheet_Change(ByVal target As Range)
    
    ' --- Validaciones Iniciales ---
    ' Salir inmediatamente si se modifican múltiples celdas.
    If target.CountLarge > 1 Then Exit Sub

    ' Salir si la celda modificada no está en el rango de interés (B2:B1000).
    If Intersect(target, Me.Range("B2:B1000")) Is Nothing Then Exit Sub

    ' Si el código llega hasta aquí, significa que UNA SOLA CELDA en el rango B2:B1000 ha cambiado.
    
    Dim fila As Long
    Dim ajuste As Double
    Dim colInicioDatos As Long, colTotal As Long
    Dim j As Long
    Dim valorOriginal As Variant
    Dim nuevoValor As Double
    Dim colBackup As Long
    Dim numColsToProcess As Long
    
    fila = target.Row
    colInicioDatos = 3 ' Columna C
    colBackup = 100    ' Columna CV
    
    ' --- Evitar bucles de eventos ---
    Application.EnableEvents = False

    ' --- INICIO DE LA CORRECCIÓN ---
    ' Determinar el ancho de la tabla usando la fila de encabezados (fila 1), que es fija y confiable.
    colTotal = Me.Cells(1, Me.Columns.count).End(xlToLeft).Column
    numColsToProcess = colTotal - colInicioDatos
    
    ' Si no hay columnas de datos (solo Producto, Ajuste y Total), no hacer nada.
    If numColsToProcess < 1 Then GoTo CleanExit
    ' --- FIN DE LA CORRECCIÓN ---
    
    
    ' --- Guardar valores originales si no existen ---
    ' La condición ahora usa un ancho de tabla confiable.
    If Application.CountA(Me.Cells(fila, colBackup).Resize(1, numColsToProcess + 1)) = 0 Then
        For j = colInicioDatos To colTotal
            ' Extraer solo el número si la celda contiene texto (ej: "1300 (3.6 cajones)")
            valorOriginal = Val(Me.Cells(fila, j).Value)
            Me.Cells(fila, colBackup + (j - colInicioDatos)).Value = valorOriginal
        Next j
    End If

    ' --- Aplicar el ajuste ---
    If IsNumeric(target.Value) Then
        ajuste = target.Value
    Else
        ajuste = 0
    End If

    ' Aplicar ajuste a las columnas de horas (no a la columna Total)
    For j = colInicioDatos To colTotal - 1
        valorOriginal = Me.Cells(fila, colBackup + (j - colInicioDatos)).Value
        If IsNumeric(valorOriginal) And Not IsEmpty(valorOriginal) Then
            nuevoValor = CDbl(valorOriginal) * (1 + ajuste)
            Me.Cells(fila, j).Value = WorksheetFunction.Round(nuevoValor, 0)
        End If
    Next j
    
    ' --- Recalcular el total numérico ---
    Me.Cells(fila, colTotal).Value = Application.WorksheetFunction.Sum(Me.Range(Me.Cells(fila, colInicioDatos), Me.Cells(fila, colTotal - 1)))
    
' Etiqueta para asegurar que los eventos siempre se reactiven
CleanExit:
    Application.EnableEvents = True
End Sub
