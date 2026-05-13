Attribute VB_Name = "Módulo20"
Option Explicit

' Módulo 20: Proceso Principal
' Calcula el gran total de la columna "Total" en "Tabla de venta"
' y lo muestra en el mensaje final.

Sub ProcesoCompletoVentaDiaria()

    Dim wsTemp As Worksheet
    Dim wsVenta As Worksheet
    Dim totalRange As Range
    Dim grandTotal As Double
    Dim colTotal As Long
    Dim filaInicio As Long, filaFin As Long
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    ' Paso 1: Descargar datos del día.
    Call ImportarVentasPorHoraFormateada

    ' Comprobar si la hoja de datos se creó correctamente.
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Sheets("Venta por hora")
    On Error GoTo ErrorHandler

    If wsTemp Is Nothing Then
        MsgBox "El módulo 'ImportarVentasPorHoraFormateada' no creó la hoja 'Venta por hora'. No se puede continuar.", vbCritical
        GoTo CleanExit
    End If

    lastRow = wsTemp.Cells(wsTemp.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "La hoja 'Venta por hora' fue creada pero está vacía. No hay datos para procesar.", vbExclamation
        GoTo CleanExit
    End If

    ' Paso 2: Procesar datos y crear la tabla final.
    Call ProcesarVentaDiariaYCrearTabla

    ' Paso 3: Calcular el Gran Total de la tabla generada.
    Set wsVenta = ThisWorkbook.Sheets("Tabla de venta")

    filaInicio = 2
    filaFin = wsVenta.Cells(wsVenta.Rows.count, "A").End(xlUp).Row
    colTotal = wsVenta.Cells(1, wsVenta.Columns.count).End(xlToLeft).Column

    If filaFin >= filaInicio Then
        Set totalRange = wsVenta.Range(wsVenta.Cells(filaInicio, colTotal), wsVenta.Cells(filaFin, colTotal))
        grandTotal = Application.WorksheetFunction.Sum(totalRange)
    Else
        grandTotal = 0
    End If

    ' Paso 4: Ocultar todas las hojas excepto las de resultado final.
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case "Tabla de venta", "Combos pollo", "Combos no encontrados"
                ws.Visible = xlSheetVisible
            Case Else
                ws.Visible = xlSheetHidden
        End Select
    Next ws

    ' Activar la hoja principal de resultados para el usuario.
    ThisWorkbook.Sheets("Tabla de venta").Activate

    MsgBox "Proceso completado con éxito." & vbCrLf & vbCrLf & _
           "Total de unidades de pollo proyectadas: " & Format(grandTotal, "#,##0"), vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error inesperado durante la ejecución: " & vbCrLf & Err.Description, vbCritical
    Resume CleanExit

End Sub


