Attribute VB_Name = "InsertarCodigo"
Option Explicit

Sub ExportarTodoElCodigoVBA()
    Dim vbComp As Object
    Dim rutaBase As String
    Dim extension As String
    Dim nombreArchivo As String
    
    rutaBase = ThisWorkbook.Path & "\vba-exportado\"
    
    If Dir(rutaBase, vbDirectory) = "" Then
        MkDir rutaBase
    End If
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        
        Select Case vbComp.Type
            Case 1 ' Módulo estándar
                extension = ".bas"
            Case 2 ' Clase
                extension = ".cls"
            Case 3 ' UserForm
                extension = ".frm"
            Case 100 ' ThisWorkbook / Hojas
                extension = ".cls"
            Case Else
                extension = ".txt"
        End Select
        
        nombreArchivo = rutaBase & vbComp.Name & extension
        
        On Error Resume Next
        vbComp.Export nombreArchivo
        On Error GoTo 0
    Next vbComp
    
    MsgBox "Código VBA exportado en:" & vbCrLf & rutaBase, vbInformation
End Sub
