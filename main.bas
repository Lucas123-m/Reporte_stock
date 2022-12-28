Attribute VB_Name = "main"
Option Explicit
Const folder_path = "C:\Users\Administracion 02\Desktop\Reposicion\"
Sub procedimiento_reporte()
    Dim ok_apertura As Boolean, key As Variant, condicion_reutilizar_vmd As Boolean
    Dim paths As Object

    Set paths = CreateObject("Scripting.Dictionary")
    paths.Add "stock", "stock.xls"
    paths.Add "vmd", "vmd.xls"
    
    
    For Each key In paths.Keys
        ok_apertura = apertura_archivos(paths(key), condicion_reutilizar_vmd)
        If Not ok_apertura Then
            Exit Sub
        End If
    Next
        
    paths("stock") = Replace(paths("stock"), "xls", "xlsx")
    paths("vmd") = Replace(paths("vmd"), "xls", "xlsx")
    
    formato.formato_archivo (paths("stock"))
    If Not condicion_reutilizar_vmd Then
        formato.formato_archivo (paths("vmd"))
    End If
    
    'Revisar si es necesaria una condicion para ejecutar el procesamiento con el formato correcto.
    procesamiento.insertar_funciones_stock (paths("stock"))
    procesamiento.insertar_funciones_vmd paths("vmd"), condicion_reutilizar_vmd
    procesamiento.copiar_faltantes_en_robot paths("vmd"), paths("stock")
    Workbooks(paths("vmd")).Close
End Sub

Function abrir_archivo(path As String, type_file As String) As Boolean
    If Not (type_file = "xls" Or type_file = "xlsx") Then
        MsgBox "El formato especificado " & type_file & " es incorrecto."
        abrir_archivo = False
        Exit Function
    End If
    
    Application.DisplayAlerts = False
    On Error GoTo error
    Workbooks.Open folder_path & path
    On Error GoTo 0
    
    If type_file = "xls" Then
        If IsWorkbookOpen(Replace(path, "xls", "xlsx")) Then
            Workbooks(Replace(path, "xls", "xlsx")).Close
        End If
        ActiveWorkbook.SaveAs folder_path & Replace(path, "xls", "xlsx"), xlOpenXMLWorkbook
    End If
    
    Application.DisplayAlerts = True
    abrir_archivo = True
    Exit Function
error:
    Dim msg
    abrir_archivo = False
    msg = "El archivo " & path & " no existe o no tiene el nombre correcto." & vbCrLf
    msg = msg & "Guardar los archivos en la carpeta 'REPOSICION' en el escritorio"
    msg = msg & " con el siguiente formato: " & vbCrLf
    msg = msg & vbCrLf & vbCrLf & "stock.xls" & vbCrLf & "vmd.xls" & vbCrLf
    MsgBox msg
    Application.DisplayAlerts = True
    Exit Function
End Function

Function apertura_archivos(ByVal path As String, ByRef condicion_reutilizar As Boolean) As Boolean
    Dim msg As String, respuesta As Variant
    If path = "stock.xls" Then
        apertura_archivos = abrir_archivo(path, "xls")
    ElseIf path = "vmd.xls" Then
        msg = "Escribir 'y' para utilizar el archivo ya existente, 'n' para abrir uno nuevo." & vbCrLf
        msg = msg & "Cualquier otra accion cerrará el programa." & vbCrLf
        respuesta = Application.InputBox(msg, "Reutilizar archivo vmd")
        
        If respuesta = "n" Then
            apertura_archivos = abrir_archivo(path, "xls")
        ElseIf respuesta = "y" Then
            apertura_archivos = abrir_archivo(Replace(path, "xls", "xlsx"), "xlsx")
        Else
            Workbooks(Replace(path, "xls", "xlsx")).Close
            MsgBox "Ha salido correctamente"
            apertura_archivos = False
        End If
        condicion_reutilizar = respuesta = "y"
    Else
        MsgBox "La ruta " & path & " es incorrecta. Se cancelará la ejecucion del programa."
        apertura_archivos = False
    End If
End Function

Function IsWorkbookOpen(file As String) As Boolean
    Dim wb As Workbook
    
    For Each wb In Workbooks
        If wb.Name = file Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next
    IsWorkbookOpen = False
End Function
