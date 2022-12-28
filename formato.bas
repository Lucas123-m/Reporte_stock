Attribute VB_Name = "formato"
Option Explicit

Function formato_archivo_stock(archivo As String)
    Dim last_row As Long, rng_titles As Range
    
    Workbooks(archivo).Activate
    With Workbooks(archivo).Worksheets(1)
        .Range("a1").EntireColumn.Delete
        .Range("a1").EntireColumn.Delete
        .Range("d1:e1").EntireColumn.Delete
        .Range("d1").Value = "VMD"
        .Range("e1").Value = "Stock en farmacia"
        .Range("f1").Value = "Venta redondeada a 10 dias"
        .Range("g1").Value = "Cantidad a reponer"
        
        Set rng_titles = .Range(.Range("a1"), .Range("a1").End(xlToRight))
        rng_titles.VerticalAlignment = xlCenter
        rng_titles.HorizontalAlignment = xlCenter
        rng_titles.WrapText = True
        
        last_row = .Range("c1").End(xlDown).Row
        .Cells(1, 3).Resize(last_row, 1).Copy
        Cells(1, 4).Resize(last_row, 4).PasteSpecial (xlPasteFormats)
        .Range("f1").ColumnWidth = 17
        Application.CutCopyMode = False
        .Cells(1, 1).Resize(last_row, 3).Sort Key1:=Range("a1"), Order1:=xlAscending, Header:=xlYes
        
        .Range("a1").EntireRow.Select
        With ActiveWindow
            .SplitRow = 1
            .SplitColumn = 0
            .FreezePanes = True
        End With
    End With
    
End Function

Function formato_archivo_vmd(archivo As String)
    Workbooks(archivo).Activate
    With Workbooks(archivo).Worksheets(1)
        .Columns("A:B").Delete
        .Columns("B:C").Delete
        .Columns("C").Delete
        
        .Columns("D:J").Delete
        .Columns("E:P").Delete
        
        .Range("o1").Value = "¿Falta en el robot?"
        .Range("o2").Value = "Si, falta en el robot"
        .Range("c1:c2").Copy
        .Range("o1:o2").PasteSpecial (xlFormats)
    End With
    Workbooks(archivo).Save
End Function

Function formato_archivo(archivo As String)
    If archivo = "stock.xlsx" Then
        formato_archivo_stock (archivo)
    ElseIf archivo = "vmd.xlsx" Then
        formato_archivo_vmd (archivo)
    Else
        formato_archivo = False
        Exit Function
    End If
End Function

Sub pruebas()
    formato_archivo_vmd ("vmd.xlsx")

End Sub



