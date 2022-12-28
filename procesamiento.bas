Attribute VB_Name = "procesamiento"
Option Explicit

Function insertar_funciones_stock(archivo As String)
    Dim last_row As Long
    Workbooks(archivo).Activate
    With Workbooks(archivo).Worksheets(1)
        .Range("d2").FormulaLocal = "=BUSCARV(A2;vmd.xlsx!$A$2:$d$8757;4;0)"
        .Range("e2").FormulaLocal = "=+BUSCARV(A2;vmd.xlsx!$A$2:$d$8757;3;0)"
        .Range("f2").FormulaLocal = "=+REDONDEAR.MAS(D2*10;0)"
        .Range("g2").FormulaLocal = "=SI(C2>F2;0;F2-C2)"
        last_row = .Range("a1").End(xlDown).Row
        .Cells(2, 4).Resize(1, 4).Copy
        .Cells(2, 4).Resize(last_row - 1, 4).PasteSpecial (xlFormulas)
    End With

End Function

Function insertar_funciones_vmd(archivo As String, reuse_vmd As Boolean)
    Dim last_row As Long
    Workbooks(archivo).Activate
    With Workbooks(archivo).Worksheets(1)
    
        If Not reuse_vmd Then
            .Range("e1").Value = "¿Está en el robot?"
            .Range("f1").Value = "¿Falta en el robot?"
        End If
        
        .Range("e2").FormulaLocal = "=+BUSCARV(A2;stock.xlsx!$A$2:$A$4336;1;0)"
        .Range("f2").FormulaLocal = "=+SI(Y(ESERROR(e2);d2>0);""Si, falta en el robot"";""No, no falta"")"
        last_row = .Range("a1").End(xlDown).Row
        .Range("e2:f2").Copy
        .Cells(3, 5).Resize(last_row - 2, 2).PasteSpecial (xlFormulas)
        
        If Not reuse_vmd Then
            .Range("d1:d" & last_row).Copy
            .Range("e1:f" & last_row).PasteSpecial (xlFormats)
        End If
        Application.CutCopyMode = False
        
        If reuse_vmd Then
            .Range("i1:n" & last_row).ClearContents
            .Range("i1:n" & last_row).ClearFormats
        End If
        .Range("a1:f" & last_row).AdvancedFilter xlFilterCopy, .Range("o1:o2"), .Range("i1")
        .Range("i1").ColumnWidth = 40
        .Range("a1:a" & last_row).EntireRow.AutoFit
        Workbooks(archivo).Save
    End With

End Function

Sub copiar_faltantes_en_robot(path_file_vmd As String, path_file_stock As String)

    If Workbooks(path_file_vmd).Worksheets(1).Range("i2").Value = "" Then
        MsgBox "No hay faltantes de stock en el robot", , "Faltantes de stock"
        Exit Sub
    End If
    
    Dim stock_ws As Worksheet, vmd_ws As Worksheet
    Dim last_row_vmd As Long, next_row_stock As Long, last_row_stock
    Set stock_ws = Workbooks(path_file_stock).Worksheets(1)
    Set vmd_ws = Workbooks(path_file_vmd).Worksheets(1)
    
    last_row_vmd = vmd_ws.Range("i1").End(xlDown).Row
    next_row_stock = stock_ws.Range("a1").End(xlDown).Offset(3).Row
    
    With stock_ws.Range("a" & next_row_stock).Offset(-1)
        .Value = "FALTANTES EN EL ROBOT:"
    End With
    
    vmd_ws.Range("i2:i" & last_row_vmd).Copy
    stock_ws.Range("a" & next_row_stock).PasteSpecial xlPasteAll
    vmd_ws.Range("l2:l" & last_row_vmd).Copy
    stock_ws.Range("d" & next_row_stock).PasteSpecial xlPasteAll
    vmd_ws.Range("k2:k" & last_row_vmd).Copy
    stock_ws.Range("e" & next_row_stock).PasteSpecial xlPasteAll
    vmd_ws.Range("j2:j" & last_row_vmd).Copy
    stock_ws.Range("b" & next_row_stock).PasteSpecial xlPasteValues
    
        
    last_row_stock = stock_ws.Range("a" & next_row_stock).End(xlDown).Row
    
    stock_ws.Range("c" & next_row_stock & ":c" & last_row_stock).Value = 0
    
    stock_ws.Range("f2:g2").Copy
    stock_ws.Range("f" & next_row_stock & ":g" & last_row_stock).PasteSpecial xlPasteFormulas
    stock_ws.Range("a2:g2").Copy
    stock_ws.Range("a" & next_row_stock & ":g" & last_row_stock).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    Workbooks(path_file_stock).Save
End Sub
