Attribute VB_Name = "Copiando_Dados_pela_Planilha"

Sub Copia_para_Top_15()

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Top 15")
    
    'limpa os conteudos antes de inserir
    With ws
        .Range(.Cells(7, "B").End(xlDown), .Cells(7, "B").End(xlToRight)).ClearContents
    End With
    
    Set ws = ActiveWorkbook.Worksheets("Raw")
    
    With ws
    
        .Range(.Cells(4, "D"), .Cells(19, "I")).Copy Destination:=ActiveWorkbook.Worksheets("Top 15").Range("C7")
        .Range(.Cells(4, "T"), .Cells(19, "V")).Copy Destination:=ActiveWorkbook.Worksheets("Top 15").Range("I7")
        .Range(.Cells(4, "J"), .Cells(19, "S")).Copy Destination:=ActiveWorkbook.Worksheets("Top 15").Range("L7")
        .Range(.Cells(4, "W"), .Cells(19, "AF")).Copy Destination:=ActiveWorkbook.Worksheets("Top 15").Range("V7")
        
    End With

End Sub

Sub Copia_para_BD()

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Base de Dados")
    
    'limpa os conteudos antes de inserir
    With ws
        .Range(.Cells(7, "B").End(xlDown), .Cells(7, "B").End(xlToRight)).ClearContents
    End With
    
    Set ws = ActiveWorkbook.Worksheets("Raw")
    
    With ws
    
        .Range(.Cells(4, "D"), .Cells(4, "I").End(xlDown)).Copy Destination:=ActiveWorkbook.Worksheets("Base de Dados").Range("B7")
        .Range(.Cells(4, "T"), .Cells(4, "V").End(xlDown)).Copy Destination:=ActiveWorkbook.Worksheets("Base de Dados").Range("H7")
        .Range(.Cells(4, "J"), .Cells(4, "S").End(xlDown)).Copy Destination:=ActiveWorkbook.Worksheets("Base de Dados").Range("K7")
        .Range(.Cells(4, "W").End(xlDown), .Cells(4, "AF")).Copy Destination:=ActiveWorkbook.Worksheets("Base de Dados").Range("U7")
        
    End With

End Sub
