Attribute VB_Name = "Arruma_Emissoes"
Sub main_arruma_emissoes()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call Scraping_Emissao
    Call Format_Celulas
    Call Formata_Tabela
    
    MsgBox ("Dados de Emissões atualizados!")
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub



Sub Format_Celulas()

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Emissão")
    Dim valor As Double
    Dim rng As Range

    With ws
    
        'valor
        Set rng = .Range(.Range("C8"), .Range("C8").End(xlDown))
        For Each i In rng
            valor = i.Value2
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'fator de proporção
        Set rng = .Range(.Range("E8"), .Range("E8").End(xlDown))
        For Each i In rng
            On Error GoTo end_filtra
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'devolvendo simbolos
        .Range(.Range("C8"), .Range("C8").End(xlDown)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range(.Range("E8"), .Range("E8").End(xlDown)).NumberFormat = "####0.0#####\% "
        
        
    
    End With
end_filtra:
End Sub

Sub Formata_Tabela()

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Emissão")
    
    With ws
        
        .ListObjects.Add(xlSrcRange, .Range(.Range("B7").End(xlDown), .Range("B7").End(xlToRight)), , xlYes).Name = "Tab_Emissao"
        .ListObjects("Tab_Emissao").TableStyle = "TableStyleMedium20"
    
    End With

End Sub
