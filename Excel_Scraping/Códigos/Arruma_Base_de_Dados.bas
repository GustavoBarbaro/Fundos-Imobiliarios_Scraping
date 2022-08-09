Attribute VB_Name = "Arruma_Base_de_Dados"
Sub main_Base_de_Dados()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call Copia_para_BD
    Call Converte_para_Tabela_BD
    Call Formata_Numeros_BD
    

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


Sub Converte_para_Tabela_BD()
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Base de Dados")
    
    With ws
    
        .ListObjects.Add(xlSrcRange, .Range(.Range("B7").End(xlDown), .Range("B7").End(xlToRight)), , xlYes).Name = "Tab_BD"
        .ListObjects("Tab_BD").TableStyle = "TableStyleMedium17"
    End With
    
End Sub

Sub Formata_Numeros_BD()

    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Worksheets("Base de Dados")
    
    With ws
        
        .Range(.Range("D8"), .Range("D8").End(xlDown)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range(.Range("E8"), .Range("E8").End(xlDown)).NumberFormat = "###,000"
        .Range(.Range("F8"), .Range("F8").End(xlDown)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range(.Range("G8"), .Range("G8").End(xlDown)).NumberFormat = "####0.0#####\% "
        .Range(.Range("H8"), .Range("I8").End(xlDown)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range(.Range("K8"), .Range("T8").End(xlDown)).NumberFormat = "####0.0#####\% "
        
    End With
    

End Sub
