Attribute VB_Name = "Arruma_Top15"
Sub main_TOP15()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call Copia_para_Top_15
    Call Converte_para_Tabela_Top15
    Call Formata_Numeros_Top15

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


Sub Converte_para_Tabela_Top15()
    
    Dim ws As Worksheet
    Dim i, cont As Integer
    Set ws = ActiveWorkbook.Worksheets("Top 15")
    
    cont = 1
    
    With ws
    
        .Range("B7").Value2 = "Ranking"
        For i = 8 To 22
            .Cells(i, "B").Value2 = cont
            cont = cont + 1
        Next i
    
        .ListObjects.Add(xlSrcRange, .Range("$B$7:$AE$22"), , xlYes).Name = "Tab_top_15"
        .ListObjects("Tab_top_15").TableStyle = "TableStyleMedium21"
    End With
    
End Sub

Sub Formata_Numeros_Top15()

    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Worksheets("Top 15")
    
    With ws
        
        .Range("E8:E22").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range("F8:F22").NumberFormat = "###,000"
        .Range("G8:G22").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range("H8:H22").NumberFormat = "####0.0#####\% "
        .Range("I8:J22").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range("L8:U22").NumberFormat = "####0.0#####\% "
        
    End With
    

End Sub
