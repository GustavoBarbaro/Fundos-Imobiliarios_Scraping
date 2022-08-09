Attribute VB_Name = "Filtra_Especial"

Sub Main_Filtra_Especial()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call Tira_NA
    Call Tira_Zeros
    Call Formata_Celulas
    Call Filtra_Patrimonio_Liquido
    Call Filtra_Liquidez
    Call Ordena_Estrategia

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub


Sub Formata_Celulas()

    Dim ws As Worksheet
    Dim valor As Double
    Dim rng As Range

    Set ws = ActiveWorkbook.Worksheets("Raw")
    
    With ws
        
        'preco atual
        Set rng = .Range(.Range("F5"), .Range("F5").End(xlDown))
        For Each i In rng
            valor = i.Value2
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'dividendo
        Set rng = .Range(.Range("H5"), .Range("H5").End(xlDown))
        For Each i In rng
            valor = i.Value2
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'patrimonio liquido
        Set rng = .Range(.Range("T5"), .Range("T5").End(xlDown))
        For Each i In rng
            valor = i.Value2
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'VPA
        Set rng = .Range(.Range("U5"), .Range("U5").End(xlDown))
        For Each i In rng
            valor = i.Value2
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'P/VPA
        Set rng = .Range(.Range("V5"), .Range("V5").End(xlDown))
        For Each i In rng
            valor = i.Value2
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'DY
        Set rng = .Range(.Range("I5"), .Range("I5").End(xlDown))
        For Each i In rng
            On Error GoTo end_filtra
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'DY (3M) A
        Set rng = .Range(.Range("J5"), .Range("J5").End(xlDown))
        For Each i In rng
            On Error GoTo end_filtra
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'DY (6M) A
        Set rng = .Range(.Range("K5"), .Range("K5").End(xlDown))
        For Each i In rng
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'DY (12M) A
        Set rng = .Range(.Range("L5"), .Range("L5").End(xlDown))
        For Each i In rng
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'DY (3M) Media
        Set rng = .Range(.Range("M5"), .Range("M5").End(xlDown))
        For Each i In rng
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'DY (6M) Media
        Set rng = .Range(.Range("N5"), .Range("N5").End(xlDown))
        For Each i In rng
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'DY (12M) Media
        Set rng = .Range(.Range("O5"), .Range("O5").End(xlDown))
        For Each i In rng
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'DY Ano
        Set rng = .Range(.Range("P5"), .Range("P5").End(xlDown))
        For Each i In rng
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'Variacao de preco
        Set rng = .Range(.Range("Q5"), .Range("Q5").End(xlDown))
        For Each i In rng
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'rentabilidade no periodo
        Set rng = .Range(.Range("R5"), .Range("R5").End(xlDown))
        For Each i In rng
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
        'rentabilidade acumulada
        Set rng = .Range(.Range("S5"), .Range("S5").End(xlDown))
        For Each i In rng
            valor = Left(i.Value2, Len(i.Value2) - 1) 'tirando o sinal de %
            i.Value2 = valor
        Next i
        Set rng = Nothing
        
    End With
    
end_filtra:

End Sub


Sub Tira_NA()

    Dim ws As Worksheet
    Dim rng As Range
    Dim i As Integer
    
    
    Set ws = ActiveWorkbook.Worksheets("Raw")
    
    With ws
    
        Set rng = .Range(.Range("D5"), .Range("V5").End(xlDown))
        
        While (1)
        
            If (rng.Find("N/A") Is Nothing) Then
                GoTo break_while
            End If
    
            rng.Find("N/A").EntireRow.Delete
        
        Wend
break_while:
    End With

End Sub

Sub Tira_Zeros()

    Dim ws As Worksheet
    Dim rng As Range
    
    
    Set ws = ActiveWorkbook.Worksheets("Raw")
    
    With ws
    
        Set rng = .Range(.Range("J5"), .Range("P5").End(xlDown))
                
        While (1)

            If (rng.Find("0,00%") Is Nothing) Then
                GoTo break_while
            End If
            rng.Find("0,00%").EntireRow.Delete

        Wend
break_while:

        Set rng = Nothing
        Set rng = .Range(.Range("H5"), .Range("H5").End(xlDown))
        
        While (1)

            If (rng.Find("R$ 0,00") Is Nothing) Then
                GoTo break_while_1
            End If
            rng.Find("R$ 0,00").EntireRow.Delete

        Wend
break_while_1:
        
    End With

End Sub

Sub Filtra_Patrimonio_Liquido()

    Dim ws As Worksheet
    Dim rng As Range
    Dim max_PatrimonioLiquido As Double
    
    
    max_PatrimonioLiquido = ActiveWorkbook.Worksheets("Home").Range("J12").Value2
    
    
    Set ws = ActiveWorkbook.Worksheets("Raw")
    
    With ws
    
        Set rng = .Range(.Range("T5"), .Range("T5").End(xlDown))
                
        For Each i In rng
            If i.Value2 < max_PatrimonioLiquido Then
                i.Value2 = 0
            End If
        Next i
        
                
        While (1)

            If (rng.Find("0", Lookat:=xlWhole) Is Nothing) Then
                GoTo break_while
            End If
        
            rng.Find("0", Lookat:=xlWhole).EntireRow.Delete

        Wend
    
    End With
break_while:
End Sub

Sub Filtra_Liquidez()

    Dim ws As Worksheet
    Dim rng As Range
    Dim min_Liquidez As Double
    
    
    min_Liquidez = 3402
    
    
    Set ws = ActiveWorkbook.Worksheets("Raw")
    
    With ws
    
        Set rng = .Range(.Range("G5"), .Range("G5").End(xlDown))
                
        For Each i In rng
            If i.Value2 <= min_Liquidez Then
                i.Value2 = 0
            End If
        Next i
        
                
        While (1)

            If (rng.Find("0", Lookat:=xlWhole) Is Nothing) Then
                GoTo break_while
            End If
        
            rng.Find("0", Lookat:=xlWhole).EntireRow.Delete

        Wend
            
    End With
break_while:
End Sub

Sub Ordena_Estrategia()

    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ActiveWorkbook.Worksheets("Raw")
    
    With ws
    
        .Range("AD4").Value2 = "Pont. DY"
        .Range("AE4").Value2 = "Pont. P/VPA"
        .Range("AF4").Value2 = "Pont. FINAL"
    
    End With

End Sub
