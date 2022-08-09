Attribute VB_Name = "Aplicando_Estrategia_G_rank"
Sub Main_Aplicando_Estrategia_G_rank()
Attribute Main_Aplicando_Estrategia_G_rank.VB_ProcData.VB_Invoke_Func = " \n14"
'

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Dim i, cont As Integer

    Set ws = ActiveWorkbook.Worksheets("Raw")

    ActiveWorkbook.Worksheets("Raw").Range("D4:AF4").AutoFilter
    
    'ORDENANDO E PONTUANDO OS DY
    ActiveWorkbook.Worksheets("Raw").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Raw").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "I4"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
        
    With ActiveWorkbook.Worksheets("Raw").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    With ws
        cont = 1
        
        For i = 5 To .Range("D5").End(xlDown).Row
            .Cells(i, "AD").Value2 = cont
            cont = cont + 1
        Next i
    End With
    
    'ORDENANDO E PONTUANDO FILTRO P/VPA
    
    ActiveWorkbook.Worksheets("Raw").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Raw").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "V4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
        
    With ActiveWorkbook.Worksheets("Raw").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    With ws
        cont = 1
        
        For i = 5 To .Range("D5").End(xlDown).Row
            .Cells(i, "AE").Value2 = cont
            cont = cont + 1
        Next i
    End With
    
    'AQUI COMEÇA AS SOMAS E A ORDENAÇÃO (PONTUAÇÃO FINAL)
    
    With ws
    
        For i = 5 To .Range("D5").End(xlDown).Row
            .Cells(i, "AF").Value2 = .Cells(i, "AD").Value2 + .Cells(i, "AE").Value2
        Next i
    
    End With
    
    ActiveWorkbook.Worksheets("Raw").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Raw").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "AF4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
        
    With ActiveWorkbook.Worksheets("Raw").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveWorkbook.Worksheets("Raw").Range("D4:AF4").AutoFilter
    
    Application.ScreenUpdating = True
    
End Sub
