Attribute VB_Name = "Scraping_Data"
Sub Scraping_Data_Base_de_Dados()

    Dim driver As New ChromeDriver
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    
    Set ws = ActiveWorkbook.Worksheets("Raw")
    
    driver.Get "https://www.fundsexplorer.com.br/ranking"
    
    
    With ws
        
        .Range(.Range("D4"), .Range("D4").End(xlDown).End(xlToRight)).ClearContents
        
        driver.FindElementById("table-ranking").AsTable.ToExcel .Range("D4")
        
    End With
    
    ActiveWorkbook.Worksheets("Home").Range("J9").Value2 = Date
    ActiveWorkbook.Worksheets("Home").Range("J10").Value2 = Time
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


Sub Scraping_Emissao()

    Dim driver As New ChromeDriver
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    
    Set ws = ActiveWorkbook.Worksheets("Emissão")
    
    driver.Get "https://www.fundsexplorer.com.br/emissoes-ipos"
    
    
    With ws
    
        'limpar conteudo antes de inserir
        .Range(.Range("B7").End(xlDown), .Range("B7").End(xlToRight)).ClearContents
        
        driver.FindElementById("DataTables_Table_0").AsTable.ToExcel .Range("B7")
        
    End With
    
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

