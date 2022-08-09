Attribute VB_Name = "Scraping_Data"
Sub Scraping_Data()

    Dim driver As New ChromeDriver
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    
    Set ws = ActiveWorkbook.Worksheets("Raw")
    
    driver.Get "https://www.fundsexplorer.com.br/ranking"
    
    
    With ws
        
        .Range("B10").Value2 = Time
        
        .Range(.Range("D4"), .Range("D4").End(xlDown).End(xlToRight)).ClearContents
        
    
        driver.FindElementById("table-ranking").AsTable.ToExcel .Range("D4")
        

        .Range("B11").Value2 = Time
    End With
      
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
