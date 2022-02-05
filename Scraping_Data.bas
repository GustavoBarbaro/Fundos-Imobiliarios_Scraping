Attribute VB_Name = "Scraping_Data"
Sub Scraping_Data()

    Dim driver As New ChromeDriver
    Dim i, j, verificador As Integer
    Dim nome As String
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    i = 5
    Set ws = ActiveWorkbook.Worksheets("Raw")
    
    driver.Get "https://www.fundsexplorer.com.br/ranking"
    
    verificador = 0
    
    With ws
        
        .Range("B10").Value2 = Time
    
        'passar por cada linha da tabela
        For Each fundo In driver.FindElementById("table-ranking").FindElementsByTag("tr")
        
            'pular o primeiro tr que eh o header da tabela
            If (verificador = 0) Then
                verificador = 1
                GoTo proximo
            End If
            
            
            For j = 1 To 26
                .Cells(i, (j + 3)).Value2 = fundo.FindElementsByTag("td").Item(j).Text
            Next j
            
        
            i = i + 1
proximo:
        Next
        .Range("B11").Value2 = Time
    End With
      
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
