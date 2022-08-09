Attribute VB_Name = "btn_Filtra_Fundos_Principal"
Sub btn_Filtra_Fundos_MASTER()

    'Atualiza base de dados
    Call Scraping_Data_Base_de_Dados
    
    'Ajeitando base de dados
    Call Main_Filtra_Especial
    
    'Aplicando estrategia
    Call Main_Aplicando_Estrategia_G_rank
    
    
    'Atualizando aba base de dados
    Call main_Base_de_Dados
    
    'Atualizando aba TOP 15
    Call main_TOP15
    
    MsgBox ("Bases de Dados atualizadas e filtradas segundo a estratégia. Favor Conferir a guia TOP 15 ou Base de Dados")
    
End Sub
