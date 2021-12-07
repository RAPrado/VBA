Sub Extrair()
    On Error GoTo Trata_Erro
    
    Dim IE As Object    
    Dim URL As String
    Dim Win As Object
    Dim Nome_Combo As String
    Dim Texto As String
    
    Dim Planilha, Tabela, el, eleRow, eleColtd, eleCol, Descricao As Object
    Dim Linha, Coluna, Linha_Html, Coluna_Html, Linha_Tabela As Integer
    Dim Inicio, Fim As String
    Dim Entrou As Boolean    
    
    Application.StatusBar = "Instanciando browser | " & Now()
    DoEvents
    
    Set IE = CreateObject("InternetExplorer.Application")
    '------------------------------------------------------------------------
    
    'Torna a Página Visível
    IE.Visible = True
    
    Application.StatusBar = "Acessando... | " & Now()
    DoEvents
    
    URL = "http://pagina.htm"
    IE.Navigate URL

    '************************************************************************************************************************************
    'Só utilizar esse loop se devido a alguma politica de segurança perde a conexão com o objeto IE, e nesse código busca pela janela aberta e reconecta com ela.
    For Each Win In CreateObject("Shell.Application").Windows
        If Win.Name Like "*Internet Explorer" Then
            Set IE = Win

            If IE.LocationURL = URL Then
                Exit For
            End If
        End If
    Next
    Set Win = Nothing
    '************************************************************************************************************************************
    
    Aguardar IE 'Aguarda tela estar carregada.
    
    'Combo 1
    Application.StatusBar = "Combo 1 | " & Now()
    DoEvents
    Nome_Combo = "ComboBox5"
    Texto = "Conteudo a selecionar"
    Seleciona_Item_Combo IE, Nome_Combo, Texto, "onchange"
    If Item_Selecionado(IE.Document.getElementbyid(Nome_Combo), Texto) Then GoTo Trata_Erro
        
    'Botão Pesquisar
    Application.StatusBar = "Pesquisando registros | " & Now()
    DoEvents
    IE.Document.getElementbyid("Button0").Click
    Aguardar IE 'Aguarda tela estar carregada.

    Set Planilha = Sheets("Nome Sheet")
    Planilha.Cells.Clear
    
    'Coluna A
    Planilha.Columns("A:A").Select
    Selection.NumberFormat = "@" 'Formato texto
    
    'Coluna C
    Planilha.Columns("C:C").Select
    Selection.NumberFormat = "dd/mm/yyyy"
        
    Planilha.Cells(1, 1).Select
    
    'Planilha.Cells.NumberFormat = "@" 'Formata todas colunas
    'Planilha.Cells.ColumnWidth = 3
    
    'Define títulos das colunas
    Planilha.Cells(1, 1) = "Coluna A"
    Planilha.Cells(1, 2) = "Coluna B"
    Planilha.Cells(1, 3) = "Coluna C"    
    
    Inicio = 1
       
    Fim = Replace(IE.Document.getElementbyid("id nome").innertext, "Pag. ", "")
    Fim = Right(Fim, Len(Fim) - InStr(Fim, "/"))
    
    Linha = 2
    Entrou = False
    
    Do While Inicio <= Int(Fim)
        Application.StatusBar = "Lendo registro : " & Linha & " - página " & Inicio & "/" & Fim & " | " & Now()
        DoEvents
        
        Set Tabela = IE.Document.getElementsByTagName("table")
        Set el = Tabela(53).getElementsByTagName("tr")
        
        Linha_Html = 0
        Linha_Tabela = 0
        
        For Each eleRow In el
            Application.StatusBar = "Lendo registro : " & Linha & " - página " & Inicio & "/" & Fim & " | " & Now()
            DoEvents
            
            Set Tabela = IE.Document.getElementsByTagName("table")
            
            'Pega os elementos "TD" contidos no "TR"
            Set eleColtd = Tabela(53).getElementsByTagName("tr")(Linha_Html).getElementsByTagName("td")
            
            Coluna = 1
            Coluna_Html = 1
                
            For Each eleCol In eleColtd
                If Coluna_Html > 1 And Coluna_Html <> 4 And Coluna_Html <> 5 And Coluna_Html < 10 Then 'Ignora as colunas do IF
                    If Coluna = 3 Then 'Data
                        Planilha.Cells(Linha, Coluna) = Format(Left(eleCol.innertext, 10), "mm/dd/yyyy")
                        Coluna = Coluna + 1
                        Planilha.Cells(Linha, Coluna) = Mid(eleCol.innertext, 12, 8)
                        Coluna = Coluna + 1
                    else
                        Planilha.Cells(Linha, Coluna) = eleCol.innertext
                        
                        Coluna = Coluna + 1
                    End If
                    Entrou = True
                End If
                Coluna_Html = Coluna_Html + 1
                DoEvents
            Next eleCol
                                   
            Linha_Html = Linha_Html + 1
            
            If Entrou Then
                Linha = Linha + 1
                Entrou = False
            End If
            DoEvents
        Next eleRow
        
        'NextButton
        If Inicio < Int(Fim) Then
            IE.Document.getElementbyid("Next Botao").Click
            Aguardar IE 'Aguarda tela estar carregada.
        End If
        
        Inicio = Inicio + 1
    Loop

    Planilha.Cells.EntireColumn.AutoFit
    
    Planilha.Cells(Linha, 1).Select
 

        
Trata_Erro:
    Application.StatusBar = ""
    
    IE.Quit
    Set IE = Nothing
End Sub
