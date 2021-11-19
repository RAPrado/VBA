'Conexão com o Internet Explorer

'Referências :
'https://stackoverflow.com/questions/54976102/the-remote-server-machine-does-not-exist-or-is-unavailable-for-specific-website
'https://www.tomasvasquez.com.br/forum/viewtopic.php?t=2854
'Criação     : Rodrigo Prado
'Colaboraçao : Cayo Gilson M Silva

Dim URL As String
Dim Win As Object

'------------------------------------------------------------------------
'Método 1 de Conexão ao Internet Explorer
'Ativar objeto Microsoft Ojbect Control em References.
Dim IE As InternetExplorer

'Cria nova instância do Internet Explorer
Set IE = New InternetExplorer

'------------------------------------------------------------------------
'Método 2 de Conexão ao Internet Explorer
Dim IE As Object
Set IE = CreateObject("InternetExplorer.Application")

'------------------------------------------------------------------------

'Torna a Página Visível
IE.Visible = True

URL = "https://www.google.com"
IE.Navigate URL

'************************************************************************************************************************************************************
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
'************************************************************************************************************************************************************

'Ao final
IE.Quit 'Fechar o browser.
Set IE = Nothing 'Limpar variável.
