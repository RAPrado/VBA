'Navega entre as janelas abertas do Windows.
'Nesse caso procura pelo browser do IE e refaz a conexão da variável com a janela, que foi perdida por devido alguma regra de segurança.

'Inicia objeto com o Internet Explorer.
Dim IE As Object
Set IE = CreateObject("InternetExplorer.Application")

Dim Win As Object

For Each Win In CreateObject("Shell.Application").Windows
    If Win.Name Like "*Internet Explorer" Then
        Set IE = Win

        If IE.LocationURL = URL Then
            Exit For
        End If
    End If
Next
Set Win = Nothing
