'Retorna o conteúdo de um ComboBox (Select) selecionado

'Inicia objeto com o Internet Explorer.
Dim IE As Object
Dim Texto as String
Dim Codigo as String
Set IE = CreateObject("InternetExplorer.Application")

'Acessa URL
IE.Navigate "https://www.google.com/"

Codigo = IE.Document.getElementbyid("Nome_Combo").Value
Texto  = IE.Document.getElementbyid("Nome_Combo").Options(IE.Document.getElementbyid("Nome_Combo").SelectedIndex).Text
