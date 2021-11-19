'Itera entre os itens de um combo box e seleciona o item desejado.

'Inicia objeto com o Internet Explorer.
Dim IE As Object
Set IE = CreateObject("InternetExplorer.Application")

'Acessa URL
IE.Navigate "https://www.google.com/"

'Chama procedure
Seleciona_Item_Combo IE, "ComboBox1", "Conteúdo", "onchange"
'Onde :
'IE = variável contendo o browser
'ComboBox1 = Id do combo que se deseja iterar
'Conteúdo  = Palavra a ser procurada no combo. Pode se procurar pelo Value ou innertext.
'onchange  = Nome do evento a ser executado vinculado ao combo.

'Procedure
Sub Seleciona_Item_Combo(Browse As Object, Combo As String, Conteudo As String, Evento As String)
    Dim obj As Object
    Dim Item As Object
    
    Set obj = Browse.Document.getElementbyid(Combo)

    For Each Item In obj.Options
        'If Item.Value = Conteudo Then
        If Item.innertext = Conteudo Then
            Item.Selected = True
            
            If Evento <> "" Then
                Browse.Document.getElementbyid(Combo).FireEvent (Evento)
                Aguardar Browse 'Aguarda tela estar carregada.
            End If
            Exit For
        End If
    Next Item
End Sub
