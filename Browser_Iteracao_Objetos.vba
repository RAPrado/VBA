'Inicia objeto com o Internet Explorer.
Dim IE As Object
Set IE = CreateObject("InternetExplorer.Application")

'Forma A
Dim obj As Object

For Each obj In IE.Document.All.Item("ComboBox").Options 'Nesse modo não aceita que o nome ComboBox seja passado por parâmetro, tem que estár fixo no código.
    'If obj.Value = "Teste" Then
    If obj.innertext = "Teste" Then
        obj.Selected = True
    End If
Next Item


'Forma B
Dim obj As Object
Dim Item As Object

Set obj = IE.Document.getElementbyid(Combo)  'Aceita que o nome seja passado por parâmetro.

For Each Item In obj.Options
    'If Item.Value = "Teste" Then
    If Item.innertext = "Teste" Then
        Item.Selected = True
    End If
Next Item
