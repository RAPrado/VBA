'Inicia objeto com o Internet Explorer.
Dim IE As Object
Set IE = CreateObject("InternetExplorer.Application")

'Acessa URL
IE.Navigate "https://www.google.com/"
  
'Aguarda enquanto o browser está processando.
'Linha de comando
Do Until Not IE.Busy And IE.ReadyState = 4: DoEvents: Loop

'Ou chamada por procedure
Sub Aguardar(Browse As Object)
    Do Until Not Browse.Busy And Browse.ReadyState = 4: DoEvents: Loop
End Sub
