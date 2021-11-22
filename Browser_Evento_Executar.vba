IE.Document.getElementbyid("Combo").FireEvent ("OnChange") 'Executa o evento de um objeto.
'ou 
IE.Document.getElementbyid("Button").Click  'Executa o evento Click de um botão.
'ou
IE.Document.parentWindow.execScript "EventoA(param1,'parm2',true);document.getElementById('Botao').click();"
