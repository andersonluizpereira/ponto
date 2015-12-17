<%
' Declaração de variáveis locais.
Dim vobj_ListaErros
Dim vobj_ListaCamposFocus


' Faz a criação do objeto collection.
Set vobj_ListaErros = Server.CreateObject("Scripting.Dictionary")
Set vobj_ListaCamposFocus = Server.CreateObject("Scripting.Dictionary")


Public Sub AddErro(pstr_NomeChave, pstr_Valor)
	
	' Tratamento de erro. --------------------
	
	' Verifica se o objeto de colleção de erros
	' está criado, se não estiver então finaliza
	' a chamada da procedure.
	If Not(ObjetoExiste()) Then
		Exit Sub
	End If
	
	
	' Adiciona o erro na coleção de erros.
	Call vobj_ListaErros.Add(pstr_NomeChave & "xxx_" & vobj_ListaErros.Count, pstr_Valor)
End Sub


Public Sub AddErroFocus(pstr_NomeChave, pstr_Valor, pstr_CampoFocus)
	
	' Tratamento de erro. --------------------
	
	' Verifica se o objeto de colleção de erros
	' está criado, se não estiver então finaliza
	' a chamada da procedure.
	If Not(ObjetoExiste()) Then
		Exit Sub
	End If
	
	
	' Adiciona o erro na coleção de erros.
	Call vobj_ListaErros.Add(pstr_NomeChave & "xxx_" & vobj_ListaErros.Count, pstr_Valor)
	Call vobj_ListaCamposFocus.Add(pstr_NomeChave & "xxx_" & vobj_ListaErros.Count, pstr_CampoFocus)
End Sub


Private Function ObjetoExiste()
	
	' Retorno para a função true se o objeto existir
	' ou false se não existir.
	ObjetoExiste = Not(vobj_ListaErros Is Nothing)
End Function


Public Function TotalErros()
	
	' Verifica se o objeto de colleção de erros
	' está criado, se não estiver então finaliza
	' a chamada da função.
	If Not(ObjetoExiste()) Then
		
		TotalErros = 0
	Else
		
		TotalErros = vobj_ListaErros.Count
	End If
End Function


Public Function ExibirErros()
	
	' Declaração de variáveis locais.
	Dim ContaErros
	Dim vstr_Retorno
	
	
	vstr_Retorno	= Empty
	ExibirErros		= Empty
	
	
	' Verifica se o objeto de colleção de erros
	' está criado, se não estiver então finaliza
	' a chamada da função.
	If Not(ObjetoExiste()) Then
		Exit Function
	End If
	
	Dim vstr_Mensagem
	
	vstr_Mensagem = vobj_ListaErros.Items
	
	Dim vint_Contador
	
	For vint_Contador = 0 To vobj_ListaErros.Count -1
		
		vstr_Retorno = "<script> alert('" & vstr_Mensagem(0) & "')</script>"
		
	Next
	
	' Chama procedimento que limpa a
	' coleção de erros do sistema.
	Call LimpaCollecaoErros()
	
	
	' Retorno todos os erros para a função.
	ExibirErros = vstr_Retorno
End Function


Private Sub PosicionarFocusCamposComProblemas()
	
	' Declaração de variáveis locais.
	Dim ContaErros
	
	
	' Tratamento de Erro. Verificando se o existe realmente
	' um objeto coleção instanciado na variavel.
	If Not(vobj_ListaCamposFocus Is Nothing) Then
	
		' Loop de todos os campos que devem possuir o
		' focus apontado.
		For Each ContaErros In vobj_ListaCamposFocus.Items
			
			' Posiciona o focus no campo.
			Call SetarFocus(ContaErros)
			Exit For
		Next
		
		
		' Remove todos os itens da coleção.
		vobj_ListaCamposFocus.RemoveAll()
		
		' Limpa o objeto da memória.
		Set vobj_ListaCamposFocus = Nothing
	End If
End Sub


Private Sub LimpaCollecaoErros()
	
	
	' Verifica se o objeto de colleção de erros
	' está criado.
	If ObjetoExiste() Then
		
		' Limpa toda a lista do objeto.
		vobj_ListaErros.RemoveAll()
	End If
	
	
	' Limpa o objeto da memória.
	Set vobj_ListaErros = Nothing 
End Sub
%>