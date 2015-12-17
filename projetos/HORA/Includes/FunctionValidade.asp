<%
' Procedimento desenvolvido para verificar se o valor passado como parametro
' para o campo pstr_Valor é do tipo Moeda. Se caso o campo for numérico, a função irá
' retornar um valor positivo, caso contrário o retorno será negativo. Para os
' casos onde o retorno é negativo, o campo pstr_CausaValidacao será retornado para
' a aplicação chamadora contendo o motivo da não validação do parametro.
Public Function ValidarValorMoeda(pstr_Valor, ByRef pstr_CausaValidacao)
	
	
	' Verificando se o valor passado como parametro 
	' possui conteudo.
	If Len(pstr_Valor) > 0 Then
		
		
		' Verificando se o valor passado como parametro
		' é do tipo numérico.
		If IsNumeric(pstr_Valor) Then
			
			
			' Verificando se o valor numérico está dentro dos padrões
			' aceitaveis de inteiro positivo longo.
			If CLng(pstr_Valor) >= 0 And CLng(pstr_Valor) <= 1000000 Then
				
				' ... o valor do ano está correto.
				ValidarValorMoeda = True
				pstr_CausaValidacao = Empty
			Else
				
				ValidarValorMoeda = False
				pstr_CausaValidacao = "O valor '" & pstr_Valor & "' não esta entre os limites validos."
			End If
		Else
			
			ValidarValorMoeda = False
			pstr_CausaValidacao = "O valor '" & pstr_Valor & "' não é numérico."
		End If
	Else
		
		ValidarValorMoeda = False
		pstr_CausaValidacao = "Sem valor para comparação."
	End If
End Function


' Procedimento desenvolvido para verificar se o valor passado como parametro
' para o campo pstr_Valor é do tipo Inteiro positivo Longo. Se caso o campo for numérico, 
' a função irá retornar um valor positivo, caso contrário o retorno será negativo. Para os
' casos onde o retorno é negativo, o campo pstr_CausaValidacao será retornado para
' a aplicação chamadora contendo o motivo da não validação do parametro.
Public Function ValidarValorInteiroPossitivoLongo(pstr_Valor, ByRef pstr_CausaValidacao)
	
	
	' Verificando se o valor passado como parametro 
	' possui conteudo.
	If Len(pstr_Valor) > 0 Then
		
		
		' Verificando se o valor passado como parametro
		' é do tipo numérico.
		If IsNumeric(pstr_Valor) Then
			
			
			' Verificando se o valor numérico está dentro dos padrões
			' aceitaveis de inteiro positivo longo.
			If CLng(pstr_Valor) >= 0 And CLng(pstr_Valor) <= 999999999 Then
				
				
				' Rotina que verifica se o parametro informado possui
				' ponto flutuante.
				' -----------------------------------------------------
				Dim vint_ContaCaracteres
				Dim vboo_PossuiSomenteNumeros
				
				
				vboo_PossuiSomenteNumeros = True
				
				
				For vint_ContaCaracteres = 1 To Len(pstr_Valor)
					
					
					' Para todos os caractres o procedimento verifica se o valor
					' é numérico, se for encontrado algum valor tipo '.' ou ',' o 
					' procedimento é intenrrompido.
					If Not (InStr("0123456789", Mid(pstr_Valor, vint_ContaCaracteres, 1)) > 0) Then
						
						vboo_PossuiSomenteNumeros = False
						Exit For
					End If
				Next
				' -----------------------------------------------------
				
				
				' Verificando se não foi encontrados valores de ponto
				' flutuante ('.', ',') na valor informado.
				If vboo_PossuiSomenteNumeros Then
					
					' ... o valor do ano está correto.
					ValidarValorInteiroPossitivoLongo = True
					pstr_CausaValidacao = Empty
				Else
					
					
					ValidarValorInteiroPossitivoLongo = False
					pstr_CausaValidacao = "O valor '" & pstr_Valor & "' deve ser inteiro, não podendo conter valor de ponto flutuante."
				End If
			Else
				
				ValidarValorInteiroPossitivoLongo = False
				pstr_CausaValidacao = "O valor '" & pstr_Valor & "' não esta entre os limites validos."
			End If
		Else
			
			ValidarValorInteiroPossitivoLongo = False
			pstr_CausaValidacao = "O valor '" & pstr_Valor & "' não é numérico."
		End If
	Else
		
		ValidarValorInteiroPossitivoLongo = False
		pstr_CausaValidacao = "Sem valor para comparação."
	End If
End Function


' Procedimento desenvolvido para verificar se o valor passado como parametro
' para o campo pstr_Valor é do tipo Ano. Se caso o campo for numérico, a função irá
' retornar um valor positivo, caso contrário o retorno será negativo. Para os
' casos onde o retorno é negativo, o campo pstr_CausaValidacao será retornado para
' a aplicação chamadora contendo o motivo da não validação do parametro.
Public Function ValidarValorAno(pstr_Valor, ByRef pstr_CausaValidacao)
	
	
	' Verificando se o valor passado como parametro 
	' possui conteudo.
	If Len(pstr_Valor) > 0 Then
		
		
		' Verificando se o valor passado como parametro
		' é do tipo numérico.
		If IsNumeric(pstr_Valor) Then
			
			
			' Verificando se o valor numérico está dentro dos padrões
			' aceitaveis de ano.
			If CInt(pstr_Valor) >= 1900 And CInt(pstr_Valor) <= 2100 Then
				
				
				' ... o valor do ano está correto.
				ValidarValorAno = True
				pstr_CausaValidacao = Empty
			Else
				
				ValidarValorAno = False
				pstr_CausaValidacao = "O valor '" & pstr_Valor & "' não esta entre os limites validos."
			End If
		Else
			
			ValidarValorAno = False
			pstr_CausaValidacao = "O valor '" & pstr_Valor & "' não é numérico."
		End If
	Else
		
		ValidarValorAno = False
		pstr_CausaValidacao = "Sem valor para comparação."
	End If
End Function


' Procedimento desenvolvido para verificar se o valor passado como parametro
' para o campo pstr_Valor é um E-MAIL válido. Se caso o campo não for preenchido 
' com um e-mail válido, a função irá retornar um valor positivo, caso contrário 
' o retorno será negativo. 
Public Function ValidarValorEmail(pstr_Valor)
	
	
	' 
	If Not Trim(pstr_Valor) = "" Then
		
		If Instr(1,Trim(pstr_Valor),"@") > 4 And Len(Mid(pstr_Valor,Instr(1,Trim(pstr_Valor),"@")+1)) > 2 Then
			
			ValidarValorEmail = True
		Else
			
			ValidarValorEmail = False
		End If
	Else
		
		ValidarValorEmail = False
	End If
End Function
%>