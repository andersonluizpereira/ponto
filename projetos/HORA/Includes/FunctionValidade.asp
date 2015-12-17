<%
' Procedimento desenvolvido para verificar se o valor passado como parametro
' para o campo pstr_Valor � do tipo Moeda. Se caso o campo for num�rico, a fun��o ir�
' retornar um valor positivo, caso contr�rio o retorno ser� negativo. Para os
' casos onde o retorno � negativo, o campo pstr_CausaValidacao ser� retornado para
' a aplica��o chamadora contendo o motivo da n�o valida��o do parametro.
Public Function ValidarValorMoeda(pstr_Valor, ByRef pstr_CausaValidacao)
	
	
	' Verificando se o valor passado como parametro 
	' possui conteudo.
	If Len(pstr_Valor) > 0 Then
		
		
		' Verificando se o valor passado como parametro
		' � do tipo num�rico.
		If IsNumeric(pstr_Valor) Then
			
			
			' Verificando se o valor num�rico est� dentro dos padr�es
			' aceitaveis de inteiro positivo longo.
			If CLng(pstr_Valor) >= 0 And CLng(pstr_Valor) <= 1000000 Then
				
				' ... o valor do ano est� correto.
				ValidarValorMoeda = True
				pstr_CausaValidacao = Empty
			Else
				
				ValidarValorMoeda = False
				pstr_CausaValidacao = "O valor '" & pstr_Valor & "' n�o esta entre os limites validos."
			End If
		Else
			
			ValidarValorMoeda = False
			pstr_CausaValidacao = "O valor '" & pstr_Valor & "' n�o � num�rico."
		End If
	Else
		
		ValidarValorMoeda = False
		pstr_CausaValidacao = "Sem valor para compara��o."
	End If
End Function


' Procedimento desenvolvido para verificar se o valor passado como parametro
' para o campo pstr_Valor � do tipo Inteiro positivo Longo. Se caso o campo for num�rico, 
' a fun��o ir� retornar um valor positivo, caso contr�rio o retorno ser� negativo. Para os
' casos onde o retorno � negativo, o campo pstr_CausaValidacao ser� retornado para
' a aplica��o chamadora contendo o motivo da n�o valida��o do parametro.
Public Function ValidarValorInteiroPossitivoLongo(pstr_Valor, ByRef pstr_CausaValidacao)
	
	
	' Verificando se o valor passado como parametro 
	' possui conteudo.
	If Len(pstr_Valor) > 0 Then
		
		
		' Verificando se o valor passado como parametro
		' � do tipo num�rico.
		If IsNumeric(pstr_Valor) Then
			
			
			' Verificando se o valor num�rico est� dentro dos padr�es
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
					' � num�rico, se for encontrado algum valor tipo '.' ou ',' o 
					' procedimento � intenrrompido.
					If Not (InStr("0123456789", Mid(pstr_Valor, vint_ContaCaracteres, 1)) > 0) Then
						
						vboo_PossuiSomenteNumeros = False
						Exit For
					End If
				Next
				' -----------------------------------------------------
				
				
				' Verificando se n�o foi encontrados valores de ponto
				' flutuante ('.', ',') na valor informado.
				If vboo_PossuiSomenteNumeros Then
					
					' ... o valor do ano est� correto.
					ValidarValorInteiroPossitivoLongo = True
					pstr_CausaValidacao = Empty
				Else
					
					
					ValidarValorInteiroPossitivoLongo = False
					pstr_CausaValidacao = "O valor '" & pstr_Valor & "' deve ser inteiro, n�o podendo conter valor de ponto flutuante."
				End If
			Else
				
				ValidarValorInteiroPossitivoLongo = False
				pstr_CausaValidacao = "O valor '" & pstr_Valor & "' n�o esta entre os limites validos."
			End If
		Else
			
			ValidarValorInteiroPossitivoLongo = False
			pstr_CausaValidacao = "O valor '" & pstr_Valor & "' n�o � num�rico."
		End If
	Else
		
		ValidarValorInteiroPossitivoLongo = False
		pstr_CausaValidacao = "Sem valor para compara��o."
	End If
End Function


' Procedimento desenvolvido para verificar se o valor passado como parametro
' para o campo pstr_Valor � do tipo Ano. Se caso o campo for num�rico, a fun��o ir�
' retornar um valor positivo, caso contr�rio o retorno ser� negativo. Para os
' casos onde o retorno � negativo, o campo pstr_CausaValidacao ser� retornado para
' a aplica��o chamadora contendo o motivo da n�o valida��o do parametro.
Public Function ValidarValorAno(pstr_Valor, ByRef pstr_CausaValidacao)
	
	
	' Verificando se o valor passado como parametro 
	' possui conteudo.
	If Len(pstr_Valor) > 0 Then
		
		
		' Verificando se o valor passado como parametro
		' � do tipo num�rico.
		If IsNumeric(pstr_Valor) Then
			
			
			' Verificando se o valor num�rico est� dentro dos padr�es
			' aceitaveis de ano.
			If CInt(pstr_Valor) >= 1900 And CInt(pstr_Valor) <= 2100 Then
				
				
				' ... o valor do ano est� correto.
				ValidarValorAno = True
				pstr_CausaValidacao = Empty
			Else
				
				ValidarValorAno = False
				pstr_CausaValidacao = "O valor '" & pstr_Valor & "' n�o esta entre os limites validos."
			End If
		Else
			
			ValidarValorAno = False
			pstr_CausaValidacao = "O valor '" & pstr_Valor & "' n�o � num�rico."
		End If
	Else
		
		ValidarValorAno = False
		pstr_CausaValidacao = "Sem valor para compara��o."
	End If
End Function


' Procedimento desenvolvido para verificar se o valor passado como parametro
' para o campo pstr_Valor � um E-MAIL v�lido. Se caso o campo n�o for preenchido 
' com um e-mail v�lido, a fun��o ir� retornar um valor positivo, caso contr�rio 
' o retorno ser� negativo. 
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