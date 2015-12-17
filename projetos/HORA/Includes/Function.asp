<%
' Força a declaração de todas as variáveis.
' -------------------------------------------------------------------
Option Explicit


' Retorna o nome do dia da semana.
Public Function DiaDaSemana(pstr_Data)
	
	Select Case Weekday(pstr_Data)
		
		Case 1  ' Domingo
			DiaDaSemana = "Domingo"

		Case 2  ' Segunda
			DiaDaSemana = "Segunda"
			
		Case 3  ' Terça
			DiaDaSemana = "Terça"

		Case 4  ' Quarta
			DiaDaSemana = "Quarta"

		Case 5  ' Quinta
			DiaDaSemana = "Quinta"

		Case 6  ' Sexta
			DiaDaSemana = "Sexta"
			
		Case 7  ' Sábado
			DiaDaSemana = "Sábado"
			
	End Select
	
End Function

' Retorna o nome do dia da semana.
Public Function DescricaoMes(pstr_Mes)
	
	Select Case Cint(pstr_Mes)
		
		Case 1  ' Janeiro
			DescricaoMes = "Janeiro"

		Case 2  ' Fevereiro
			DescricaoMes = "Fevereiro"
			
		Case 3  ' Março
			DescricaoMes = "Março"

		Case 4  ' Abril
			DescricaoMes = "Abril"

		Case 5  ' Maio
			DescricaoMes = "Maio"

		Case 6  ' Junho
			DescricaoMes = "Junho"
			
		Case 7  ' Julho
			DescricaoMes = "Julho"
		
		Case 8  ' Agosto
			DescricaoMes = "Agosto"
			
		Case 9  ' Setembro
			DescricaoMes = "Setembro"
			
		Case 10  ' Outubro
			DescricaoMes = "Outubro"
			
		Case 11  ' Novembro
			DescricaoMes = "Novembro"
			
		Case 12  ' Dezembro
			DescricaoMes = "Dezembro"
			
	End Select
	
End Function

Public Function converterDataParaSQL(Data)
	
	' Declaração de variáveis auxiliares
	' para conseguir a data de retorno.
	Dim vstr_Dia
	Dim vstr_Mes
	Dim vstr_Ano
	
	' Conseguindo os valores que serão concatenados
	' para o retorno da função
	
	'vstr_Dia = Mid(Data, 1, 2)
	'vstr_Mes = Mid(Data, 4, 2)
	'vstr_Ano = Mid(Data, 7, 4)
	If Not Trim(Data) = "" Then
		
		vstr_Dia = Day(Data)
		vstr_Mes = Month(Data)
		vstr_Ano = DatePart("yyyy", Data)
		
		If Len(vstr_Dia) = 1 Then
			
			vstr_Dia = "0" & vstr_Dia
		End if
		
		If Len(vstr_Mes) = 1 Then
			
			vstr_Mes = "0" & vstr_Mes
		End if
		
		' Retornando a data formatada
		' para a função.
		converterDataParaSQL = vstr_Ano & "/" & vstr_Mes & "/" & vstr_Dia
	Else
		
		converterDataParaSQL = ""
	End If
End Function


Public Function converterHoraParaSQL(Hora)
	
	Dim vstr_Hora
	Dim vstr_Minuto
	Dim vint_Cont
	
	vstr_Hora = CStr(Hour(Hora))
	vstr_Minuto = CStr(Minute(Hora))
	
	If Len(vstr_Hora) = 1 Then
		
		vstr_Hora = "0" & vstr_Hora
	End If
	
	If Len(vstr_Minuto) = 1 Then
		
		vstr_Minuto = "0" & vstr_Minuto
	End If
	
	converterHoraParaSQL = vstr_Hora & ":" & vstr_Minuto
	
End Function

Public Function converterDataParaHtml(Data)
	
	' Declaração de variáveis auxiliares
	' para conseguir a data de retorno.
	Dim vstr_Dia
	Dim vstr_Mes
	Dim vstr_Ano
	
	If Not Data = "" Then
		
		' Conseguindo os valores que serão concatenados
		' para o retorno da função
		vstr_Dia = Day(Data)
		vstr_Mes = Month(Data)
		vstr_Ano = DatePart("yyyy", Data)
		
		If Len(vstr_Dia) = 1 Then
			
			vstr_Dia = "0" & vstr_Dia
		End if
		
		If Len(vstr_Mes) = 1 Then
			
			vstr_Mes = "0" & vstr_Mes
		End if
		
		
		' Retornando a data formatada
		' para a função.
		converterDataParaHtml = vstr_Dia & "/" & vstr_Mes & "/" & vstr_Ano
	Else
		
		converterDataParaHtml = ""
	End If
	
End Function

Public Function IIF(pboo_ExpressaoBooleana, pstr_ValorTrue, pstr_ValorFalse)
	
	' Verifica se a expressão booleana passado
	' como parametro para está função é verdadeira.
	If pboo_ExpressaoBooleana Then 
		
		' ... neste caso devolve o valor positivo
		' para a função.
		IIF = pstr_ValorTrue
	Else
		
		' ... neste caso devolve o valor negativo
		' para a função.
		IIF = pstr_ValorFalse
	End If
End Function


' Função desenvolvida para retornar a data atual
' do relógio do computador. A data retornada será
' no formato dd/mm/aaaa. Ex.: 12/06/2008
Public Function GetData()
	
	' Declaração de variáveis locais.
	Dim vstr_Dia
	Dim vstr_Mes
	Dim vstr_Ano
	
	' Conseguindo os valores da data atual.
	vstr_Dia = Day(Date)
	vstr_Mes = Month(Date)
	vstr_Ano = Year(Date)
	
	' Formatando os valores para retorno.
	If Len(vstr_Dia) = 1 Then vstr_Dia = "0" & vstr_Dia
	If Len(vstr_Mes) = 1 Then vstr_Mes = "0" & vstr_Mes
	If Len(vstr_Ano) = 2 Then vstr_Ano = "20" & vstr_Ano
	
	' Devolvendo o valor formatado para
	' a chamada da função.
	GetData = vstr_Dia & "/" & vstr_Mes & "/" & vstr_Ano
End Function


Public Function ValidaLogin(pstr_Login)
	
	' ---------------------------------------------------------------------
	'		VERIFICA SE O LOGIN JÁ EXISTE NO BANCO DE DADOS
	' ---------------------------------------------------------------------
	
	' Variavel que verificará se existe 
	' o login informado ja existe. 
	
	Dim vobj_commandProc
	Dim vobj_rs
	
	Set vobj_commandProc = Server.CreateObject("ADODB.Command")
	Set vobj_rs = Server.CreateObject("ADODB.Recordset")
	Set vobj_commandProc.ActiveConnection = vobj_conexao
	
	vobj_commandProc.CommandType					= adCmdStoredProc
	vobj_commandProc.CommandText					= "spFunction"
	vobj_commandProc.Parameters.Refresh
	
	vobj_commandProc.Parameters("@TIPO_OPER")		= "VALIDAR_LOGIN_TODOS"
	vobj_commandProc.Parameters("@CD_LOGIN")		= pstr_Login
	' ---------------------------------------------------------------------
	
	
	Set vobj_rs = vobj_commandProc.Execute
	
	' Se o login existir = true.
	If Not vobj_rs.EOF Then
		ValidaLogin = True
	Else
		ValidaLogin = False
	End If
	
	' Fechando os objetos
	vobj_rs.Close
	Set vobj_rs = Nothing
	Set vobj_commandProc = Nothing
	
End Function

' Função desenvolvida para retirar caracter inválido;.
' Obs: Como foi desenvolvido para a criação de pastas
' de fotos do sistema de inventário, todo caracter
' de espaço será subustituído por UNDERLINE.

Public Function retiraCaracterInvalido(pstr_Registro)
	
	pstr_Registro = Lcase(pstr_Registro)
	
	pstr_Registro = Replace(pstr_Registro,"â","a" )
	pstr_Registro = Replace(pstr_Registro,"á","a" )
	pstr_Registro = Replace(pstr_Registro,"à","a" )
	pstr_Registro = Replace(pstr_Registro,"ã","a" )
	pstr_Registro = Replace(pstr_Registro,"é","e" )
	pstr_Registro = Replace(pstr_Registro,"è","e" )
	pstr_Registro = Replace(pstr_Registro,"ê","e" )
	pstr_Registro = Replace(pstr_Registro,"í","i" )
	pstr_Registro = Replace(pstr_Registro,"ì","i" )
	pstr_Registro = Replace(pstr_Registro,"î","i" )
	pstr_Registro = Replace(pstr_Registro,"ô","o" )
	pstr_Registro = Replace(pstr_Registro,"ó","o" )
	pstr_Registro = Replace(pstr_Registro,"ò","o" )
	pstr_Registro = Replace(pstr_Registro,"õ","o" )
	pstr_Registro = Replace(pstr_Registro,"ú","u" )
	pstr_Registro = Replace(pstr_Registro,"ù","u" )
	pstr_Registro = Replace(pstr_Registro,"û","u" )
	pstr_Registro = Replace(pstr_Registro,"(","_" )
	pstr_Registro = Replace(pstr_Registro,")","_" )
	pstr_Registro = Replace(pstr_Registro,"´","" )
	pstr_Registro = Replace(pstr_Registro,"`","" )
	pstr_Registro = Replace(pstr_Registro,"~","" )
	pstr_Registro = Replace(pstr_Registro,"^","" )
	pstr_Registro = Replace(pstr_Registro,"*","" )
	pstr_Registro = Replace(pstr_Registro,"%","" )
	pstr_Registro = Replace(pstr_Registro,"$","" )
	pstr_Registro = Replace(pstr_Registro,"!","" )
	pstr_Registro = Replace(pstr_Registro,"@","" )
	pstr_Registro = Replace(pstr_Registro,"#","" )
	pstr_Registro = Replace(pstr_Registro,"¨","" )
	pstr_Registro = Replace(pstr_Registro,"+","" )
	pstr_Registro = Replace(pstr_Registro,"º","" )
	pstr_Registro = Replace(pstr_Registro,";","" )
	pstr_Registro = Replace(pstr_Registro,":","" )
	pstr_Registro = Replace(pstr_Registro,"-","" )
	pstr_Registro = Replace(pstr_Registro," ","_" )

	retiraCaracterInvalido = pstr_Registro

End Function


' Procedimento desenvolvido para enviar e-mail.
Public Sub SendEmail(pstr_From, pstr_To, pstr_Cc, pstr_Bcc, pstr_Subject, pstr_Body)
	
	Dim vobj_Mail
	Set vobj_Mail = server.CreateObject("CDONTS.NewMail")
	
	vobj_Mail.From 		= pstr_From			'***email de quem esta enviando
	vobj_Mail.To 		= pstr_To		'***Para quem Vai o E-mail
	vobj_Mail.Cc 		= pstr_Cc
	vobj_Mail.Bcc 		= pstr_Bcc		
	vobj_Mail.Subject 	= pstr_Subject		'***Asunto do Email
	vobj_Mail.BodyFormat 	= 0
	vobj_Mail.MailFormat 	= 0
	vobj_Mail.Body 		= pstr_Body		'***Corpo do Email
	
	vobj_Mail.Send 
	
	Set vobj_Mail = nothing
End Sub


' ---------------------------------------------------------
Public Function doubleToMoneyDisplay(pdbl_Valor)
    
    Dim vdbl_Decimal
    Dim vlng_Inteiro
    Dim vstr_Retorno
    
    
    If IsNull(pdbl_Valor) Or Trim(pdbl_Valor) = "" Then
		doubleToMoneyDisplay = "0,00"
		Exit Function
    End If
    
    
    vstr_Retorno = Empty
    vdbl_Decimal = (CDbl(pdbl_Valor) - Fix(pdbl_Valor))
     
    vlng_Inteiro = pdbl_Valor - vdbl_Decimal
    vdbl_Decimal = Replace(vdbl_Decimal, ".", ",")
    
    
    If vdbl_Decimal = 0 Then
        vdbl_Decimal = "0,00"
    Else
        If Len(vdbl_Decimal) >= 4 Then
            
            vdbl_Decimal = Round(vdbl_Decimal,2)
			
        Else
            
            vdbl_Decimal = vdbl_Decimal & String(4 - Len(vdbl_Decimal), "0")
			
        End If
    End If
    
    
    Dim vint_AuxContaAdicaoPonto
    Dim vint_ContaCaracteres
    
    
    vint_AuxContaAdicaoPonto = 0
    
    
    For vint_ContaCaracteres = Len(vlng_Inteiro) To 1 Step -1
        
        vint_AuxContaAdicaoPonto = vint_AuxContaAdicaoPonto + 1
        
        If vint_AuxContaAdicaoPonto = 4 Then
            vint_AuxContaAdicaoPonto = 1
            vstr_Retorno = "." & vstr_Retorno
        End If
        
        vstr_Retorno = Mid(vlng_Inteiro, vint_ContaCaracteres, 1) & vstr_Retorno
    Next
    
    vstr_Retorno = "R$ " & Round(vstr_Retorno & Mid(vdbl_Decimal, 2),2)
    
    If Instr(1,vstr_Retorno,",") = 0 Then
		vstr_Retorno = vstr_Retorno & ",00"
    End If
    
    doubleToMoneyDisplay = vstr_Retorno
End Function


Public Function EncriptaString(pValor)
	
	Dim strRet
    Dim i
	
    ' =======================================================
    ' Monta string com a diferença entre o maior possivel 255
    ' e o valor asc do caracter e inverte os caracteres
    ' =======================================================
    strRet = ""
    If IsNull(pValor) Then
		pValor = ""
    End If
    
    If  Not(pValor = "") Then
		
		For i = 1 To Len(pValor)
			strRet = Chr(255 - Asc(Mid(pValor, i, 1))) & strRet
		Next
		
	Else
	
		strRet = "0"
	End if
	
    EncriptaString = strRet
	
End Function

Public Function DesencriptaString(pValor)
	
	DesencriptaString = EncriptaString(pValor)
	
End Function

Public Function ConverterMinutoParaHora(pint_Minuto)
	
	Dim vint_Minuto
	Dim vint_Hora
	
	vint_Hora = (Cint(pint_Minuto) \ 60)
	vint_Minuto = Cint(pint_Minuto) Mod 60
	
	If Len(vint_Hora) = 1 Then
		
		vint_Hora = "0" & Cstr(vint_Hora)
		
	End If
	
	If Len(vint_Minuto) = 1 Then
		
		vint_Minuto = "0" & Cstr(vint_Minuto)
		
	End If
	
	converterMinutoParaHora = vint_Hora & ":" & vint_Minuto

End Function

Public Function GetUltimoDiaMes(Mes,Ano)
  Select Case Mes
    Case 1,3,5,7,8,10,12: GetUltimoDiaMes = 31
    Case 4,6,9,11: GetUltimoDiaMes = 30
    Case Else
      If Ano Mod 4 = 0 And (Ano Mod 100 <> 0 Or Ano Mod 400 = 0) Then
        GetUltimoDiaMes = 29
      Else
        GetUltimoDiaMes = 28
      End If
  End Select
End Function


' Verificar se há novas mensagens para esse usario
Public Function VerificarNovaMensagem(pstr_usuario)
	
	' Declaração de variáveis auxiliares
	' para obter as informações do registro.
	Dim vobj_commandNovaMensagem
	Dim vobj_rsNovaMensagem
	
	
	' ---------------------------------------------------------------------
	' Selecionando os dados do registro.
	' ---------------------------------------------------------------------
	Set vobj_commandNovaMensagem = Server.CreateObject("ADODB.Command")
	Set vobj_commandNovaMensagem.ActiveConnection = vobj_conexao
	
	
	vobj_commandNovaMensagem.CommandType					= adCmdStoredProc
	vobj_commandNovaMensagem.CommandText					= "consultaNovaMensagem"
	
	
	vobj_commandNovaMensagem.Parameters.Append vobj_commandNovaMensagem.CreateParameter("param1",adChar, adParamInput, 10, pstr_Usuario)
	
	
	' Cria o recordset e posiciona a páginação do recordset.
	Set vobj_rsNovaMensagem = vobj_commandNovaMensagem.Execute
	
	' Verifica se registros foram encontrados.
	If Not vobj_rsNovaMensagem.EOF Then
		
		VerificarNovaMensagem = True
	
	Else
		
		VerificarNovaMensagem = False
		
	End If
	
	vobj_rsNovaMensagem.Close
	Set vobj_rsNovaMensagem = Nothing
	Set vobj_commandNovaMensagem = Nothing
	
End Function

%>
