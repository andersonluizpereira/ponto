<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- #include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->

<%

If	Not Session("sboo_fladministrador") = True Then
	
	Response.Redirect getBaseLink("/horas/horaslancamento.asp")
	
End If


' Declaração de variáveis locais. ==============================================

' Guarda a operação que será executa nesta tela.
' Obs.: Seus valores podem ser A = Alteração, I = Inclusão, V = Visualização.
Dim vstr_Operacao

' Variável flag que indica se a página deve ser 
' processada, apenas disponivel para as operações de
' A e I.
Dim vstr_Processar

'Variável de controle e fluxo de acoes
Dim vstr_Executar

' Armazena o código de referência do registro que será alterado, Inclusso
' ou visualizado.
Dim vstr_IdUsuario

' Declaração de variáveis utilizadas para armazenar os
' valores dos campos da tela.
Dim vstr_CdUsuario
Dim vstr_DsUsuario
Dim vstr_DsPer
Dim vstr_DsHorasen
Dim vstr_DsCPF
Dim vstr_DsRG
Dim vint_IdFuncao
Dim vint_IdEquipe
Dim vint_FlPerfil
Dim vstr_DtNascimento
Dim vstr_DsTelefone
Dim vstr_DsRamal
Dim vstr_DsLocalAlocado


Dim vstr_DtAniversario

'Dim vint_FlAtivo
Dim vstr_CdSenha
Dim vstr_CdConfirmaSenha





' para está página.
vstr_Operacao		= Request.Form("pstr_Operacao")
vstr_Processar		= Request.Form("hdnProcessar")
vstr_Executar		= Request.Form("hdnExecutar")

' Verifica se o parametro que defini o tipo
' de operação a ser executado na página é
' igual a branco(vazio).
If Trim(vstr_Operacao) = "" Then
	
	' ... neste caso a operação
	' padrão de a de visualização apenas
	' do registro.
	vstr_Operacao = "V"
End If

' Analizando a variável que indica o fluxo de 
' operação desta página.
If vstr_Executar = "DESATIVAR" Then
	
	
	' -->> Operação de Exclusão de Registros.
		
		
	' ... operação de exclusão de registros.
		
		
	' Declaração de variáveis auxiliares que
	' auxiliarão na exclusão dos registros selecionados.
	Dim vstr_DesativarIdRegistro
	Dim vobj_commandDesativar
		
		
		
		
	' Conseguindo todos os registros selecionados para
	' a exclusão do banco de dados.
	vstr_DesativarIdRegistro = Request.Form("hdnIdRegistro")
		
		
	' ---------------------------------------------------------------------
	' Exclusão de Registros do banco de dados.
	' ---------------------------------------------------------------------
	Set vobj_commandDesativar = Server.CreateObject("ADODB.Command")
	Set vobj_commandDesativar.ActiveConnection = vobj_conexao
			
			
	vobj_commandDesativar.CommandType					= adCmdStoredProc
	vobj_commandDesativar.CommandText					= "excluiUsuario"
	vobj_commandDesativar.CommandTimeout					= 0	
		
			
			
	' Iguinorando os erros que ocorrem na exclusão
	' do registro do banco de dados.
	On Error Resume Next
		
	' Passando o código do Registro a ser
	' excluido do banco de dados.
	vobj_commandDesativar.Parameters.Append vobj_commandDesativar.CreateParameter("param1",adChar, adParamInput, 10, vstr_DesativarIdRegistro)
		
		
	' Chamando comando para excluir o registro
	Call vobj_commandDesativar.Execute
			
	' Analizando os erros que podem ter ocorrido
	' na exclusão do registros selecionados pelo
	' usuário.
	Select Case Err.number 
				
		' Verificando se o erro de integridade referencial
		' ocorreu na exclusão do registro acima.
		Case -2147217900
					
			%><script>alert("Atenção!!!\n\nO(s) registro(s) que não foi(ram) excluido(s) possui(em) dados relacionados. Exclua os dados relacionados para poder excluir este(s) registro(s).");</script><%
					
	End Select
			
			
			
	' Habilitando a mensagem de erro quando um
	' erro acontecer.
	On Error Goto 0
			
			
			
	' Limpa a variável utilizada para excluir os
	' registros do banco de dados.
	Set vobj_commandDesativar = Nothing
		
	Response.Redirect("usuarioslistagem.asp")
	
End If


' Analizando a variável que indica o fluxo de 
' operação desta página.
If vstr_Executar = "ATIVAR" Then
	
	
	' -->> Operação de Exclusão de Registros.
		
		
	' ... operação de exclusão de registros.
		
		
	' Declaração de variáveis auxiliares que
	' auxiliarão na exclusão dos registros selecionados.
	Dim vstr_AtivarIdRegistro
	Dim vobj_commandAtivar
		
		
		
		
	' Conseguindo todos os registros selecionados para
	' a exclusão do banco de dados.
	vstr_AtivarIdRegistro = Request.Form("hdnIdRegistro")
		
		
	' ---------------------------------------------------------------------
	' Exclusão de Registros do banco de dados.
	' ---------------------------------------------------------------------
	Set vobj_commandAtivar = Server.CreateObject("ADODB.Command")
	Set vobj_commandAtivar.ActiveConnection = vobj_conexao
			
			
	vobj_commandAtivar.CommandType					= adCmdStoredProc
	vobj_commandAtivar.CommandText					= "ativarUsuario"
	vobj_commandAtivar.CommandTimeout					= 0	
		
			
			
	' Iguinorando os erros que ocorrem na exclusão
	' do registro do banco de dados.
	On Error Resume Next
		
	' Passando o código do Registro a ser
	' excluido do banco de dados.
	vobj_commandAtivar.Parameters.Append vobj_commandAtivar.CreateParameter("param1",adChar, adParamInput, 10, vstr_AtivarIdRegistro)
	
	
	' Chamando comando para excluir o registro
	Call vobj_commandAtivar.Execute
			
	' Analizando os erros que podem ter ocorrido
	' na exclusão do registros selecionados pelo
	' usuário.
	Select Case Err.number 
				
		' Verificando se o erro de integridade referencial
		' ocorreu na exclusão do registro acima.
		Case -2147217900
					
			%><script>alert("Atenção!!!\n\nO(s) registro(s) que não foi(ram) excluido(s) possui(em) dados relacionados. Exclua os dados relacionados para poder excluir este(s) registro(s).");</script><%
					
	End Select
			
			
			
	' Habilitando a mensagem de erro quando um
	' erro acontecer.
	On Error Goto 0
			
			
			
	' Limpa a variável utilizada para excluir os
	' registros do banco de dados.
	Set vobj_commandAtivar = Nothing
		
	Response.Redirect("usuarioslistagem.asp")
	
End If


' *******************************************************
' INICIO DA ROTINA QUE CONSEGUE OS DADOS DO REGISTRO
' *******************************************************

' Veririfa se a operação a ser executada nesta página é a 
' operação de Alteração ou Visualização e se a página não
' foi processada ainda.
If (vstr_Operacao = "A" or vstr_Operacao = "V") And vstr_Processar <> "S" Then
	
	
	' ... neste caso deve ser solicitado o código do registro
	' e encontrar suas informações no banco de dados para exibir para
	' as informações do registro na tela.
	' Conseguindo o código do registro.
	vstr_IdUsuario				= Request.Form("hdnIdRegistro")
	
	
	' Declaração de variáveis auxiliares
	' para obter as informações do registro.
	Dim vobj_rsRegistro
	Dim vobj_commandRegistro
	
	
	
	' ---------------------------------------------------------------------
	' Selecionando os dados do registro.
	' ---------------------------------------------------------------------
	Set vobj_commandRegistro = Server.CreateObject("ADODB.Command")
	Set vobj_commandRegistro.ActiveConnection = vobj_conexao
	
	
    vobj_commandRegistro.CommandType					= adCmdStoredProc
  	vobj_commandRegistro.CommandText					= "consultaUsuario"
	
	vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("param1", adChar, adParamInput, 10, vstr_IdUsuario)
	' ---------------------------------------------------------------------
	
	
	' Cria o objeto recordset com as informações do registro.	
	Set vobj_rsRegistro = vobj_commandRegistro.Execute
	
	
	If Not vobj_rsRegistro.EOF Then
		
		' Conseguindo os dados do registro.
		vstr_CdUsuario			= vobj_rsRegistro("ID_USUARIO")
        vstr_DsUsuario			= vobj_rsRegistro("DS_USUARIO")
		vstr_DsPer              = vobj_rsRegistro("DS_PER")
		vstr_CdSenha			= vobj_rsRegistro("CD_SENHA")
		vstr_DsHorasen     		= vobj_rsRegistro("DS_HORASEN")
		vstr_DsCPF				= vobj_rsRegistro("DS_CPF")
		vstr_DsRG				= vobj_rsRegistro("DS_RG")
		vint_IdFuncao			= vobj_rsRegistro("ID_FUNCAO")
		vint_IdEquipe			= vobj_rsRegistro("ID_EQUIPE")
		vint_FlPerfil			= vobj_rsRegistro("FL_PERFIL")
		'vint_FlAtivo			= vobj_rsRegistro("FL_ATIVO")
		
		vstr_DtNascimento		= converterDataParaHtml(vobj_rsRegistro("DS_NASCIMENTO"))
		vstr_DsTelefone			= vobj_rsRegistro("DS_TELEFONE")
		vstr_DsRamal			= vobj_rsRegistro("DS_RAMAL")
		vstr_DsLocalAlocado		= vobj_rsRegistro("DS_LOCAL_ALOCADO")
		
        
		
	End If
	
	vobj_rsRegistro.Close
	Set vobj_rsRegistro = Nothing
	Set vobj_commandRegistro = Nothing
Else
	
	
	' Verifica se a operação a ser executada nesta página é
	' a operação de inclusão e verifica se a página não foi
	' processada ainda.
	If vstr_Operacao = "I" And vstr_Processar <> "S" Then
		
		' Neste caso todas as variáveis devem ser vazias
		' para o usuário poder preencher seu novo cadastro
		' do registro.
		
        vstr_IdUsuario			= Empty
		vstr_CdUsuario			= Empty
		vstr_DsUsuario			= Empty
        vstr_DsPer              = Empty
		vstr_DsHorasen          = Empty
		vstr_DsCPF				= Empty
		vstr_DsRG				= Empty
		vint_IdFuncao			= Empty
		vint_IdEquipe			= Empty
		vint_FlPerfil			= Empty
		'vint_FlAtivo			= Empty
		vstr_CdSenha			= Empty
		vstr_CdConfirmaSenha	= Empty
		
		vstr_DtNascimento		= Empty
		vstr_DsTelefone			= Empty
		vstr_DsRamal			= Empty
		vstr_DsLocalAlocado		= Empty
		
		
	Else
		
		' ... está opção acontecerá quando o usuário processar
		' a página, neste caso todas os dados da tela serão
		' submetidos e devem ser pegos neste lugar.
		
		vstr_CdUsuario			= Request.Form("txtCdUsuario")
		vstr_DsUsuario			= Request.Form("txtDsUsuario")
		vstr_DsPer			    = Request.Form("txtDsPer")
		vstr_DsHorasen		    = Request.Form("txtDshorasen")
		vstr_DsCPF				= Request.Form("txtDsCPF")
		vstr_DsRG				= Request.Form("txtDsRG")
		vint_IdFuncao			= Request.Form("cmbComboFuncao")
		vint_IdEquipe			= Request.Form("txtDsEquipe")
		vint_FlPerfil			= Request.Form("cmbComboPerfil")
		'vint_FlAtivo			= Request.Form("txtFlAtivo")
		vstr_CdSenha			= Request.Form("txtCdSenha")
		vstr_CdConfirmaSenha	= Request.Form("txtCdConfirmaSenha")
		
		vstr_DtNascimento		= Request.Form("txtDtNascimento")
		vstr_DsTelefone			= Request.Form("txtDsTelefone")
		vstr_DsRamal			= Request.Form("txtDsRamal")
		vstr_DsLocalAlocado		= Request.Form("txtDsLocalAlocado")
		
		
				
	End If
End If
' *******************************************************
' FINAL DA ROTINA QUE CONSEGUE OS DADOS DO REGISTRO
' *******************************************************



' *******************************************************
' INICIO DA ROTINA QUE FAZ O PROCESSAMENTO DOS DADOS
' DO REGISTRO.
' *******************************************************

' Verifica se a variável flag está setada como S, 
' isto indica que um processamento deve ser feito.
If vstr_Processar = "S" Then
	
	
	' Declaração de variáveis auxiliares
	' para fazer o processamento da página.
	Dim vobj_commandProc
	
	
	' Analiza a operação a ser executada na página
	' para descobrir o processamento que deve ser feito.
	Select Case vstr_Operacao
		
		
		Case "A"						' Operação de alteração do registro.
			
			' ... processamento de alteração do registro.
			
			vstr_IdUsuario				= Request.Form("hdnIdRegistro")
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
				
				
				' ---------------------------------------------------------------------
				' Procedimento desenvolvimento para tratar a entrada de umas mesma
				' area
				' ---------------------------------------------------------------------
				
				' Declaração de variáveis auxiliares
				' para obter as informações do registro.
				Dim vobj_rsRegistroConsultaAltera
				Dim vobj_commandRegistroConsultaAltera
				
				If Not Trim(vstr_CdUsuario) = Trim(vstr_IdUsuario) Then
					
					' ---------------------------------------------------------------------
					' Selecionando os dados do registro.
					' ---------------------------------------------------------------------
					Set vobj_commandRegistroConsultaAltera = Server.CreateObject("ADODB.Command")
					Set vobj_commandRegistroConsultaAltera.ActiveConnection = vobj_conexao
				
				
					vobj_commandRegistroConsultaAltera.CommandType					= adCmdStoredProc
					vobj_commandRegistroConsultaAltera.CommandText					= "consultaUsuario"
					
					
					vobj_commandRegistroConsultaAltera.Parameters.Append vobj_commandRegistroConsultaAltera.CreateParameter("param1",adChar, adParamInput, 10, vstr_CdUsuario)
					' ---------------------------------------------------------------------
					
					
					' Cria o objeto recordset com as informações do registro.	
					Set vobj_rsRegistroConsultaAltera = vobj_commandRegistroConsultaAltera.Execute
					
					'Verificando se ja ha registro no banco com mesma area
					'Obs. É verificado soment campo Area, campo nome pode haver dois iguais.
					If Not vobj_rsRegistroConsultaAltera.EOF Then
						
						Call AddErro("Erro", "Há um registro com o mesmo nome de usuário alterado.")
						
					Else
						
						' ---------------------------------------------------------------------
						' Alterando os dados do registro no banco de dados.
						' ---------------------------------------------------------------------
						Set vobj_commandProc = Server.CreateObject("ADODB.Command")
						Set vobj_commandProc.ActiveConnection = vobj_conexao
						
						vobj_commandProc.CommandType					= adCmdStoredProc
						vobj_commandProc.CommandText					= "alteraUsuario"
						
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_IdUsuario))
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_CdUsuario))
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 100, Trim(vstr_DsUsuario))
                        vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adChar, adParamInput, 11, Trim(vstr_DsHorasen))
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adChar, adParamInput, 11, Trim(vstr_DsCPF))
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adChar, adParamInput, 15, Trim(vstr_DsRG))
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adInteger, adParamInput,, vint_IdFuncao)
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param7",adChar, adParamInput, 25, vint_IdEquipe)
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param8",adInteger, adParamInput,, vint_FlPerfil)
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param9",adChar, adParamInput, 15, EncriptaString(Trim(vstr_CdSenha)))
						
						
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param10",adChar, adParamInput, 10, converterDataParaSQL(vstr_DtNascimento))
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param11",adChar, adParamInput, 20, Trim(vstr_DsTelefone))
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param12",adChar, adParamInput, 15, Trim(vstr_DsRamal))
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param13",adChar, adParamInput, 30, Trim(vstr_DsLocalAlocado))
					
						
						If Not Trim(vstr_DtNascimento) = "" Then
							
							vstr_DtAniversario = DateSerial(2000, Month(vstr_DtNascimento), Day(vstr_DtNascimento))
							
						End If
						
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param12",adChar, adParamInput, 10, converterDataParaSQL(vstr_DtAniversario))
						
						Call vobj_commandProc.Execute
						
						Set vobj_commandProc = Nothing
						' ---------------------------------------------------------------------
						
						
						If Trim(Session("sstr_IdUsuario")) = Trim(vstr_IdUsuario) Then
							
							Session("sstr_IdUsuario")		= Trim(vstr_CdUsuario)
							Session("sstr_DsUsuario")		= Trim(vstr_DsUsuario)
							
						End If
						
						
						' Redireciona para a página de listagem
						' dos registros.
						Response.Redirect("usuarioslistagem.asp")
					End If
					
				Else
					
					' ---------------------------------------------------------------------
					' Alterando os dados do registro no banco de dados.
					' ---------------------------------------------------------------------
					Set vobj_commandProc = Server.CreateObject("ADODB.Command")
					Set vobj_commandProc.ActiveConnection = vobj_conexao
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "alteraUsuario"
					
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_IdUsuario))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 10, Trim(vstr_CdUsuario))
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adChar, adParamInput, 100, Trim(vstr_DsUsuario))
                    vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adChar, adParamInput, 11, Trim(vstr_DsPer))
                    vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adChar, adParamInput, 11, Trim(vstr_DsHorasen))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adChar, adParamInput, 11, Trim(vstr_DsCPF))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param7",adChar, adParamInput, 15, Trim(vstr_DsRG))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param8",adInteger, adParamInput,, vint_IdFuncao)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param9",adChar, adParamInput, 25, vint_IdEquipe)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param10",adInteger, adParamInput,, vint_FlPerfil)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param11",adChar, adParamInput, 15, EncriptaString(Trim(vstr_CdSenha)))
					
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param12",adChar, adParamInput, 10, Trim(vstr_DtNascimento))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param13",adChar, adParamInput, 20, Trim(vstr_DsTelefone))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param14",adChar, adParamInput, 15, Trim(vstr_DsRamal))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param15",adChar, adParamInput, 30, Trim(vstr_DsLocalAlocado))
					
					
					If Not Trim(vstr_DtNascimento) = "" Then
						
						vstr_DtAniversario = DateSerial(2000, Month(vstr_DtNascimento), Day(vstr_DtNascimento))
							
					End If
						
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param12",adChar, adParamInput, 10, converterDataParaSQL(vstr_DtAniversario))
					
					Call vobj_commandProc.Execute
					
					Set vobj_commandProc = Nothing
					' ---------------------------------------------------------------------
					
					
					If Trim(Session("sstr_IdUsuario")) = Trim(vstr_IdUsuario) Then
						
						Session("sstr_IdUsuario")		= Trim(vstr_CdUsuario)
						Session("sstr_DsUsuario")		= Trim(vstr_DsUsuario)
						
					End If
					
					' Redireciona para a página de listagem
					' dos registros.
					Response.Redirect("usuarioslistagem.asp")
				End If
				
				vobj_rsRegistroConsultaAltera.Close
				Set vobj_rsRegistroConsultaAltera = Nothing
				Set vobj_commandRegistroConsultaAltera = Nothing
				
			End If
			
			
		Case "I"						' Operação de inclusão do registro.
			
			
			' ... processamento de inclusão de registro.
			
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
				
				' ---------------------------------------------------------------------
				' Procedimento desenvolvimento para tratar a entrada de umas mesma
				' area
				' ---------------------------------------------------------------------
				
				' Declaração de variáveis auxiliares
				' para obter as informações do registro.
				Dim vobj_rsRegistroConsulta
				Dim vobj_commandRegistroConsulta
				
				
				' ---------------------------------------------------------------------
				' Selecionando os dados do registro.
				' ---------------------------------------------------------------------
				Set vobj_commandRegistroConsulta = Server.CreateObject("ADODB.Command")
				Set vobj_commandRegistroConsulta.ActiveConnection = vobj_conexao
				
				
				vobj_commandRegistroConsulta.CommandType					= adCmdStoredProc
				vobj_commandRegistroConsulta.CommandText					= "consultaUsuario"
				
				
				vobj_commandRegistroConsulta.Parameters.Append vobj_commandRegistroConsulta.CreateParameter("param1",adChar, adParamInput, 10, vstr_CdUsuario)
				' ---------------------------------------------------------------------
				
				
				' Cria o objeto recordset com as informações do registro.	
				Set vobj_rsRegistroConsulta = vobj_commandRegistroConsulta.Execute
				
				'Verificando se ja ha registro no banco com mesma area
				'Obs. É verificado soment campo Area, campo nome pode haver dois iguais.
				If Not vobj_rsRegistroConsulta.EOF Then
					
					Call AddErro("Erro", "Há um registro com o mesmo nome de Usuário.")
					
				Else
					
					Dim vobj_rs
					
					' ---------------------------------------------------------------------
					' Incluindo os dados do registro no banco de dados.
					' ---------------------------------------------------------------------
					Set vobj_commandProc = Server.CreateObject("ADODB.Command")
					Set vobj_commandProc.ActiveConnection = vobj_conexao
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "incluiUsuario"
					
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_CdUsuario))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 100, Trim(vstr_DsUsuario))
                    vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adChar, adParamInput, 11, Trim(vstr_DsPer))
                    vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adChar, adParamInput, 11, Trim(vstr_DsHorasen))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adChar, adParamInput, 11, Trim(vstr_DsCPF))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adChar, adParamInput, 15, Trim(vstr_DsRG))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param7",adInteger, adParamInput,, vint_IdFuncao)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param8",adChar, adParamInput, 25, vint_IdEquipe)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param9",adInteger, adParamInput,, vint_FlPerfil)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param10",adChar, adParamInput, 15, EncriptaString(Trim(vstr_CdSenha)))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param11",adDate, adParamInput,, converterDataParaSQL(Date()))
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param12",adChar, adParamInput, 10, Trim(vstr_DtNascimento))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param13",adChar, adParamInput, 20, Trim(vstr_DsTelefone))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param14",adChar, adParamInput, 15, Trim(vstr_DsRamal))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param15",adChar, adParamInput, 30, Trim(vstr_DsLocalAlocado))
					
					
					
					
					If Not Trim(vstr_DtNascimento) = "" Then
						
						vstr_DtAniversario = DateSerial(2000, Month(vstr_DtNascimento), Day(vstr_DtNascimento))
						
					End If
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param13",adChar, adParamInput, 10, converterDataParaSQL(vstr_DtAniversario))
					
					vobj_commandProc.Execute
					
					
					Set vobj_commandProc = Nothing
					
					
					' Altera a variável que indica o tipo de
					' operação que é executada na página.
					vstr_Operacao = "A"
					
					
					' Redireciona para a página de listagem
					' dos registros.
					Response.Redirect("usuarioslistagem.asp")
					
				End If
				
				vobj_rsRegistroConsulta.Close
				Set vobj_rsRegistroConsulta = Nothing
				Set vobj_commandRegistroConsulta = Nothing
				
				
				
			End If
	End Select
	
End If
' *******************************************************
' FINAL DA ROTINA QUE FAZ O PROCESSAMENTO DOS DADOS
' DO REGISTRO.
' *******************************************************
%>

<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/usuariosmanutencao.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<form name="thisForm" action="usuariosmanutencao.asp" method="post">
				
				<input type="hidden" name="hdnProcessar" value="S">
				<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>">
				<input type="hidden" name="hdnIdRegistro" value="<%=vstr_IdUsuario%>">
				
				<i><b class="TituloPagina">Usuários</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td><%=ExibirErros()%></td>
					</tr>
					<tr>
						<td colspan="2">
						<fieldset style="LEFT: 0px; WIDTH: 595px; HEIGHT: 150px">
							<legend>
							   <b>Dados do Usuário</b>
							</legend>
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabPesquisa" id="tabPesquisa" width="100%" style="FILTER: alpha(opacity  =80)">
								<tr>
									<td align="left">Usuário:&nbsp;</td>
									<td align="left" colspan="3"><input name="txtCdUsuario" id="User_ID" class="TextBox" size="15" maxlength="10" value="<%=vstr_CdUsuario%>"></td>
									<td align="left">Nome:&nbsp;</td>
									<td colspan="3" align="left"><input name="txtDsUsuario" id="Nome" class="TextBox" size="35" maxlength="100" value="<%=vstr_DsUsuario%>"></td>
									<tr>
									<td align="left">Entrada Ponto:&nbsp;</td>
									<td colspan="3" align="left"><input name="txtDshorasen" id="txtDshorasen" class="TextBox" size="6" maxlength="6" value="<%=vstr_DsHorasen%>"></td>
									
									<td align="left">Periodo Trab:&nbsp;</td>
									<td colspan="3" align="left"><input name="txtDsPer" id="Text1" class="TextBox" size="6" maxlength="6" value="<%=vstr_DsPer%>"></td>
									
								</tr>
								<tr>
									<td align="left">Data Nascimento:&nbsp;</td>
									<td align="left">
										<input type="text" name="txtDtNascimento" class="TextBox" size="15" maxlength="10" value="<%=vstr_DtNascimento%>">
									</td>
								</tr>
								<tr>
									<td align="left">CPF:&nbsp;</td>
									<td align="left" colspan="3"><input name="txtDsCPF" id="CPF" class="TextBox" size="15" maxlength="11" value="<%=vstr_DsCPF%>"></td>
									<td align="left">RG:&nbsp;</td>
									<td colspan="3" align="left"><input name="txtDsRG" id="RG" class="TextBox" maxlength="15" value="<%=vstr_DsRG%>"></td>
								</tr>
								<tr>
								</tr>
								<tr>
									<td align="left">Função:&nbsp;</td>
									
									<td align="left" colspan="3"><%Call CriarComboFuncao("cmbComboFuncao", vint_IdFuncao,Empty,Empty) %></td>
									<td align="left">Perfil:&nbsp;</td>
									<td colspan="3" align="left">
										<select name="cmbComboPerfil" class="TextBox">
											<option value="">Selecione</option>
											
											<%
											
											If vint_FlPerfil = "0" Then
												
												%>
												
												<option selected value="0">Colaborador Nivel - 1</option>
												<option value="2">Colaborador Nivel - 2</option>								
												<option value="1">Administrador</option>
												
												<%
												
											ElseIf vint_FlPerfil = "1" Then
												
												%>
												
												<option value="0">Colaborador Nivel - 1</option>
												<option value="2">Colaborador Nivel - 2</option>								
												<option selected value="1">Administrador</option>
												
												<%
												
											ElseIf vint_FlPerfil = "2" Then
												
												%>
												<option value="0">Colaborador Nivel - 1</option>
												<option selected value="2">Colaborador Nivel - 2</option>	
												<option value="1">Administrador</option>
												
												<%
												
											Else
												
												%>
												
												<option value="0">Colaborador Nivel - 1</option>
												<option value="2">Colaborador Nivel - 2</option>						
												<option value="1">Administrador</option>
												
												<%
												
											End If
											
											%>
											
										</select>
									</td>
								</tr>
								<tr>
			       <!-- ------------------------------------------------------------------------------ -->
			       						
									
									
									<td align="left">Equipe:&nbsp;</td>
									
									
								   <td> <select name="txtDsEquipe" class="TextBox">
									<option selected value="">Selecione</option>
									
									
									<%
									
									If vint_IdEquipe="Azul" Then
									
									 %>
									
									<option selected value="Azul">Azul</option>
												
									
									<%
									
									Else If vint_IdEquipe="Laranja" Then
									
									 %>
									
									            
												<option selected value="Laranja">Laranja</option>								
												
									
									<%
									
									Else If vint_IdEquipe="Inativo" Then
									
									 %>
									
									            
												<option selected value="Inativo">Inativo</option>								
												
									
									
									
									<%
									
									Else If vint_IdEquipe="Vermelha" Then
									
									 %>
									
									            
												<option selected value="Vermelha">Vermelha</option>
												
									<%
									
									Else If vint_IdEquipe="Verde" Then
									
									 %>
									
									            
												<option selected value="Vermelha">Verde</option>
												
									
									
									<%
									
									Else If vint_IdEquipe="Roxa" Then
									
									 %>
																		            
												<option selected value="Roxa">Roxa</option>
									
									
									
									<%
									
									End IF
									End IF
									End IF
									End IF
									End IF
									End IF
									
									 %>
									
									            <option value="Azul">Azul</option>
									            <option value="Inativo">Inativo</option>
												<option value="Laranja">Laranja</option>								
												<option value="Vermelha">Vermelha</option>
												<option value="Verde">Verde</option>
												<option value="Roxa">Roxa</option>
									
									
									</select>
									</td>
									
									
									<td align="left">Telefone:&nbsp;</td>
									<td align="left">
										<input type="text" name="txtDsTelefone" class="TextBox" size="17" maxlength="20" value="<%=vstr_DsTelefone%>">
									</td>
									<td nowrap align="left">Ramal:&nbsp;</td>
									<td align="left">
										<input type="text" name="txtDsRamal" class="TextBox" size="7" maxlength="15" value="<%=vstr_DsRamal%>">
									</td>
									<td nowrap align="left">Local Alocado:&nbsp;</td>
									<td align="left">
										<input type="text" name="txtDsLocalAlocado" class="TextBox" size="25" maxlength="30" value="<%=vstr_DsLocalAlocado%>">
									</td>
								</tr>
								<tr>
									<td align="left">Senha:&nbsp;</td>
									<td align="left" colspan="3">
										<input type="password" name="txtCdSenha" id="txtCdSenha" class="TextBox" size="25" maxlength="15" value="<%=DesencriptaString(vstr_CdSenha)%>">
									</td>
									<td nowrap align="left">Confirma Senha:&nbsp;</td>
									<td align="left">
										<input type="password" name="txtCdConfirmaSenha" id="txtCdConfirmaSenha" class="TextBox" size="25" maxlength="15" value="<%=DesencriptaString(vstr_CdSenha)%>">
									</td>
								</tr>
							</table>
						</fieldset>
						</td>
					</tr>
					<tr>
						<td colspan="2" align="middle">
						&nbsp;
						</td>
					</tr>
					<tr>
						<td colspan="2" align="middle">
							<table ALIGN="center" BORDER="0" CELLSPACING="1" CELLPADDING="1">
								<tr>
									<td><input type="Submit" name="cmdSalvar" value="Salvar" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Gravar dados"></td>
									<td><input type="button" name="cmdRetornar" value="Retornar" onClick="voltar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Retornar a tela anterior"></td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>

<!-- #include file = "../includes/LayoutEnd.asp" -->

<%
' =============================================================================================
' DECLARAÇÃO DE FUNÇÕES E PROCEDIMENTOS LOCAIS DA PÁGINA.
' =============================================================================================

' Função desenvolvida para fazer o tratamento do
' formulário de dados.
Private Function ValidarForm()
	
	' Tratamento de campos do formulário. =============================
	
	If Trim(vstr_CdUsuario) = "" Then
		
		Call AddErro("CdUsario", "Favor, preencher o campo Usuário.")
	End If
	
	If Trim(vstr_DsUsuario) = "" Then
		
		Call AddErro("Nome", "Favor, preencher o campo Nome.")
	End If
	
	If Trim(vstr_DsCPF) = "" Then
		
		Call AddErro("CPF", "Favor, preencher o campo CPF.")
	Else
		
		If Not Len(Trim(vstr_DsCPF)) = 11 Then
			
			Call AddErro("CPF", "Favor, preencher o campo CPF.")
		Else
		
			If Not IsNumeric(vstr_DsCPF)Then
				
				Call AddErro("CPF", "Favor, preencher o campo CPF apenas com números, Ex. 12345678901.")
			End If
		End If
	End If
	
	If Trim(vstr_DsRG) = "" Then
		
		Call AddErro("RG", "Favor, preencher o campo RG.")
	End If
	
	If Trim(vint_IdFuncao) = "" Then
		
		Call AddErro("Funcao", "Favor, Selecionar uma Função.")
	End If
	
	If Trim(vint_FlPerfil) = "" Then
		
		Call AddErro("Perfil", "Favor, Selecionar um Perfil.")
	End If
	
	If Trim(vstr_CdSenha) = "" Then
		
		Call AddErro("Senha", "Favor, preencher o campo Senha.")
	Else
		
		If Trim(vstr_CdConfirmaSenha) = "" Then
		
			Call AddErro("ConfirmaSenha", "Favor, preencher o campo Confirma Senha.")
		Else
			If Not vstr_CdSenha = vstr_CdConfirmaSenha Then
				
				Call AddErro("ConfirmarSenha", "Favor, digitar a mesma senha.")
			End If
		End If
	End If
	
	If Not isDate(vstr_DtNascimento) Then
		
		Call AddErro("Data", "Favor, digite uma data de nascimento válida. Ex 01/01/1900")
		
	End If
	
	
	' Verifica se algum tipo de erro
	' ocorreu na validação do formulário.
	If TotalErros > 0 Then
		
		' Formulário inválido.
		ValidarForm = False
	Else
		
		' Formulário válido.
		ValidarForm = True
	End If
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
								' COMBO DE FUNÇÃO
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub CriarComboFuncao(pstr_Nome, pstr_ValorDefault, pstr_onChange, pstr_Evento)
	
	
	' Declaração de variáveis locais.
	Dim vobj_command
	Dim vobj_rs
	
	
	' ---------------------------------------------------------------------
	' Selecionando todos os Registros
	' ---------------------------------------------------------------------
	Set vobj_command = Server.CreateObject("ADODB.Command")
	Set vobj_command.ActiveConnection = vobj_conexao
	
	
	vobj_command.CommandType				= adCmdStoredProc
	vobj_command.CommandText				= "consultaFuncoes"
	vobj_command.Parameters.Refresh
	
	
	Set vobj_rs = vobj_command.Execute
	' ---------------------------------------------------------------------
	
	
	%>
	<SELECT name="<%=pstr_Nome%>" onChange="<%=pstr_onChange%>" <%=pstr_Evento%> class="TextBox">
		<option value="<%=Empty%>">SELECIONE</option>
		<%
		
		
		If Not vobj_rs.EOF Then
			
			' Loop de todos os registros encontrados.
			Do While Not vobj_rs.EOF
				
				
				' Verificando se o registro
				' é o Registro default que deve ser selecionado.
				If Trim(pstr_ValorDefault) = Trim(vobj_rs("ID_FUNCAO")) Then
					%><OPTION selected value="<%=vobj_rs("ID_FUNCAO")%>"><%=vobj_rs("DS_FUNCAO")%></OPTION><%
				Else
					%><OPTION value="<%=vobj_rs("ID_FUNCAO")%>"><%=vobj_rs("DS_FUNCAO")%></OPTION><%
				End If
				
				
				vobj_rs.MoveNext
			Loop
		End If
		
		
		%>
	</SELECT>
	<%
		
			
	vobj_rs.Close
	Set vobj_rs = Nothing
	Set vobj_command = Nothing
	
End Sub

%>

<!-- #include file = "../includes/CloseConnection.asp" -->