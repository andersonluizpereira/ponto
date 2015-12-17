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
Dim vstr_IdProjeto

' Declaração de variáveis utilizadas para armazenar os
' valores dos campos da tela.
Dim vstr_DsProjeto
Dim vstr_DsDescricao
Dim vstr_DtInicio
Dim vstr_DtFinal
'Dim vint_FlAtivo
Dim vstr_DsArea

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
	vobj_commandDesativar.CommandText					= "excluiProjeto"
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
		
	Response.Redirect("projetoslistagem.asp")
	
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
	vobj_commandAtivar.CommandText					= "ativarProjeto"
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
		
	Response.Redirect("projetoslistagem.asp")
	
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
	vstr_IdProjeto				= Request.Form("hdnIdRegistro")
	
	
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
	vobj_commandRegistro.CommandText					= "consultaProjeto"
	
	vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("param1", adChar, adParamInput, 10, vstr_IdProjeto)
	' ---------------------------------------------------------------------
	
	
	' Cria o objeto recordset com as informações do registro.	
	Set vobj_rsRegistro = vobj_commandRegistro.Execute
	
	
	If Not vobj_rsRegistro.EOF Then
		
		' Conseguindo os dados do registro.
		vstr_DsProjeto			= vobj_rsRegistro("DS_PROJETO")
		vstr_DsDescricao		= vobj_rsRegistro("DS_DESCRICAO")
		vstr_DtInicio			= converterDataParaHtml(vobj_rsRegistro("DT_INICIO"))
		vstr_DtFinal			= converterDataParaHtml(vobj_rsRegistro("DT_FINAL"))
		'vint_FlAtivo			= vobj_rsRegistro("FL_ATIVO")
		vstr_DsArea				= vobj_rsRegistro("DS_AREA")
		
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
		
		vstr_IdProjeto			= Empty
		vstr_DsProjeto			= Empty
		vstr_DsDescricao		= Empty
		vstr_DtInicio			= Empty
		vstr_DtFinal			= Empty
		'vint_FlAtivo			= Empty
		vstr_DsArea				= Empty
		
	Else
		
		' ... está opção acontecerá quando o usuário processar
		' a página, neste caso todas os dados da tela serão
		' submetidos e devem ser pegos neste lugar.
		
		vstr_DsProjeto			= Request.Form("txtDsProjeto")
		vstr_DsDescricao		= Request.Form("txtDsDescricao")
		vstr_DtInicio			= Request.Form("txtDtInicio")
		vstr_DtFinal			= Request.Form("txtDtFinal")
		'vint_FlAtivo			= Request.Form("txtFlAtivo")
		vstr_DsArea				= Request.Form("cmbComboArea")
		
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
			
			vstr_IdProjeto				= Request.Form("hdnIdRegistro")
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
				
				
				' ---------------------------------------------------------------------
				' Alterando os dados do registro no banco de dados.
				' ---------------------------------------------------------------------
				Set vobj_commandProc = Server.CreateObject("ADODB.Command")
				Set vobj_commandProc.ActiveConnection = vobj_conexao
				
				
				vobj_commandProc.CommandType					= adCmdStoredProc
				vobj_commandProc.CommandText					= "alteraProjeto"
				
				'If vint_FlAtivo = "" Then
				'		
				'	vint_FlAtivo = 0
				'Else
				'	If vint_FlAtivo = "on" Then
				'			
				'		vint_FlAtivo = 1
				'	End If
				'End If
				
				
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_IdProjeto))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 10, Trim(vstr_DsArea))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adChar, adParamInput, 100, Trim(vstr_DsDescricao))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adDate, adParamInput,, converterDataParaSQL(vstr_DtInicio))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adDate, adParamInput,, converterDataParaSQL(vstr_DtFinal))
				'vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adBoolean, adParamInput,, vint_FlAtivo)
					
				Call vobj_commandProc.Execute
				Set vobj_commandProc = Nothing
				' ---------------------------------------------------------------------
				
				
				' Redireciona para a página de listagem
				' dos registros.
				Response.Redirect("projetoslistagem.asp")
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
				vobj_commandRegistroConsulta.CommandText					= "consultaProjeto"
				
				
				vobj_commandRegistroConsulta.Parameters.Append vobj_commandRegistroConsulta.CreateParameter("param1",adChar, adParamInput, 10, vstr_DsProjeto)
				' ---------------------------------------------------------------------
				
				
				' Cria o objeto recordset com as informações do registro.	
				Set vobj_rsRegistroConsulta = vobj_commandRegistroConsulta.Execute
				
				'Verificando se ja ha registro no banco com mesma area
				'Obs. É verificado soment campo Area, campo nome pode haver dois iguais.
				If Not vobj_rsRegistroConsulta.EOF Then
					
					Call AddErro("Erro", "Há um registro com o mesma nome de Projeto.")
					
				Else
					
					Dim vobj_rs
					
					' ---------------------------------------------------------------------
					' Incluindo os dados do registro no banco de dados.
					' ---------------------------------------------------------------------
					Set vobj_commandProc = Server.CreateObject("ADODB.Command")
					Set vobj_commandProc.ActiveConnection = vobj_conexao
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "incluiProjeto"
					
					'If vint_FlAtivo = "" Then
					'	
					'	vint_FlAtivo = 0
					'Else
					'	If vint_FlAtivo = "on" Then
					'		
					'		vint_FlAtivo = 1
					'	End If
					'End If
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_DsProjeto))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 10, Trim(vstr_DsArea))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adChar, adParamInput, 100, Trim(vstr_DsDescricao))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adDate, adParamInput,, converterDataParaSQL(vstr_DtInicio))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adDate, adParamInput,, converterDataParaSQL(vstr_DtFinal))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adDate, adParamInput,, converterDataParaSQL(Date()))
					'vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param7",adBoolean, adParamInput,, vint_FlAtivo)
					
					vobj_commandProc.Execute
					
					
					Set vobj_commandProc = Nothing
					
					
					' Altera a variável que indica o tipo de
					' operação que é executada na página.
					vstr_Operacao = "A"
					
					
					' Redireciona para a página de listagem
					' dos registros.
					Response.Redirect("projetoslistagem.asp")
					
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

<script type="text/javascript" src="js/projetosmanutencao.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td valign="top">
			<form name="thisForm" action="projetosmanutencao.asp" method="post">
						
				<input type="hidden" name="hdnProcessar" value="S">
				<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>">
				<input type="hidden" name="hdnIdRegistro" value="<%=vstr_IdProjeto%>">
						
				<i><b class="TituloPagina">Projetos</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td><%=ExibirErros()%></td>
					</tr>
					<tr>
						<td colspan="2">
							<fieldset style="LEFT: 0px; WIDTH: 596px; HEIGHT: 120px">
								<legend>
									<b>Dados do Projeto</b>
								</legend>
								<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabPesquisa" id="tabPesquisa" width="100%" >
									<tr>
										<td align="left">Projeto:&nbsp;</td>
										<td align="left"><input name="txtDsProjeto" id="IDE" class="TextBox" size="15" maxlength="10" value="<%=vstr_DsProjeto%>"></td>
										<td align="left">Nome:&nbsp;</td>
										<td align="left"><input name="txtDsDescricao" id="Descricao" class="TextBox" size="45" maxlength="100" value="<%=vstr_DsDescricao%>"></td>
									</tr>
									<tr>
										<td align="left">Objetivo(s):&nbsp;</td>
										<td align="left"><input name="txtDtInicio" id="Inicio" class="TextBox" size="12" maxlength="10" value="<%=vstr_DtInicio%>"></td>
										<td align="left">Final:&nbsp;</td>
										<td align="left"><input name="txtDtFinal" id="Final" class="TextBox" size="12" maxlength="10" value="<%=vstr_DtFinal%>"></td>
									</tr>
									<tr>
									</tr>
									<tr>
										<td align="left">Área:&nbsp;</td>
										<td colspan="3" align="left">
											<%Call CriarComboArea("cmbComboArea", vstr_DsArea,Empty,Empty) %>
										</td>
									</tr>
								</table>
							</fieldset>
						</td>
					</tr>
					<tr>
						<td align="center">
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
	
	If Trim(vstr_DsProjeto) = "" Then
		
		Call AddErro("Projeto", "Favor, preencher o campo Projeto.")
	End If
	
	If Trim(vstr_DsArea) = "" Then
		
		Call AddErro("Projeto", "Favor, selecionar uma Area para o projeto.")
	End If
	
	If Trim(vstr_DsDescricao) = "" Then
		
		Call AddErro("Nome", "Favor, preencher o campo Nome.")
	End If
	
	If Trim(vstr_DtInicio) = "" Then
		
		Call AddErro("Inicio", "Favor, preencher o campo Objetivo.")
	Else
		
		If Not isDate(vstr_DtInicio) Then
			
			Call AddErro("isDataInicio", "Favor, preencher o campo Objetivo com uma data válida, Ex. 01/01/2009.")
		End If
	End If
	
	If Trim(vstr_DtFinal) = "" Then
		
		Call AddErro("Final", "Favor, preencher o campo Final.")
	Else
		
		If Not isDate(vstr_Dtfinal) Then
			
			Call AddErro("isDataFinal", "Favor, preencher o campo Final com uma data válida, Ex. 01/01/2009.")
		End If
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
								' COMBO DE AREA
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub CriarComboArea(pstr_Nome, pstr_ValorDefault, pstr_onChange, pstr_Evento)
	
	
	' Declaração de variáveis locais.
	Dim vobj_command
	Dim vobj_rs
	
	
	' ---------------------------------------------------------------------
	' Selecionando todos os Registros
	' ---------------------------------------------------------------------
	Set vobj_command = Server.CreateObject("ADODB.Command")
	Set vobj_command.ActiveConnection = vobj_conexao
	
	
	vobj_command.CommandType				= adCmdStoredProc
	vobj_command.CommandText				= "consultaAreasAtivo"
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
				If Trim(pstr_ValorDefault) = Trim(vobj_rs("DS_AREA")) Then
					%><OPTION selected value="<%=vobj_rs("DS_AREA")%>"><%=UCASE(vobj_rs("DS_AREA"))%></OPTION><%
				Else
					%><OPTION value="<%=vobj_rs("DS_AREA")%>"><%=UCASE(vobj_rs("DS_AREA"))%></OPTION><%
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