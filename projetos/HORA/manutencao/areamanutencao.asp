<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- S#include file = "../includes/GetConnection.asp" -->
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
Dim vstr_IdArea

' Declaração de variáveis utilizadas para armazenar os
' valores dos campos da tela.
Dim vstr_DsArea
Dim vstr_DsNome

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
	vobj_commandDesativar.CommandText					= "excluiArea"
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
		
	Response.Redirect("arealistagem.asp")
	
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
	vobj_commandAtivar.CommandText					= "ativarArea"
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
		
	Response.Redirect("arealistagem.asp")
	
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
	vstr_IdArea				= Request.Form("hdnIdRegistro")
	
	
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
	vobj_commandRegistro.CommandText					= "consultaArea"
	
	vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("param1",adChar, adParamInput, 10, vstr_IdArea)
	' ---------------------------------------------------------------------
	
	
	' Cria o objeto recordset com as informações do registro.	
	Set vobj_rsRegistro = vobj_commandRegistro.Execute
	
	
	If Not vobj_rsRegistro.EOF Then
		
		' Conseguindo os dados do registro.
		vstr_IdArea		= vobj_rsRegistro("DS_AREA")
		vstr_DsArea		= vobj_rsRegistro("DS_AREA")
		vstr_DsNome		= vobj_rsRegistro("DS_NOME")
		
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
		
		vstr_IdArea				= Empty
		vstr_DsArea				= Empty
		vstr_DsNome				= Empty
		
	Else
		
		' ... está opção acontecerá quando o usuário processar
		' a página, neste caso todas os dados da tela serão
		' submetidos e devem ser pegos neste lugar.
		
		vstr_IdArea					= Request.Form("hdnIdRegistro")
		vstr_DsArea					= Request.Form("txtDsArea")
		vstr_DsNome					= Request.Form("txtDsNome")
		
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
			
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
				
				
				' ---------------------------------------------------------------------
				' Alterando os dados do registro no banco de dados.
				' ---------------------------------------------------------------------
				Set vobj_commandProc = Server.CreateObject("ADODB.Command")
				Set vobj_commandProc.ActiveConnection = vobj_conexao
				
				
				vobj_commandProc.CommandType					= adCmdStoredProc
				vobj_commandProc.CommandText					= "alteraArea"
				
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, vstr_IdArea)
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 10, Trim(vstr_DsArea))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adChar, adParamInput, 100, Trim(vstr_DsNome))
				
				
				Call vobj_commandProc.Execute
				Set vobj_commandProc = Nothing
				' ---------------------------------------------------------------------
				
				
				' Redireciona para a página de listagem
				' dos registros.
				Response.Redirect("arealistagem.asp")
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
				vobj_commandRegistroConsulta.CommandText					= "consultaArea"
				
				
				vobj_commandRegistroConsulta.Parameters.Append vobj_commandRegistroConsulta.CreateParameter("param1",adChar, adParamInput, 10, vstr_DsArea)
				' ---------------------------------------------------------------------
				
				
				' Cria o objeto recordset com as informações do registro.	
				Set vobj_rsRegistroConsulta = vobj_commandRegistroConsulta.Execute
				
				'Verificando se ja ha registro no banco com mesma area
				'Obs. É verificado soment campo Area, campo nome pode haver dois iguais.
				If Not vobj_rsRegistroConsulta.EOF Then
					
					Call AddErro("Erro", "Há um registro com a mesma sigla de area.")
					
				Else
					
					Dim vobj_rs
					
					' ---------------------------------------------------------------------
					' Incluindo os dados do registro no banco de dados.
					' ---------------------------------------------------------------------
					Set vobj_commandProc = Server.CreateObject("ADODB.Command")
					Set vobj_commandProc.ActiveConnection = vobj_conexao
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "incluiArea"
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_DsArea))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 100, Trim(vstr_DsNome))
					
					vobj_commandProc.Execute
					
					
					Set vobj_commandProc = Nothing
					
					
					' Altera a variável que indica o tipo de
					' operação que é executada na página.
					vstr_Operacao = "A"
					
					
					' Redireciona para a página de listagem
					' dos registros.
					Response.Redirect("arealistagem.asp")
					
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

<script type="text/javascript" src="js/areamanutencao.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
		<form name="thisForm" action="areamanutencao.asp" method="post">
			
			<input type="hidden" name="hdnProcessar" value="S">
			<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>">
			<input type="hidden" name="hdnIdRegistro" value="<%=vstr_IdArea%>">
			
			<i><b class="TituloPagina">Áreas</b></i>
			<table border="0" class="font" cellpadding="0" cellspacing="0">
				<tr>
					<td><%=ExibirErros()%></td>
				</tr>
				<tr>
					<td>
					<fieldset style="LEFT: 0px; WIDTH: 595px; HEIGHT: 59px">
						<legend>
						   <b>Dados da Área</b>
						</legend>
						<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabPesquisa" id="tabPesquisa" width="100%" style="FILTER: alpha(opacity  =80)">
							<tr>
								<td>Sigla:&nbsp;</td>
								<td><input name="txtDsArea" id="Area" class="TextBox" size="15" maxlength="10" Value="<%=vstr_DsArea%>"></td>
								<td>Descrição:&nbsp;</td>
								<td><input name="txtDsNome" id="Nome" class="TextBox" size="45" maxlength="100" Value="<%=vstr_DsNome%>"></td>
							</tr>
						</table>
					</fieldset>
					</td>
				</tr>
				<tr>
					<td align="middle">
					&nbsp;
					</td>
				</tr>
				<tr>
					<td align="middle">
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
</TABLE>

<!-- #include file = "../includes/LayoutEnd.asp" -->


<%
' =============================================================================================
' DECLARAÇÃO DE FUNÇÕES E PROCEDIMENTOS LOCAIS DA PÁGINA.
' =============================================================================================

' Função desenvolvida para fazer o tratamento do
' formulário de dados.
Private Function ValidarForm()
	
	' Tratamento de campos do formulário. =============================
	
	If Trim(vstr_DsArea) = "" Then
		Call AddErro("Sigla", "Favor, preencher da Siglas.")
	Else
		If Len(Trim(vstr_DsArea)) > 10 Then
			Call AddErro("QtdSigla", "Digite apenas 10 caracteres no campo Sigla.")
		End If
	End If
	
	If Trim(vstr_DsNome) = "" Then
		Call AddErro("Descrição", "Favor, preencher o campo Descrição.")
	Else
		If Len(Trim(vstr_DsNome)) > 100 Then
			Call AddErro("QtdDescrição", "Digite apenas 100 caracteres no campo Descrição.")
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

%>

<!-- #include file = "../includes/CloseConnection.asp" -->