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


' Declara��o de vari�veis locais. ==============================================

' Guarda a opera��o que ser� executa nesta tela.
' Obs.: Seus valores podem ser A = Altera��o, I = Inclus�o, V = Visualiza��o.
Dim vstr_Operacao

' Vari�vel flag que indica se a p�gina deve ser 
' processada, apenas disponivel para as opera��es de
' A e I.
Dim vstr_Processar

'Vari�vel de controle e fluxo de acoes
Dim vstr_Executar

' Armazena o c�digo de refer�ncia do registro que ser� alterado, Inclusso
' ou visualizado.
Dim vint_IdAtividade

' Declara��o de vari�veis utilizadas para armazenar os
' valores dos campos da tela.
Dim vstr_DsAtividade

' para est� p�gina.
vstr_Operacao		= Request.Form("pstr_Operacao")
vstr_Processar		= Request.Form("hdnProcessar")
vstr_Executar		= Request.Form("hdnExecutar")

' Verifica se o parametro que defini o tipo
' de opera��o a ser executado na p�gina �
' igual a branco(vazio).
If Trim(vstr_Operacao) = "" Then
	
	' ... neste caso a opera��o
	' padr�o de a de visualiza��o apenas
	' do registro.
	vstr_Operacao = "V"
End If


' Analizando a vari�vel que indica o fluxo de 
' opera��o desta p�gina.
If vstr_Executar = "DESATIVAR" Then
	
	
	' -->> Opera��o de Exclus�o de Registros.
		
		
	' ... opera��o de exclus�o de registros.
		
		
	' Declara��o de vari�veis auxiliares que
	' auxiliar�o na exclus�o dos registros selecionados.
	Dim vstr_DesativarIdRegistro
	Dim vobj_commandDesativar
		
		
		
		
	' Conseguindo todos os registros selecionados para
	' a exclus�o do banco de dados.
	vstr_DesativarIdRegistro = Request.Form("hdnIdRegistro")
		
		
	' ---------------------------------------------------------------------
	' Exclus�o de Registros do banco de dados.
	' ---------------------------------------------------------------------
	Set vobj_commandDesativar = Server.CreateObject("ADODB.Command")
	Set vobj_commandDesativar.ActiveConnection = vobj_conexao
			
			
	vobj_commandDesativar.CommandType					= adCmdStoredProc
	vobj_commandDesativar.CommandText					= "excluiAtividade"
	vobj_commandDesativar.CommandTimeout					= 0	
		
			
			
	' Iguinorando os erros que ocorrem na exclus�o
	' do registro do banco de dados.
	On Error Resume Next
		
		
	' Passando o c�digo do Registro a ser
	' excluido do banco de dados.
	vobj_commandDesativar.Parameters.Append vobj_commandDesativar.CreateParameter("param1", adInteger, adParamInput, , vstr_DesativarIdRegistro)
		
		
	' Chamando comando para excluir o registro
	Call vobj_commandDesativar.Execute
			
	' Analizando os erros que podem ter ocorrido
	' na exclus�o do registros selecionados pelo
	' usu�rio.
	Select Case Err.number 
			
		' Verificando se o erro de integridade referencial
		' ocorreu na exclus�o do registro acima.
		Case -2147217900
					
			%><script>alert("Aten��o!!!\n\nO(s) registro(s) que n�o foi(ram) excluido(s) possui(em) dados relacionados. Exclua os dados relacionados para poder excluir este(s) registro(s).");</script><%
					
	End Select
			
			
			
	' Habilitando a mensagem de erro quando um
	' erro acontecer.
	On Error Goto 0
			
			
			
	' Limpa a vari�vel utilizada para excluir os
	' registros do banco de dados.
	Set vobj_commandDesativar = Nothing
		
	Response.Redirect("atividadelistagem.asp")
	
End If

' Analizando a vari�vel que indica o fluxo de 
' opera��o desta p�gina.
If vstr_Executar = "ATIVAR" Then
	
	
	' -->> Opera��o de Exclus�o de Registros.
		
		
	' ... opera��o de exclus�o de registros.
		
		
	' Declara��o de vari�veis auxiliares que
	' auxiliar�o na exclus�o dos registros selecionados.
	Dim vstr_AtivarIdRegistro
	Dim vobj_commandAtivar
		
		
		
		
	' Conseguindo todos os registros selecionados para
	' a exclus�o do banco de dados.
	vstr_AtivarIdRegistro = Request.Form("hdnIdRegistro")
		
		
	' ---------------------------------------------------------------------
	' Exclus�o de Registros do banco de dados.
	' ---------------------------------------------------------------------
	Set vobj_commandAtivar = Server.CreateObject("ADODB.Command")
	Set vobj_commandAtivar.ActiveConnection = vobj_conexao
			
			
	vobj_commandAtivar.CommandType					= adCmdStoredProc
	vobj_commandAtivar.CommandText					= "ativarAtividade"
	vobj_commandAtivar.CommandTimeout					= 0	
		
			
			
	' Iguinorando os erros que ocorrem na exclus�o
	' do registro do banco de dados.
	On Error Resume Next
		
	' Passando o c�digo do Registro a ser
	' excluido do banco de dados.
	vobj_commandAtivar.Parameters.Append vobj_commandAtivar.CreateParameter("param1",adChar, adParamInput, 10, vstr_AtivarIdRegistro)
	
	
	' Chamando comando para excluir o registro
	Call vobj_commandAtivar.Execute
			
	' Analizando os erros que podem ter ocorrido
	' na exclus�o do registros selecionados pelo
	' usu�rio.
	Select Case Err.number 
				
		' Verificando se o erro de integridade referencial
		' ocorreu na exclus�o do registro acima.
		Case -2147217900
					
			%><script>alert("Aten��o!!!\n\nO(s) registro(s) que n�o foi(ram) excluido(s) possui(em) dados relacionados. Exclua os dados relacionados para poder excluir este(s) registro(s).");</script><%
					
	End Select
			
			
			
	' Habilitando a mensagem de erro quando um
	' erro acontecer.
	On Error Goto 0
			
			
			
	' Limpa a vari�vel utilizada para excluir os
	' registros do banco de dados.
	Set vobj_commandAtivar = Nothing
		
	Response.Redirect("atividadelistagem.asp")
	
End If

' *******************************************************
' INICIO DA ROTINA QUE CONSEGUE OS DADOS DO REGISTRO
' *******************************************************

' Veririfa se a opera��o a ser executada nesta p�gina � a 
' opera��o de Altera��o ou Visualiza��o e se a p�gina n�o
' foi processada ainda.
If (vstr_Operacao = "A" or vstr_Operacao = "V") And vstr_Processar <> "S" Then
	
	
	' ... neste caso deve ser solicitado o c�digo do registro
	' e encontrar suas informa��es no banco de dados para exibir para
	' as informa��es do registro na tela.
	' Conseguindo o c�digo do registro.
	vint_IdAtividade				= Request.Form("hdnIdRegistro")
	
	
	' Declara��o de vari�veis auxiliares
	' para obter as informa��es do registro.
	Dim vobj_rsRegistro
	Dim vobj_commandRegistro
	
	
	
	' ---------------------------------------------------------------------
	' Selecionando os dados do registro.
	' ---------------------------------------------------------------------
	Set vobj_commandRegistro = Server.CreateObject("ADODB.Command")
	Set vobj_commandRegistro.ActiveConnection = vobj_conexao
							
							
	vobj_commandRegistro.CommandType					= adCmdStoredProc
	vobj_commandRegistro.CommandText					= "consultaAtividade"
	
	vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("param1", adInteger, adParamInput, , vint_IdAtividade)
	' ---------------------------------------------------------------------
	
	
	' Cria o objeto recordset com as informa��es do registro.	
	Set vobj_rsRegistro = vobj_commandRegistro.Execute
	
	
	If Not vobj_rsRegistro.EOF Then
		
		' Conseguindo os dados do registro.
		vint_IdAtividade		= vobj_rsRegistro("ID_ATIVIDADE")
		vstr_DsAtividade		= vobj_rsRegistro("DS_ATIVIDADE")
		
	End If
	
	vobj_rsRegistro.Close
	Set vobj_rsRegistro = Nothing
	Set vobj_commandRegistro = Nothing
Else
	
	
	' Verifica se a opera��o a ser executada nesta p�gina �
	' a opera��o de inclus�o e verifica se a p�gina n�o foi
	' processada ainda.
	If vstr_Operacao = "I" And vstr_Processar <> "S" Then
		
		' Neste caso todas as vari�veis devem ser vazias
		' para o usu�rio poder preencher seu novo cadastro
		' do registro.
		
		vint_IdAtividade		= Empty
		vstr_DsAtividade		= Empty
		
	Else
		
		' ... est� op��o acontecer� quando o usu�rio processar
		' a p�gina, neste caso todas os dados da tela ser�o
		' submetidos e devem ser pegos neste lugar.
		
		vint_IdAtividade					= Request.Form("hdnIdRegistro")
		vstr_DsAtividade					= Request.Form("txtDsAtividade")
		
	End If
End If
' *******************************************************
' FINAL DA ROTINA QUE CONSEGUE OS DADOS DO REGISTRO
' *******************************************************



' *******************************************************
' INICIO DA ROTINA QUE FAZ O PROCESSAMENTO DOS DADOS
' DO REGISTRO.
' *******************************************************

' Verifica se a vari�vel flag est� setada como S, 
' isto indica que um processamento deve ser feito.
If vstr_Processar = "S" Then
	
	
	' Declara��o de vari�veis auxiliares
	' para fazer o processamento da p�gina.
	Dim vobj_commandProc
	
	
	' Analiza a opera��o a ser executada na p�gina
	' para descobrir o processamento que deve ser feito.
	Select Case vstr_Operacao
		
		
		Case "A"						' Opera��o de altera��o do registro.
			
			' ... processamento de altera��o do registro.
			
			
			' Verificando se o formul�rio foi
			' devidamente v�lidado pelo sistema.
			If ValidarForm = True Then
				
				
				' ---------------------------------------------------------------------
				' Alterando os dados do registro no banco de dados.
				' ---------------------------------------------------------------------
				Set vobj_commandProc = Server.CreateObject("ADODB.Command")
				Set vobj_commandProc.ActiveConnection = vobj_conexao
				
				
				vobj_commandProc.CommandType					= adCmdStoredProc
				vobj_commandProc.CommandText					= "alteraAtividade"
				
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1", adInteger, adParamInput, , vint_IdAtividade)
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 100, Trim(vstr_DsAtividade))
				
				Call vobj_commandProc.Execute
				Set vobj_commandProc = Nothing
				' ---------------------------------------------------------------------
				
				
				' Redireciona para a p�gina de listagem
				' dos registros.
				Response.Redirect("atividadelistagem.asp")
			End If
			
			
		Case "I"						' Opera��o de inclus�o do registro.
			
			
			' ... processamento de inclus�o de registro.
			
			
			' Verificando se o formul�rio foi
			' devidamente v�lidado pelo sistema.
			If ValidarForm = True Then
				
				' ---------------------------------------------------------------------
				' Procedimento desenvolvimento para tratar a entrada de umas mesma
				' area
				' ---------------------------------------------------------------------
				
				' Declara��o de vari�veis auxiliares
				' para obter as informa��es do registro.
				Dim vobj_rsRegistroConsulta
				Dim vobj_commandRegistroConsulta
				
				
				' ---------------------------------------------------------------------
				' Selecionando os dados do registro.
				' ---------------------------------------------------------------------
				Set vobj_commandRegistroConsulta = Server.CreateObject("ADODB.Command")
				Set vobj_commandRegistroConsulta.ActiveConnection = vobj_conexao
				
				
				vobj_commandRegistroConsulta.CommandType					= adCmdStoredProc
				vobj_commandRegistroConsulta.CommandText					= "consultaAtividade"
				
				
				vobj_commandRegistroConsulta.Parameters.Append vobj_commandRegistroConsulta.CreateParameter("param1", adInteger, adParamInput, , 0)
				' ---------------------------------------------------------------------
				
				
				' Cria o objeto recordset com as informa��es do registro.	
				Set vobj_rsRegistroConsulta = vobj_commandRegistroConsulta.Execute
				
				'Verificando se ja ha registro no banco com mesma area
				'Obs. � verificado soment campo Area, campo nome pode haver dois iguais.
				If Not vobj_rsRegistroConsulta.EOF Then
					
					Call AddErro("Erro", "H� um registro com o mesmo nome d Atividade.")
					
				Else
					
					Dim vobj_rs
					
					' ---------------------------------------------------------------------
					' Incluindo os dados do registro no banco de dados.
					' ---------------------------------------------------------------------
					Set vobj_commandProc = Server.CreateObject("ADODB.Command")
					Set vobj_commandProc.ActiveConnection = vobj_conexao
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "incluiAtividade"
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 100, Trim(vstr_DsAtividade))
					
					vobj_commandProc.Execute
					
					
					Set vobj_commandProc = Nothing
					
					
					' Altera a vari�vel que indica o tipo de
					' opera��o que � executada na p�gina.
					vstr_Operacao = "A"
					
					
					' Redireciona para a p�gina de listagem
					' dos registros.
					Response.Redirect("atividadelistagem.asp")
					
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

<script type="text/javascript" src="js/atividademanutencao.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
		<form name="thisForm" action="atividademanutencao.asp" method="post">
			
			<input type="hidden" name="hdnProcessar" value="S">
			<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>">
			<input type="hidden" name="hdnIdRegistro" value="<%=vint_IdAtividade%>">
			
			<i><b class="TituloPagina">Atividade</b></i>
			<table border="0" class="font" cellpadding="0" cellspacing="0">
				<tr>
					<td><%=ExibirErros()%></td>
				</tr>
				<tr>
					<td>
					<fieldset style="LEFT: 0px; WIDTH: 595px; HEIGHT: 59px">
						<legend>
						   <b>Dados da Atividade</b>
						</legend>
						<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabPesquisa" id="tabPesquisa" style="FILTER: alpha(opacity  =80)">
							<tr>
								<td>Descri��o:&nbsp;</td>
								<td><input name="txtDsAtividade" id="Descricao" class="TextBox" size="54" maxlength="100" Value="<%=vstr_DsAtividade%>"></td>
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
' DECLARA��O DE FUN��ES E PROCEDIMENTOS LOCAIS DA P�GINA.
' =============================================================================================

' Fun��o desenvolvida para fazer o tratamento do
' formul�rio de dados.
Private Function ValidarForm()
	
	' Tratamento de campos do formul�rio. =============================
	
	If Trim(vstr_DsAtividade) = "" Then
		Call AddErro("Descri��o", "Favor, preencher o campo da Descri��o.")
	End If
	
	' Verifica se algum tipo de erro
	' ocorreu na valida��o do formul�rio.
	If TotalErros > 0 Then
		
		' Formul�rio inv�lido.
		ValidarForm = False
	Else
		
		' Formul�rio v�lido.
		ValidarForm = True
	End If
End Function

%>

<!-- #include file = "../includes/CloseConnection.asp" -->