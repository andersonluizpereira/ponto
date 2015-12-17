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
Dim vint_FlAtivo
Dim vstr_IdUsuario
Dim vstr_IdUsuarioProjeto

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


vstr_IdProjeto				= Request.Form("cmbComboProjeto")

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
	'vstr_IdProjeto				= Request.Form("hdnIdRegistro")
	
	
Else
	
	
	' Verifica se a operação a ser executada nesta página é
	' a operação de inclusão e verifica se a página não foi
	' processada ainda.
	If vstr_Operacao = "I" And vstr_Processar <> "S" Then
		
		' Neste caso todas as variáveis devem ser vazias
		' para o usuário poder preencher seu novo cadastro
		' do registro.
		
		vstr_IdProjeto			= Empty
		vint_FlAtivo			= Empty
		vstr_IdUsuario			= Empty
		vstr_IdUsuarioProjeto	= Empty
	Else
		
		' ... está opção acontecerá quando o usuário processar
		' a página, neste caso todas os dados da tela serão
		' submetidos e devem ser pegos neste lugar.
		
		vint_FlAtivo			= Request.Form("txtFlAtivo")
		vstr_IdUsuario			= Request.Form("cmbComboUsuario")
		vstr_IdUsuarioProjeto	= Request.Form("cmbComboUsuarioProjeto")
		
	End If
End If


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
	Select Case vstr_Executar
		
		'Operação de associar o usuario ao projeto
		Case "ASSOCIAR"
			
			' ... processamento de alteração do registro.
			
			vstr_IdProjeto				= Request.Form("hdnIdRegistro")
			
			'Validando Item Projeto a ser associado
			If Not Trim(vstr_IdProjeto) = "" Then
				
				If Not Trim(vstr_IdUsuario) = "" Then
					
					' ---------------------------------------------------------------------
					' Alterando os dados do registro no banco de dados.
					' ---------------------------------------------------------------------
					Set vobj_commandProc = Server.CreateObject("ADODB.Command")
					Set vobj_commandProc.ActiveConnection = vobj_conexao
					
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "associaUsuarioProjeto"
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_IdProjeto))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 10, Trim(vstr_IdUsuario))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adBoolean, adParamInput,, True)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adDate, adParamInput, 10, converterDataParaSQL(Date()))
					
					
					Call vobj_commandProc.Execute
					Set vobj_commandProc = Nothing
					' ---------------------------------------------------------------------
					
					vstr_IdUsuarioProjeto = vstr_IdUsuario
					
				End If
			End If
			
		
		'Operação de desvincular o usuario do projeto
		Case "REMOVER"
			
			
			' ... processamento de alteração do registro.
			
			vstr_IdProjeto				= Request.Form("hdnIdRegistro")
			
			'Validando Item Projeto a ser associado
			If Not Trim(vstr_IdProjeto) = "" Then
				
				If Not Trim(vstr_IdUsuarioProjeto) = "" Then
					
					' ---------------------------------------------------------------------
					' Alterando os dados do registro no banco de dados.
					' ---------------------------------------------------------------------
					Set vobj_commandProc = Server.CreateObject("ADODB.Command")
					Set vobj_commandProc.ActiveConnection = vobj_conexao
					
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "removeUsuarioProjeto"
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_IdProjeto))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 10, Trim(vstr_IdUsuarioProjeto))
					
						
					Call vobj_commandProc.Execute
					Set vobj_commandProc = Nothing
					' ---------------------------------------------------------------------
					
				End If
			End If

			'Operação de desvincular o usuario do projeto
		Case "MUDAR_ATIVO"
			
			
			' ... processamento de alteração do registro.
			
			vstr_IdProjeto				= Request.Form("hdnIdRegistro")
			
			'Validando Item Projeto a ser associado
			If Not Trim(vstr_IdProjeto) = "" Then
				
				If Not Trim(vstr_IdUsuarioProjeto) = "" Then
					
					' ---------------------------------------------------------------------
					' Alterando os dados do registro no banco de dados.
					' ---------------------------------------------------------------------
					Set vobj_commandProc = Server.CreateObject("ADODB.Command")
					Set vobj_commandProc.ActiveConnection = vobj_conexao
					
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "alteraUsuarioProjetoAtivo"
					
					
					If vint_FlAtivo = "" Then
						
						vint_FlAtivo = 0
					Else
						If vint_FlAtivo = "on" Or vint_FlAtivo = True Then
							
							vint_FlAtivo = 1
						Else
							vint_FlAtivo = 0
						End If
					End If
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_IdProjeto))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 10, Trim(vstr_IdUsuarioProjeto))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adBoolean, adParamInput,, vint_FlAtivo)
					
					
					Call vobj_commandProc.Execute
					Set vobj_commandProc = Nothing
					' ---------------------------------------------------------------------
					
				End If
			End If		
	End Select
	
End If
' *******************************************************
' FINAL DA ROTINA QUE FAZ O PROCESSAMENTO DOS DADOS
' DO REGISTRO.
' *******************************************************

%>

<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/usuariosprojetos.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<i><b class="TituloPagina">Projetos</b></i>
			<form name="thisForm" action="usuariosprojetos.asp" method="post">
				
				<input type="hidden" name="hdnProcessar" value="S" />
				<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>" />
				<input type="hidden" name="hdnIdRegistro" value="<%=vstr_IdProjeto%>" />
				<input type="hidden" name="hdnExecutar" />
				
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<table border="0" class="font" cellpadding="0" cellspacing="0">
							<tr>
								<td>Projeto:&nbsp;</td>
								<td><%Call CriarComboProjeto("cmbComboProjeto", vstr_IdProjeto,"atualizarComboProjeto();",Empty) %></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td>&nbsp;
						</td>
					</tr>
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 595px; HEIGHT: 220px">
							<legend>
							   <b>Usuários relacionados ao Projeto</b>
							</legend>
							<TABLE class="font" BORDER=0 CELLSPACING=5 CELLPADDING=1>
								<TR>
									<TD>Colaboradores</TD>
									<TD>&nbsp;</TD>
									<TD>Associados</TD>
								</TR>
								<TR>
									<TD valign="top">
										<%
										
										If vstr_IdProjeto = "" Then
											
											%>
											
											<SELECT class="TextBox" name="cmbComboUsuario" multiple style="WIDTH: 239px; HEIGHT: 142px"></SELECT>
											
											<%
										
										Else
											
											Call CriarComboUsuario("cmbComboUsuario", vstr_IdUsuario, vstr_IdProjeto,Empty)
										
										End If
										
										%>
									</TD>
									<TD align="middle">
										<INPUT onclick="Associar();" type="button" name="cmdAssociar" value="Adicionar" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" style="WIDTH: 66px; HEIGHT: 19px" size=13><br><br>
										<INPUT onclick="Remover();" type="button" name="cmdRemover" value="Remover" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" style="WIDTH: 66px; HEIGHT: 19px" size=13>
									</TD>
									<TD>
										
										<%
										
										If vstr_IdProjeto = "" Then
											
											%>
											
											<SELECT class="TextBox" name="cmbComboUsuarioProjeto" multiple style="WIDTH: 239px; HEIGHT: 142px"></SELECT>
											
											<%
										
										Else
											
											Call CriarComboUsuarioProjeto("cmbComboUsuarioProjeto", vstr_IdUsuarioProjeto, "atualizarComboAssociados();",vstr_IdProjeto)
										
										End If
										
										%>
									</TD>
								</TR>
								<TR>
									<TD>&nbsp;</TD>
									<TD>&nbsp;</TD>
									<TD valign="top"><input type="checkbox" name="txtFlAtivo" onclick="AtivoOnOff()" <%=VerificarUsuarioProjetoAtivo(vstr_IdProjeto, vstr_IdUsuarioProjeto)%>>&nbsp;Ativo</TD>
								</TR>
							</TABLE>
						</fieldset>
						</td>
					</tr>
					<tr>
						<td align="middle">
							<table ALIGN="center" BORDER="0" CELLSPACING="1" CELLPADDING="1">
								<tr>
									<td><input type="button" name="cmdRetornar" value="Retornar" onClick="voltar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Retornar ao Início"></td>
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


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
								' COMBO DE PROJETOS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub CriarComboProjeto(pstr_Nome, pstr_ValorDefault, pstr_onChange, pstr_Evento)
	
	
	' Declaração de variáveis locais.
	Dim vobj_command
	Dim vobj_rs
	
	
	' ---------------------------------------------------------------------
	' Selecionando todos os Registros
	' ---------------------------------------------------------------------
	Set vobj_command = Server.CreateObject("ADODB.Command")
	Set vobj_command.ActiveConnection = vobj_conexao
	
	
	vobj_command.CommandType				= adCmdStoredProc
	vobj_command.CommandText				= "consultaProjetosAtivos"
	vobj_command.Parameters.Refresh
	
	
	Set vobj_rs = vobj_command.Execute
	' ---------------------------------------------------------------------
	
	
	%>
	<SELECT name="<%=pstr_Nome%>" onChange="<%=pstr_onChange%>" <%=pstr_Evento%> class="TextBox">
		<option value="<%=Empty%>">Selecione</option>
		<%
		
		
		If Not vobj_rs.EOF Then
			
			' Loop de todos os registros encontrados.
			Do While Not vobj_rs.EOF
				
				
				' Verificando se o registro
				' é o Registro default que deve ser selecionado.
				If Trim(pstr_ValorDefault) = Trim(vobj_rs("ID_PROJETO")) Then
					%><OPTION selected value="<%=vobj_rs("ID_PROJETO")%>"><%=vobj_rs("ID_PROJETO") & " - " & vobj_rs("DS_PROJETO")%></OPTION><%
				Else
					%><OPTION value="<%=vobj_rs("ID_PROJETO")%>"><%=vobj_rs("ID_PROJETO") & " - " & vobj_rs("DS_PROJETO")%></OPTION><%
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


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
								' COMBO DE USUARIO
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub CriarComboUsuario(pstr_Nome, pstr_ValorDefault, pstr_ValorAux, pstr_Evento)
	
	
	' Declaração de variáveis locais.
	Dim vobj_command
	Dim vobj_rs
	
	
	' ---------------------------------------------------------------------
	' Selecionando todos os Registros
	' ---------------------------------------------------------------------
	Set vobj_command = Server.CreateObject("ADODB.Command")
	Set vobj_command.ActiveConnection = vobj_conexao
	
	
	vobj_command.CommandType				= adCmdStoredProc
	vobj_command.CommandText				= "consultaUsuarioNaoAssociadoProjeto"
	vobj_command.Parameters.Refresh
	
	vobj_command.Parameters.Append vobj_command.CreateParameter("param1", adChar, adParamInput, 10, pstr_ValorAux)
	
	Set vobj_rs = vobj_command.Execute
	' ---------------------------------------------------------------------
	
	%>
	<SELECT name="<%=pstr_Nome%>" <%=pstr_Evento%> class="TextBox" multiple style="WIDTH: 239px; HEIGHT: 142px">
		<!--<option value="<%=Empty%>">Selecione</option>-->
		<%
		
		
		If Not vobj_rs.EOF Then
			
			' Loop de todos os registros encontrados.
			Do While Not vobj_rs.EOF
				
				
				' Verificando se o registro
				' é o Registro default que deve ser selecionado.
				If Trim(pstr_ValorDefault) = Trim(vobj_rs("ID_USUARIO")) Then
					%><OPTION selected value="<%=vobj_rs("ID_USUARIO")%>"><%=vobj_rs("DS_USUARIO")%></OPTION><%
				Else
					%><OPTION value="<%=vobj_rs("ID_USUARIO")%>"><%=vobj_rs("DS_USUARIO")%></OPTION><%
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

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
								' COMBO DE USUARIO PROJETO
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub CriarComboUsuarioProjeto(pstr_Nome, pstr_ValorDefault, pstr_onChange, pstr_ValorAux)
	
	
	' Declaração de variáveis locais.
	Dim vobj_command
	Dim vobj_rs
	
	
	' ---------------------------------------------------------------------
	' Selecionando todos os Registros
	' ---------------------------------------------------------------------
	Set vobj_command = Server.CreateObject("ADODB.Command")
	Set vobj_command.ActiveConnection = vobj_conexao
	
	
	vobj_command.CommandType				= adCmdStoredProc
	vobj_command.CommandText				= "consultaUsuarioAssociadoProjeto"
	vobj_command.Parameters.Refresh
	
	vobj_command.Parameters.Append vobj_command.CreateParameter("param1", adChar, adParamInput, 10, pstr_ValorAux)
	
	Set vobj_rs = vobj_command.Execute
	' ---------------------------------------------------------------------
	
	%>
	<SELECT name="<%=pstr_Nome%>" onChange="<%=pstr_onChange%>" class="TextBox" multiple style="WIDTH: 239px; HEIGHT: 142px">
		<!--<option value="<%=Empty%>">Selecione</option>-->
		<%
		
		
		If Not vobj_rs.EOF Then
			
			' Loop de todos os registros encontrados.
			Do While Not vobj_rs.EOF
				
				
				' Verificando se o registro
				' é o Registro default que deve ser selecionado.
				If Trim(pstr_ValorDefault) = Trim(vobj_rs("ID_USUARIO")) Then
					%><OPTION selected value="<%=vobj_rs("ID_USUARIO")%>"><%=vobj_rs("DS_USUARIO")%></OPTION><%
				Else
					%><OPTION value="<%=vobj_rs("ID_USUARIO")%>"><%=vobj_rs("DS_USUARIO")%></OPTION><%
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

'Verifica se o usuario associado ao projeto esta ativo no banco
Private Function VerificarUsuarioProjetoAtivo(pstr_IdProjeto, pstr_IdUsuario)
	
	If Not Trim(pstr_IdProjeto) = "" Then
		
		If Not Trim(pstr_IdUsuario) = "" Then
			
			' Declaração de variáveis locais.
			Dim vobj_command
			Dim vobj_rs
			
			
			' ---------------------------------------------------------------------
			' Selecionando todos os Registros
			' ---------------------------------------------------------------------
			Set vobj_command = Server.CreateObject("ADODB.Command")
			Set vobj_command.ActiveConnection = vobj_conexao
			
			
			vobj_command.CommandType				= adCmdStoredProc
			vobj_command.CommandText				= "consultaUsuarioProjetoAtivo"
			vobj_command.Parameters.Refresh
			
			vobj_command.Parameters.Append vobj_command.CreateParameter("param1", adChar, adParamInput, 10, pstr_IdProjeto)
			vobj_command.Parameters.Append vobj_command.CreateParameter("param2", adChar, adParamInput, 10, pstr_IdUsuario)
			
			Set vobj_rs = vobj_command.Execute
			' ---------------------------------------------------------------------
			
			If Not vobj_rs.EOF Then
				
				If vobj_rs("FL_ATIVO") = "on" Or vobj_rs("FL_ATIVO") = True Then
					
					VerificarUsuarioProjetoAtivo = "Checked=""checked"""
				End If
				
			Else
				
				VerificarUsuarioProjetoAtivo = ""
				
			End If
			
			
			vobj_rs.Close
			Set vobj_rs = Nothing
			Set vobj_command = Nothing
		Else
			
			VerificarUsuarioProjetoAtivo = ""
			
		End If
	Else
		
		VerificarUsuarioProjetoAtivo = ""
		
	End If
	
	
End Function

%>

<!-- #include file = "../includes/CloseConnection.asp" -->