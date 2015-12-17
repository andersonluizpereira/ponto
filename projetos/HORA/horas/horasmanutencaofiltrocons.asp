<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- #include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->

<%

If Not Session("sboo_flconsulta") = True Then
	
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

' Declara��o de vari�veis utilizadas para armazenar os
' valores dos campos da tela.
Dim vstr_IdUsuario
Dim vint_IdMes
Dim vstr_DsAno

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

' Veririfa se a opera��o a ser executada nesta p�gina � a 
' opera��o de Altera��o ou Visualiza��o e se a p�gina n�o
' foi processada ainda.
If (vstr_Operacao = "A" or vstr_Operacao = "V") And vstr_Processar <> "S" Then
	
	
	' ... neste caso deve ser solicitado o c�digo do registro
	' e encontrar suas informa��es no banco de dados para exibir para
	' as informa��es do registro na tela.
	' Conseguindo o c�digo do registro.
	vstr_IdUsuario				= Request.Form("hdnIdRegistro")
	
	
Else
	
	
	' Verifica se a opera��o a ser executada nesta p�gina �
	' a opera��o de inclus�o e verifica se a p�gina n�o foi
	' processada ainda.
	If vstr_Operacao = "I" And vstr_Processar <> "S" Then
		
		' Neste caso todas as vari�veis devem ser vazias
		' para o usu�rio poder preencher seu novo cadastro
		' do registro.
		
		
		vstr_IdUsuario 			= Empty
		vint_IdMes				= Empty
		vstr_DsAno				= Empty
		
	Else
		
		' ... est� op��o acontecer� quando o usu�rio processar
		' a p�gina, neste caso todas os dados da tela ser�o
		' submetidos e devem ser pegos neste lugar.
		
		vstr_IdUsuario			= Request.Form("cmbComboUsuario")
		vint_IdMes				= Request.Form("cmbComboMes")
		vstr_DsAno				= Request.Form("txtDsAno")
		
	End If
End If

%>

<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/horasmanutencaofiltro.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<form name="thisForm" action="horasmanutencaocons.asp" method="post">
				
				<input type="hidden" name="hdnProcessar" value="S" />
				<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>" />
				<input type="hidden" name="hdnIdRegistro" value="<%=vstr_IdUsuario%>" />
				<input type="hidden" name="hdnExecutar" />
				
				<i><b class="TituloPagina">Filtro para Manuten��o de Horas</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 595px; HEIGHT: 50px">
							<legend>
							   <b>Filtro Relat�rio - M�dulo consulta</b>
							</legend>
							<TABLE valign="center" class="font" BORDER=0 CELLSPACING=1 CELLPADDING=1>
								<tr>
									<td>Usu�rio</td>
									<td><%Call CriarComboUsuario("cmbComboUsuario", vstr_IdUsuario, Empty,Empty)%></td>
									<td>M�s</td>
									<td>
										<select name="cmbComboMes" id="Data" class="TextBox">
											<option value="" selected>Selecione</option>
											<%
											
											Dim vvar_ArrayMes(11)
											
											vvar_ArrayMes(0) = "Janeiro"
											vvar_ArrayMes(1) = "Fevereiro"
											vvar_ArrayMes(2) = "Mar�o"
											vvar_ArrayMes(3) = "Abril"
											vvar_ArrayMes(4) = "Maio"
											vvar_ArrayMes(5) = "Junho"
											vvar_ArrayMes(6) = "Julho"
											vvar_ArrayMes(7) = "Agosto"
											vvar_ArrayMes(8) = "Setembro"
											vvar_ArrayMes(9) = "Outubro"
											vvar_ArrayMes(10) = "Novembro"
											vvar_ArrayMes(11) = "Dezembro"
											
											Dim vint_Contator
											
											For vint_Contator = 1 To 12
												
												If Month(Date) = vint_Contator Then
													
													%><option value="<%=vint_Contator%>" selected><%=vvar_ArrayMes(vint_Contator - 1)%></option><%
													
												Else
													
													%><option value="<%=vint_Contator%>"><%=vvar_ArrayMes(vint_Contator - 1)%></option><%
													
												End If
											Next
											
											%>
											
										</select>
									</td>
									<td>Ano</td>
									<td>
										<input class="TextBox" type="text" maxlength="4" size="5" name="txtDsAno" Value="<%=Year(Date())%>">
									</td>
								</tr>
							</TABLE>
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
									<td><input type="button" value="Enviar" onclick="enviar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Lista horas do m�s selecionado"></td>
									<td><input type="button" value="Retornar" onclick="voltar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Retorna a p�gina anterior"></td>
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
								' COMBO DE USUARIO
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub CriarComboUsuario(pstr_Nome, pstr_ValorDefault, pstr_ValorAux, pstr_Evento)
	
	
	' Declara��o de vari�veis locais.
	Dim vobj_command
	Dim vobj_rs
	
	
	' ---------------------------------------------------------------------
	' Selecionando todos os Registros
	' ---------------------------------------------------------------------
	Set vobj_command = Server.CreateObject("ADODB.Command")
	Set vobj_command.ActiveConnection = vobj_conexao
	
	
	vobj_command.CommandType				= adCmdStoredProc
	

     If Session("sint_TipoUsuario") = "3" Then

	
	vobj_command.CommandText				= "consultaUsuarioAtivovdi"
	
	Else
	
	vobj_command.CommandText				= "consultaUsuarioAtivo"
	
	End If
	
	vobj_command.Parameters.Refresh
	
	Set vobj_rs = vobj_command.Execute
	' ---------------------------------------------------------------------
	
	%>
	<SELECT name="<%=pstr_Nome%>" <%=pstr_Evento%> class="TextBox">
		<option value="<%=Empty%>">Selecione</option>
		<%
		
		
		If Not vobj_rs.EOF Then
			
			' Loop de todos os registros encontrados.
			Do While Not vobj_rs.EOF
				
				
				' Verificando se o registro
				' � o Registro default que deve ser selecionado.
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

%>

<!-- #include file = "../includes/CloseConnection.asp" -->