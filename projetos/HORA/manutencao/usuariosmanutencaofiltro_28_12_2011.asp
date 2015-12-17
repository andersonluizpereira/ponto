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

' Declaração de variáveis utilizadas para armazenar os
' valores dos campos da tela.
Dim vstr_IdUsuario
Dim vint_IdMes
Dim vstr_DsAno

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

' Veririfa se a operação a ser executada nesta página é a 
' operação de Alteração ou Visualização e se a página não
' foi processada ainda.
If (vstr_Operacao = "A" or vstr_Operacao = "V") And vstr_Processar <> "S" Then
	
	
	' ... neste caso deve ser solicitado o código do registro
	' e encontrar suas informações no banco de dados para exibir para
	' as informações do registro na tela.
	' Conseguindo o código do registro.
	vstr_IdUsuario				= Request.Form("hdnIdRegistro")
	
	
Else
	
	
	' Verifica se a operação a ser executada nesta página é
	' a operação de inclusão e verifica se a página não foi
	' processada ainda.
	If vstr_Operacao = "I" And vstr_Processar <> "S" Then
		
		' Neste caso todas as variáveis devem ser vazias
		' para o usuário poder preencher seu novo cadastro
		' do registro.
		
		
		vstr_IdUsuario 			= Empty
		vint_IdMes				= Empty
		vstr_DsAno				= Empty
		
	Else
		
		' ... está opção acontecerá quando o usuário processar
		' a página, neste caso todas os dados da tela serão
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
			<form name="thisForm" action="Consultanew.asp" method="post">
				
				<input type="hidden" name="hdnProcessar" value="S" />
				<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>" />
				<input type="hidden" name="hdnIdRegistro" value="<%=vstr_IdUsuario%>" />
				<input type="hidden" name="hdnExecutar" />
				
				<i><b class="TituloPagina">Cosultar Horas diárias</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 595px; HEIGHT: 50px">
							<legend>
							   <b>Filtro Relatório</b>
							</legend>
							<TABLE valign="center" class="font" BORDER=0 CELLSPACING=1 CELLPADDING=1>
								<tr>
									<td>Dia:</td>
									<td><input type="text" maxlength="10" size="20" name="Data" id="Data">   </td>
									
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
 									
									<td><input type="Reset" value="Limpar" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Limpar campo!"></td>
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
	
	
	' Declaração de variáveis locais.
	Dim vobj_command
	Dim vobj_rs
	
	
	' ---------------------------------------------------------------------
	' Selecionando todos os Registros
	' ---------------------------------------------------------------------
	Set vobj_command = Server.CreateObject("ADODB.Command")
	Set vobj_command.ActiveConnection = vobj_conexao
	
	
	vobj_command.CommandType				= adCmdStoredProc
	vobj_command.CommandText				= "consultaUsuarioAtivo"
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

%>

<!-- #include file = "../includes/CloseConnection.asp" -->