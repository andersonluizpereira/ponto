<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- #include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->

<%
' Declaração de variáveis locais. ==============================================

' Variável que indica o fluxo de funcionamento da página;
Dim vstr_Executar


' Conseguindo os valores submetidos
' para está página.
vstr_Executar			= Request("hdnExecutar")

' ============================================================
' Inicio do fluxo de funcionamento da Página.
' ============================================================


' Analizando a variável que indica o fluxo de 
' operação desta página.
Select Case vstr_Executar
	
	
	Case "LOGAR"											' -->> Operação de Logar o Usuário Na aplicação.
		
		
		' ... operação de logar o usuário.
		
		
		' Declaração de variáveis auxiliares
		Dim vobj_command
		Dim vobj_rs
		
		
		' Declaração dos parametros submetidos a está página.
		Dim vstr_CdLogin
		Dim vstr_CdSenha
		
		
		
		' Recuperando os parametros do formulário.
		vstr_CdLogin		= Request.Form("txtCdLogin")
		vstr_CdSenha		= Request.Form("txtCdSenha")
		
		
		
		
		'//Response.Write(vstr_CdLogin)
		'//Response.Write(vstr_CdSenha)
		'//Response.Write("xxxxxxxxxxxxxxxxxxxxxxxxx")
		'//Response.End 
		
		
				
		
		' ---------------------------------------------------------------------
		' Verificando se o Usuário Informado existe no banco de dados.
		' ---------------------------------------------------------------------
		Set vobj_command = Server.CreateObject("ADODB.Command")
		Set vobj_command.ActiveConnection = vobj_conexao
			
			
		vobj_command.CommandType					= adCmdStoredProc
		vobj_command.CommandText					= "consultaLogin"
		vobj_command.Parameters.Refresh
		vobj_command.CommandTimeout					= 0
		
		vobj_command.Parameters.Append vobj_command.CreateParameter("param1",adChar, adParamInput, 10, vstr_CdLogin)
		vobj_command.Parameters.Append vobj_command.CreateParameter("param1",adChar, adParamInput, 15, EncriptaString(vstr_CdSenha))
		
		Set vobj_rs = vobj_command.Execute()
		' ---------------------------------------------------------------------
		
		
		' Verificando se o usuário foi encontrado no
		' banco de dados.
		If Not vobj_rs.EOF Then
			
			
			' Armazenando os parametros do usuário
			' na sessão. Isto garantirá o acesso do usuário no sistema.
			
			Session("sstr_IdUsuario")		= vobj_rs("ID_USUARIO")
			Session("sstr_DsUsuario")		= vobj_rs("DS_USUARIO")
			Session("sint_TipoUsuario")		= vobj_rs("FL_PERFIL")
			
			
			' Encerrando a transação (operação) com
			' o recordset e com o objeto command.
			vobj_rs.Close
			Set vobj_rs = Nothing
			
			
			
			' Se caso o usuário logado for um usuário do tipo pátio, então
			' devemos verificar se o usuário é do tipo administrador.
			If Session("sint_TipoUsuario") = "1" Then
				
				Session("sboo_fladministrador")	= True
				
			Else
				
				' ... o usuário é um usuário comum.
				Session("sboo_fladministrador")	= False
			End If
			
			
			
			' Chamando procedimento que fecha a conexão com o banco
			' de dados.
			Call FecharConexao()
			
			
			
			' Redirecionando para a página home da aplicação.
			Response.Redirect getBaseLink("/horas/horaslancamento.asp")
		Else
			
			' Se caso o usuário não for encontrado, então adiciona uma mensagem
			' de erro na tela.
			Call AddErro("MsgErro", "Usuário Não Encontrado !!!")
			
		End If
		
		
		' Encerrando a transação (operação) com
		' o recordset e com o objeto command.
		vobj_rs.Close
		Set vobj_rs = Nothing
		Set vobj_command = Nothing
		
End Select
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Stefanini - Controle de Horas</TITLE>
<LINK rel="stylesheet" type="text/css" href="<%=getBaseLink("/css/chs.css")%>">
<SCRIPT LANGUAGE=javascript>
	// Procedimento desenvolvido para verificar se o Usuário
	// Digitou corretamento seu login e senha.
	function validaLogin()
	{
		// Verificando se login e senha estão preenchidos.
		
		
		
		//if ((document.thisForm.txtCdLogin.value == '') || (document.thisForm.txtCdSenha.value == '')){
			
		//	window.alert("Digite seu Login e Senha para entrar no sistema.");
		//	document.thisForm.txtCdLogin.focus();
		//	return false;
		//}
		//else
		//{
			
			// ... login e senha preenchidos.
			
			
			document.thisForm.hdnExecutar.value = "LOGAR";
			return true;
		//}
	}
</SCRIPT>
</HEAD>

<BODY onload="document.thisForm.txtCdLogin.focus()">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr height="55" vAlign="top" class="top">
		<td align="right"><!-- include file="../inc/topo.asp" --></td>
	</tr>
	<tr height="27">
		<td class="MenuBackground">
		</td>
	</tr>
	<tr>
		<td align="middle">
		<br /><br /><br /><br /><br />
			<form name="thisForm" action="login.asp" method="post" onSubmit="return validaLogin();">
				
				<input type="hidden" name="hdnprocessar" value="S">
				<input type="hidden" name="hdnExecutar">
				<fieldset style="LEFT: 0px; WIDTH: 590px; HEIGHT: 94px">
					<legend>
					   <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b> Login </b></font>
					</legend>
				<TABLE cellSpacing=1 cellPadding=1 align=center border=0>
				  <TR>
				    <TD>
				      <TABLE cellSpacing=1 cellPadding=1 align=center border=0>
				        <TR>
				          <TD class=font><B>Usuário:</B> </TD>
				          <TD><INPUT class=TextBox name="txtCdLogin" maxlength="15"></TD></TR>
				        <TR>
				          <TD class=font><B>Senha: </B></TD>
				          <TD><INPUT class=TextBox type=password maxlength="15" Name="txtCdSenha"></TD>
				        </TR>
				        <tr>
							<td class="font" colspan="2"><%=ExibirErros()%></td>
						</tr>
				      </TABLE>
				    </TD>
				  </TR>
				  <TR>
				    <TD>&nbsp;<INPUT class=TextBox type=hidden name=txtXML></TD>
				  </TR>
				  <TR>
				    <TD>
				      <TABLE cellSpacing=1 cellPadding=1 align=center border=0>
				        <TR>
				          <TD><INPUT class="BotaoOff" onmouseover="this.className='BotaoOn'" onmouseout="this.className='BotaoOff'" type="submit" value="Ok" name="cmdEnviar"></TD>
				          <TD><INPUT class="BotaoOff" onmouseover="this.className='BotaoOn'" onmouseout="this.className='BotaoOff'" type="button" value="Sair" onclick="window.close()"></TD>
				        </TR>
					  </TABLE>
					</TD>
				  </TR>
				</TABLE>
			</FORM>
		</td>
	</tr>
</TABLE>
</BODY>
</HTML>
