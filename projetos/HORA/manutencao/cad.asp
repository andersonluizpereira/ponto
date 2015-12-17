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


Dim vstr_CdUsuario
Dim vstr_DsUsuario
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

Dim vobj_rs
Dim vobj_commandProc
					
						
	    vstr_IdUsuario			= Empty
		vstr_CdUsuario			= Empty
		vstr_DsUsuario			= Empty
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
		
		

	   vstr_CdUsuario			= Request.Form("txtCdUsuario")
		vstr_DsUsuario			= Request.Form("txtDsUsuario")
		vstr_DsCPF				= Request.Form("txtDsCPF")
		vstr_DsRG				= Request.Form("txtDsRG")
		vint_IdFuncao			= Request.Form("cmbComboFuncao")
		vint_IdEquipe			= Request.Form("cmbComboEquipe")
		vint_FlPerfil			= Request.Form("cmbComboPerfil")
		'vint_FlAtivo			= Request.Form("txtFlAtivo")
		vstr_CdSenha			= Request.Form("txtCdSenha")
		vstr_CdConfirmaSenha	= Request.Form("txtCdConfirmaSenha")
		
		vstr_DtNascimento		= Request.Form("txtDtNascimento")
		vstr_DsTelefone			= Request.Form("txtDsTelefone")
		vstr_DsRamal			= Request.Form("txtDsRamal")
		vstr_DsLocalAlocado		= Request.Form("txtDsLocalAloc")
		
						
				
				
					
					' ---------------------------------------------------------------------
					' Incluindo os dados do registro no banco de dados.
					' ---------------------------------------------------------------------
					Set vobj_commandProc = Server.CreateObject("ADODB.Command")
					Set vobj_commandProc.ActiveConnection = vobj_conexao
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "incluiUsuario"
					
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_CdUsuario))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 100, Trim(vstr_DsUsuario))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adChar, adParamInput, 11, Trim(vstr_DsCPF))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adChar, adParamInput, 15, Trim(vstr_DsRG))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adInteger, adParamInput,, vint_IdFuncao)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adChar, adParamInput, 25, vint_IdEquipe)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param7",adInteger, adParamInput,, vint_FlPerfil)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param8",adChar, adParamInput, 15, EncriptaString(Trim(vstr_CdSenha)))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param9",adDate, adParamInput,, converterDataParaSQL(Date()))
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param10",adChar, adParamInput, 10, Trim(vstr_DtNascimento))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param11",adChar, adParamInput, 20, Trim(vstr_DsTelefone))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param12",adChar, adParamInput, 15, Trim(vstr_DsRamal))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param13",adChar, adParamInput, 30, Trim(vstr_DsLocalAlocado))
					
					If Not Trim(vstr_DtNascimento) = "" Then
						
						vstr_DtAniversario = DateSerial(2000, Month(vstr_DtNascimento), Day(vstr_DtNascimento))
						
					End If
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param13",adChar, adParamInput, 10, converterDataParaSQL(vstr_DtAniversario))
					
					vobj_commandProc.Execute
					
					
					Set vobj_commandProc = Nothing
 %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Cadastrar Usuários</title>
</head>
<body>
   <!-- #include file = "../includes/LayoutBegin.asp" -->
  


<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<form name="thisForm" action="cad.asp" method="post">
				
				
				
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
									<td align="left" colspan="3"><input name="txtCdUsuario" id="User_ID" class="TextBox" size="15" maxlength="10" ></td>
									<td align="left">Nome:&nbsp;</td>
									<td colspan="3" align="left"><input name="txtDsUsuario" id="Nome" class="TextBox" size="35" maxlength="100" ></td>
									
								</tr>
								<tr>
									<td align="left">Data Nascimento:&nbsp;</td>
									<td align="left">
										<input type="text" name="txtDtNascimento" class="TextBox" size="15" maxlength="10">
									</td>
								</tr>
								<tr>
									<td align="left">CPF:&nbsp;</td>
									<td align="left" colspan="3"><input name="txtDsCPF" id="CPF" class="TextBox" size="15" maxlength="11"></td>
									<td align="left">RG:&nbsp;</td>
									<td colspan="3" align="left"><input name="txtDsRG" id="RG" class="TextBox" maxlength="15"></td>
								</tr>
								<tr>
								</tr>
								<tr>
									<td align="left">Função:&nbsp;</td>
									
									<td align="left" colspan="3"></td>
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
									
									
								   <td> <select name="cmbComboEquipe" class="TextBox">
									<option selected value="">Selecione</option>
									<option value="Azul">Azul</option>
												<option value="Laranja">Laranja</option>								
												<option value="Vermelha">Vermelha</option>
												<option value="Roxa">Roxa</option>
									</select>
									</td>
									
									
									<td align="left">Telefone:&nbsp;</td>
									<td align="left">
										<input type="text" name="txtDsTelefone" class="TextBox" size="17" maxlength="20">
									</td>
									<td nowrap align="left">Ramal:&nbsp;</td>
									<td align="left">
										<input type="text" name="txtDsRamal" class="TextBox" size="7" maxlength="15">
									</td>
									<td nowrap align="left">Local Alocado:&nbsp;</td>
									<td align="left">
										<input type="text" name="txtDsLocalAlocado" class="TextBox" size="25" maxlength="30">
									</td>
								</tr>
								<tr>
									<td align="left">Senha:&nbsp;</td>
									<td align="left" colspan="3">
										<input type="password" name="txtCdSenha" id="txtCdSenha" class="TextBox" size="25" maxlength="15">
									</td>
									<td nowrap align="left">Confirma Senha:&nbsp;</td>
									<td align="left">
										<input type="password" name="txtCdConfirmaSenha" id="txtCdConfirmaSenha" class="TextBox" size="25" maxlength="15">
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

  
</body>
</html>
<!-- #include file = "../includes/CloseConnection.asp" -->