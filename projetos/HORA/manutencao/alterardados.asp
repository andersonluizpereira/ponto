<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- S#include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->

<%

' Declaração de variáveis locais. ==============================================

' Guarda a operação que será executa nesta tela.
' Obs.: Seus valores podem ser A = Alteração, I = Inclusão, V = Visualização.
Dim vstr_Operacao

' Variável flag que indica se a página deve ser 
' processada, apenas disponivel para as operações de
' A e I.
Dim vstr_Processar


Dim vstr_CdSenha
Dim vstr_CdNovaSenha
Dim vstr_CdConfirmaSenha

' para está página.
vstr_Operacao		= Request.Form("pstr_Operacao")
vstr_Processar		= Request.Form("hdnProcessar")


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
			
			vstr_CdSenha			= Request.Form("txtCdSenha")
			vstr_CdNovaSenha		= Request.Form("txtCdNovaSenha")
			vstr_CdConfirmaSenha	= Request.Form("txtCdConfirmaSenha")
			
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
				
				
				' ---------------------------------------------------------------------
				' Alterando os dados do registro no banco de dados.
				' ---------------------------------------------------------------------
				Set vobj_commandProc = Server.CreateObject("ADODB.Command")
				Set vobj_commandProc.ActiveConnection = vobj_conexao
				
				
				vobj_commandProc.CommandType					= adCmdStoredProc
				vobj_commandProc.CommandText					= "alteraSenha"
				
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Session("sstr_IdUsuario"))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 15, EncriptaString(vstr_CdNovaSenha))
				
				
				Call vobj_commandProc.Execute
				Set vobj_commandProc = Nothing
				' ---------------------------------------------------------------------
				
				
			End If
			
	End Select
End If
' *******************************************************
' FINAL DA ROTINA QUE FAZ O PROCESSAMENTO DOS DADOS
' DO REGISTRO.
' *******************************************************

%>


<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/alterardados.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td align="middle">
			<form name="thisForm" action="alterardados.asp" method="post">
				
				<input type="hidden" name="hdnProcessar" value="S">
				<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>">
				
				<i><b class="TituloPagina">Alterar Dados</b></i>
				<TABLE cellSpacing=1 cellPadding=1 align=center border=0>
					<tr>
						<td height="20"></td>
					</tr>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1 align=center border=0>
								<TR>
									<TD class=font>
										<strong>Senha:</strong>
									</TD>
									<TD>
										<input class=TextBox type=password name="txtCdSenha" maxlength="15" />
									</TD>
								</TR>
								<TR>
									<TD class=font>
										<strong>Nova Senha: </strong>
									</TD>
									<TD>
										<input class=TextBox type=password name="txtCdNovaSenha" maxlength="15" />
									</TD>
								</TR>
									<TD class=font>
										<strong>Confirma Nova Senha:</strong>
									</TD>
									<TD>
										<input class=TextBox type=password name="txtCdConfirmaSenha" maxlength="15" />
									</TD>
								</TR>
								<tr>
									<td colspan="2" class="font" align="center"><%=ExibirErros()%></td>
								</tr>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1 align=center border=0>
								`<tr>
									<td height="5"></td>
								</tr>
						        <TR>
									<td>
										<input class="BotaoOff" onmouseover="this.className='BotaoOn'" onmouseout="this.className='BotaoOff'" type="button" value="Ok" onclick="trocarsenha();" />&nbsp;&nbsp;
										<input type="button" name="cmdRetornar" value="Retornar" onClick="voltar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Retornar a tela anterior" />
									</td>
								</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
				<script>
					document.thisForm.txtCdSenha.focus();
				</script>
			</form>
		</td>
	</tr>
</TABLE>

<!-- #include file = "../includes/LayoutEnd.asp" -->


<%

' Função desenvolvida para fazer o tratamento do
' formulário de dados.
Private Function ValidarForm()
	
	' Tratamento de campos do formulário. =============================
	
	
	If Trim(vstr_CdSenha) = "" Then
		
		Call AddErro("Senha", "Favor, preencher o campo Senha.")
		
	End If
		
	If Trim(vstr_CdNovaSenha) = "" Then
		
		Call AddErro("NovsSenha", "Favor, preencher o campo Nova Senha.")
		
	Else
		
		If Trim(vstr_CdConfirmaSenha) = "" Then
			
			Call AddErro("ConfirmaSenha", "Favor, preencher o campo Confirma Senha.")
				
		Else
			If Not vstr_CdNovaSenha = vstr_CdConfirmaSenha Then
					
				Call AddErro("ConfirmarSenha", "Favor, digitar a mesma senha <br />nos campos Nova Senha e Confirmar Senha.")
				
			Else	
				
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
				vobj_commandRegistro.CommandText					= "consultaConfirmaSenha"
				
				vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("param1",adChar, adParamInput, 10, Session("sstr_IdUsuario"))
				vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("param1",adChar, adParamInput, 15, EncriptaString(vstr_CdSenha))
				' ---------------------------------------------------------------------
				
				
				' Cria o objeto recordset com as informações do registro.	
				Set vobj_rsRegistro = vobj_commandRegistro.Execute
				
				
				If vobj_rsRegistro.EOF Then
				
					Call AddErro("SenhaErrada", "Favor, digitar a sua senha correntamente.")
					
				End If
				
				vobj_rsRegistro.Close
				Set vobj_rsRegistro = Nothing
				Set vobj_commandRegistro = Nothing
				
			End If
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