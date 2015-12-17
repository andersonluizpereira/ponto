<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- S#include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->

<%

' Declara��o de vari�veis locais. ==============================================

' Guarda a opera��o que ser� executa nesta tela.
' Obs.: Seus valores podem ser A = Altera��o, I = Inclus�o, V = Visualiza��o.
Dim vstr_Operacao

' Vari�vel flag que indica se a p�gina deve ser 
' processada, apenas disponivel para as opera��es de
' A e I.
Dim vstr_Processar


Dim vstr_CdSenha
Dim vstr_CdNovaSenha
Dim vstr_CdConfirmaSenha

' para est� p�gina.
vstr_Operacao		= Request.Form("pstr_Operacao")
vstr_Processar		= Request.Form("hdnProcessar")


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
			
			vstr_CdSenha			= Request.Form("txtCdSenha")
			vstr_CdNovaSenha		= Request.Form("txtCdNovaSenha")
			vstr_CdConfirmaSenha	= Request.Form("txtCdConfirmaSenha")
			
			
			' Verificando se o formul�rio foi
			' devidamente v�lidado pelo sistema.
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

' Fun��o desenvolvida para fazer o tratamento do
' formul�rio de dados.
Private Function ValidarForm()
	
	' Tratamento de campos do formul�rio. =============================
	
	
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
				vobj_commandRegistro.CommandText					= "consultaConfirmaSenha"
				
				vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("param1",adChar, adParamInput, 10, Session("sstr_IdUsuario"))
				vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("param1",adChar, adParamInput, 15, EncriptaString(vstr_CdSenha))
				' ---------------------------------------------------------------------
				
				
				' Cria o objeto recordset com as informa��es do registro.	
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