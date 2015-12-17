<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- #include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->



<%


' Guarda a operação que será executa nesta tela.
' Obs.: Seus valores podem ser A = Alteração, I = Inclusão, V = Visualização.
Dim vstr_Operacao

' Variável flag que indica se a página deve ser 
' processada, apenas disponivel para as operações de
' A e I.
Dim vstr_Processar


' Declaração de variáveis utilizadas para armazenar os
' valores dos campos da tela.
Dim vstr_DsMural
Dim vboo_FlAnnonimo
Dim vint_UltimaMensagem

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
		
		Case "I"
			
			' ... processamento de inclusão de registro.
			
			vstr_DsMural = Trim(Request.Form("txtAreaMural"))
			'vboo_FlAnnonimo = Request.Form("chbFlAnonimo")
			vboo_FlAnnonimo = ""
			
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
			vobj_commandRegistroConsulta.CommandText					= "consultaUltimaMensagemMural"
				
				
			' ---------------------------------------------------------------------
				
				
			' Cria o objeto recordset com as informações do registro.	
			Set vobj_rsRegistroConsulta = vobj_commandRegistroConsulta.Execute
				
			'Verificando se ja ha registro no banco com mesma area
			'Obs. É verificado soment campo Area, campo nome pode haver dois iguais.
			If Not vobj_rsRegistroConsulta.EOF Then
					
				vint_UltimaMensagem = cInt(vobj_rsRegistroConsulta("ID_MURAL")) + 1
				
			End If
				
			vobj_rsRegistroConsulta.Close
			Set vobj_rsRegistroConsulta = Nothing
			Set vobj_commandRegistroConsulta = Nothing
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
					
					Dim vstr_Usuario
					
					' ---------------------------------------------------------------------
					' Incluindo os dados do registro no banco de dados.
					' ---------------------------------------------------------------------
					Set vobj_commandProc = Server.CreateObject("ADODB.Command")
					Set vobj_commandProc.ActiveConnection = vobj_conexao
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "incluiMural"
					vobj_commandProc.Parameters.Refresh
					
					'If vboo_FlAnnonimo = "" Then
						
					'	vstr_Usuario = Session("sstr_IdUsuario")
					'Else
					'	If vboo_FlAnnonimo = "on" Or vboo_FlAnnonimo = True Then
							
					'		vstr_Usuario = "Anônimo"
					'	End If
					'End If
					
					vstr_Usuario = Session("sstr_IdUsuario")
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 10, converterHoraParaMural(Time()))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adChar, adParamInput, 10, converterDataParaHtml(Date()))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adChar, adParamInput, 10, vstr_Usuario)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adChar, adParamInput, 255, vstr_DsMural)
					
					
					vobj_commandProc.Execute
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "alterarUsuarioNovaMensagem"
					vobj_commandProc.Parameters.Refresh
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adChar, adParamInput, 10, vstr_Usuario)
					
					vobj_commandProc.Execute
					
					Set vobj_commandProc = Nothing
					
					
			End If
			
		Case "EXCLUIR"											' -->> Operação de Exclusão de Registros.
		
		
		' ... operação de exclusão de registros.
		
		
		' Declaração de variáveis auxiliares que
		' auxiliarão na exclusão dos registros selecionados.
		Dim vstr_ExcluirIdRegistros
		Dim varrvar_IdRegistro
		Dim vvar_IdRegistro
		Dim vobj_commandExclusao
		
		
		' Conseguindo todos os registros selecionados para
		' a exclusão do banco de dados.
		vstr_ExcluirIdRegistros = Request("hdnExcluir")
		
		
		' Dividindo todos os códigos a serem excluidos em 
		' uma matriz de array.
		varrvar_IdRegistro = Split(vstr_ExcluirIdRegistros, ",")
		
		
		' ---------------------------------------------------------------------
		' Exclusão de Registros do banco de dados.
		' ---------------------------------------------------------------------
		Set vobj_commandExclusao = Server.CreateObject("ADODB.Command")
		Set vobj_commandExclusao.ActiveConnection = vobj_conexao
		
		
		' ---------------------------------------------------------------------
		
		
		' Ignorando os erros que ocorrem na exclusão
		' do registro do banco de dados.
		On Error Resume Next
		
		
		' Loop por todos os registros selecionados para ser
		' excluso do banco de dados.
		For Each vvar_IdRegistro In varrvar_IdRegistro
			
			vobj_commandExclusao.CommandType					= adCmdStoredProc
			vobj_commandExclusao.CommandText					= "excluiMensagem"
			vobj_commandExclusao.Parameters.Refresh
			
			
			' Passando o código do Registro a ser
			' excluido do banco de dados.
			vobj_commandExclusao.Parameters.Append vobj_commandExclusao.CreateParameter("param1",adInteger, adParamInput, , CInt(vvar_IdRegistro))
			
			' Chamando comando para excluir o registro
			Call vobj_commandExclusao.Execute
		Next
		
		
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
		Set vobj_commandExclusao = Nothing
		
	End Select
End If


' Chamando funcao que desmarca nova mensagem pra o usuario
Call MuralVisitado(Session("sstr_IdUsuario"))

%>

<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/muralrecado.js"></script>
<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<form name="thisForm" action="muralrecado.asp" method="post">
					
				<input type="hidden" name="hdnProcessar" value="S">
				<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>">
				<input type="hidden" name="hdnExcluir">
					
				<i><b class="TituloPagina">Mural de Recados</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
			        <TBODY>
						<tr>	
							<td colspan="2">
								<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabPesquisa" id="tabPesquisa" width="100%" style="FILTER: alpha(opacity  =80)"><tr>
									<tr>
										<td><iframe src="frmensagem.asp" name="iframeMensagem" frameBorder="no" width="800" height="220" style="overflow-x:hidden" scrolling="auto"></iframe></td>
									</tr>
									
									<%
									
									If	Session("sboo_fladministrador") = True Then
										
										%>
									
										<tr>
											<td>
												<TEXTAREA id=area2 style="WIDTH: 800px; HEIGHT: 80px" name="txtAreaMural" rows=23 cols=93 maxlength="255"></TEXTAREA>
											</td>
										</tr>
										
										<%
										
									End If
									
									%>
									
								</table>
								<table width="100%">
								<TR>
									<td align="middle">
										<table  width="100%" ALIGN="center" BORDER="0" CELLSPACING="1" CELLPADDING="1">
											<tr>
												<td width="10%" valign="top">
													
													<%
													
													If	Session("sboo_fladministrador") = True Then
														
														%>
														
														<input type="button" value="Enviar" onclick="enviar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Enviar Mensagem">
														<input type="button" value="Excluir" onclick="excluir();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Excluir Mensagens Selecionadas">
														
														<%
														
													End If
													
													%>
													
													<input type="button" value="Retornar" onclick="voltar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Retornar a tela anterior">
												</td>
												<!--<td width="10%" valign="top">
														&nbsp; 
												</td>
												<td class="font" width="10%" valign="top">Anônimo:&nbsp;</td>
												<td width="10%" valign="top">
												<input name="chbFlAnonimo" id="Anonimo" type="checkbox" size="54" maxlength="100">
												</td>
												<td width="50%" valign="top">
													&nbsp;
												</td>-->
											</tr>
										</table>
									</td>
								</TR>
								</table>
							</TD>
						</TR>
					</TBODY>
				</table>
			</FORM>
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
	
	
	If Trim(vstr_DsMural) = "" Then
		
		Call AddErro("Sigla", "Favor, Escrever uma Mensagem")
		
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

Private Function converterHoraParaMural(Hora)
	
	Dim vstr_Hora
	Dim vstr_Minuto
	Dim vint_Segundo
	
	vstr_Hora = CStr(Hour(Hora))
	vstr_Minuto = CStr(Minute(Hora))
	vint_Segundo = CStr(Second(Hora))
	
	If Len(vstr_Hora) = 1 Then
		
		vstr_Hora = "0" & vstr_Hora
	End If
	
	If Len(vstr_Minuto) = 1 Then
		
		vstr_Minuto = "0" & vstr_Minuto
	End If
	
	If Len(vint_Segundo) = 1 Then
		
		vint_Segundo = "0" & vint_Segundo
	End If
	
	converterHoraParaMural = vstr_Hora & ":" & vstr_Minuto & ":" & vint_Segundo
	
End Function


' Procedimento desenvolvido para informar ao sistema que o usuario visitou a pagina de mural,
' com isso não informar mais que há mensagens novas para ele.
Private Sub MuralVisitado(pstr_Usuario)
	
	' Declaração de variáveis auxiliares
	' para obter as informações do registro.
	Dim vobj_commandMuralVisitado
	
	
	' ---------------------------------------------------------------------
	' Selecionando os dados do registro.
	' ---------------------------------------------------------------------
	Set vobj_commandMuralVisitado = Server.CreateObject("ADODB.Command")
	Set vobj_commandMuralVisitado.ActiveConnection = vobj_conexao
	
	
	vobj_commandMuralVisitado.CommandType					= adCmdStoredProc
	vobj_commandMuralVisitado.CommandText					= "consultaMuralVisitado"
	
	
	vobj_commandMuralVisitado.Parameters.Append vobj_commandMuralVisitado.CreateParameter("param1",adChar, adParamInput, 10, pstr_Usuario)
	
	
	'Executar Command
	Call vobj_commandMuralVisitado.Execute
	
	Set vobj_commandMuralVisitado = Nothing
	
End Sub

%>

<!-- #include file = "../includes/CloseConnection.asp" -->