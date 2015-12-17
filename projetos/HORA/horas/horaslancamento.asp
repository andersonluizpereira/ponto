<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- #include file = "../includes/GetConnection.asp" -->
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

'Variável de controle e fluxo de acoes
Dim vstr_Executar

' Declaração de variáveis utilizadas para armazenar os
' valores dos campos da tela.
Dim vstr_IdProjetoUsuario
Dim vint_IdAtividade
Dim vboo_FlHoraSaidaLancada
Dim vint_FlTipo
Dim vstr_HrEntrada
Dim vstr_HrSaida

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
		Case "REGISTRAR"
			
			vstr_IdProjetoUsuario = Request.Form("cmbComboProjetoUsuario")
			vint_IdAtividade = Request.Form("cmbComboAtividade")
			vint_FlTipo = Request.Form("txtFlTipo")
			vstr_HrEntrada = converterHoraParaSQL(Time())
			
			vint_FlTipo = CInt(vint_FlTipo) + 1
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
				
				' ---------------------------------------------------------------------
				' Incuindo dados no banco de dados.
				' ---------------------------------------------------------------------
				Set vobj_commandProc = Server.CreateObject("ADODB.Command")
				Set vobj_commandProc.ActiveConnection = vobj_conexao
				
				
				vobj_commandProc.CommandType					= adCmdStoredProc
				vobj_commandProc.CommandText					= "incluiRegistroHora"
				
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Session("sstr_IdUsuario"))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(Date()))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adInteger, adParamInput,, vint_FlTipo)
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adChar, adParamInput, 10, vstr_IdProjetoUsuario)
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adInteger, adParamInput,, vint_IdAtividade)
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adChar, adParamInput, 10, EncriptaString(vstr_HrEntrada))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param7",adDate, adParamInput, 10, Now())
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param8",adChar, adParamInput, 25, request.ServerVariables("REMOTE_ADDR"))
				
				Call vobj_commandProc.Execute
				Set vobj_commandProc = Nothing
				' ---------------------------------------------------------------------
				
				' Abrindo a tela para não ter problemas com F5
				%>
					<script language="javascript">
						window.location = "horaslancamento.asp";
					</script>
				
				<%
				
			End If
			
		Case "REGISTRAR_SAIDA"
			
			vstr_HrSaida = converterHoraParaSQL(Time())
			
			vint_FlTipo = CInt(Request.Form("txtFlTipo"))
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
				
				' ---------------------------------------------------------------------
				' Incuindo dados no banco de dados.
				' ---------------------------------------------------------------------
				Set vobj_commandProc = Server.CreateObject("ADODB.Command")
				Set vobj_commandProc.ActiveConnection = vobj_conexao
				
				vobj_commandProc.CommandType					= adCmdStoredProc
				vobj_commandProc.CommandText					= "alteraRegistroHoraSaida"
				
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Session("sstr_IdUsuario"))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(Date()))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adInteger, adParamInput,, vint_FlTipo)
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adChar, adParamInput, 10, EncriptaString(converterHoraParaSQL(vstr_HrSaida)))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adDate, adParamInput, 10, Now())
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adChar, adParamInput, 25, request.ServerVariables("REMOTE_ADDR"))
				
				Call vobj_commandProc.Execute
				Set vobj_commandProc = Nothing
				' ---------------------------------------------------------------------
			
				' Abrindo a tela para não ter problemas com F5
				%>
					<script language="javascript">
						window.location = "horaslancamento.asp";
					</script>
				
				<%
				
			End If
	End Select
	
End If


' Declaração de variáveis locais.
Dim vobj_commandHoraSaida
Dim vobj_rsHoraSaida
			
			
' ---------------------------------------------------------------------
' Selecionando todos os Registros
' ---------------------------------------------------------------------
Set vobj_commandHoraSaida = Server.CreateObject("ADODB.Command")
Set vobj_commandHoraSaida.ActiveConnection = vobj_conexao
			
			
vobj_commandHoraSaida.CommandType				= adCmdStoredProc
vobj_commandHoraSaida.CommandText				= "consultaHoraSaidaLancada"
vobj_commandHoraSaida.Parameters.Refresh
			
vobj_commandHoraSaida.Parameters.Append vobj_commandHoraSaida.CreateParameter("param1",adChar, adParamInput, 10, Session("sstr_IdUsuario"))
vobj_commandHoraSaida.Parameters.Append vobj_commandHoraSaida.CreateParameter("param2",adDate, adParamInput, , ConverterDataParaSQL(Date()))
	
Set vobj_rsHoraSaida = vobj_commandHoraSaida.Execute
' ---------------------------------------------------------------------
			
If Not vobj_rsHoraSaida.EOF Then
				
	If vobj_rsHoraSaida("FL_SAIDA") = True Or vobj_rsHoraSaida("FL_SAIDA") = "on" Then
		
		' True neste caso, esta sendo para se há uma registro faltando data de saída.
		vboo_FlHoraSaidaLancada = True
		vstr_IdProjetoUsuario	= vobj_rsHoraSaida("ID_PROJETO")
		vint_IdAtividade		= vobj_rsHoraSaida("ID_ATIVIDADE")
	Else
		' False neste caso, esta sendo usado verificar se não tem registro pra date atual Date(), ou
		' se não tem resgitro faltando hora saida.
		vboo_FlHoraSaidaLancada = False
	End If
				
Else
	' False neste caso, esta sendo usado verificar se não tem registro pra date atual Date(), ou
	' se não tem resgitro faltando hora saida.
	vboo_FlHoraSaidaLancada = False
				
End If
			
			
vobj_rsHoraSaida.Close
Set vobj_rsHoraSaida = Nothing
Set vobj_commandHoraSaida = Nothing

%>

<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/horaslancamento.js"></script>

<table class=font width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
		<i><b class="TituloPagina">Lançamento de Horas</b></i><br><br>
		
			<form name="thisForm" action="horaslancamento.asp" method="post">
				
				<input type="hidden" name="hdnProcessar" value="S" />
				<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>" />
				<input type="hidden" name="hdnExecutar" value="<%=IIF(vboo_FlHoraSaidaLancada = True, "REGISTRAR_SAIDA", "REGISTRAR")%>"/>
				
				<TABLE class=font BORDER=0 CELLSPACING=1 CELLPADDING=1>
					<tr>
						<td align="middle">
						<fieldset style="LEFT: 0px; WIDTH: 590px; HEIGHT: 94px">
							<legend>
							   <b>Registrar Horário para <%=Session("sstr_DsUsuario")%></b>
							</legend>
							<TABLE valign="center" class=font BORDER=0 CELLSPACING=1 CELLPADDING=1>
								<tr>
									<td colspan="2"><%=ExibirErros()%></td>
								</tr>
								<tr>
									<td>Projeto</td>
									<td><%Call CriarComboProjetoUsuario("cmbComboProjetoUsuario", vstr_IdProjetoUsuario, Empty, Session("sstr_IdUsuario"))%></td>
									<td>Atividade</td>
									<td><%Call CriarComboAtividade("cmbComboAtividade", vint_IdAtividade, Empty, Empty)%></td>
								</tr>
								<tr>
									<td colspan="4" align="middle">
										<br>
											<input type="button" id="Registrar" value="Registrar" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" onclick="document.all.thisForm.submit();">
										<!--&nbsp;<input type="button" value="Voltar" onclick="voltar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';">-->
									</td>
								</tr>
							</TABLE>
						</fieldset>
						</td>
					</tr>
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 590px;">
							<legend>
							   <b>Horário Lançado</b>
							</legend>
							<table width="590" height="30" border="0" cellpadding="0" cellspacing="0" class="Tabela">
								<tr>
									<td>
									<table class=font border="0" cellSpacing="1" cellPadding="1" name="tabCabec" id="tabCabec">
										<COLGROUP>
											<col align="middle" width="100">
											<col align="middle" width="70">
											<col align="middle"	width="70">
											<col align="middle" width="80">
											<col align="middle" width="190">
											<col align="middle"	width="80">
										</COLGROUP>
										<TR class="Cabecalho">
											<TD>Data</TD>
											<TD>Entrada</TD>
											<TD>Saída</TD>
											<TD>Total</TD>
											<TD>Atividade</TD>
											<TD>Projeto</TD>
										</TR>
										
										
										<%
										' Declaração de variáveis locais.
										Dim vobj_command
										Dim vobj_rsRegistro
										Dim vint_ContadorRegistro
										
										
										' ---------------------------------------------------------------------
										' Selecionando todos os registros cadastrados na tabela.
										' ---------------------------------------------------------------------
										Set vobj_command = Server.CreateObject("ADODB.Command")
										Set vobj_command.ActiveConnection = vobj_conexao
									
										vobj_command.CommandType					= adCmdStoredProc
										vobj_command.CommandText					= "consultaHorarioLancado"
										
										
										vobj_command.Parameters.Append vobj_command.CreateParameter("param1",adChar, adParamInput, 10, Session("sstr_IdUsuario"))
										vobj_command.Parameters.Append vobj_command.CreateParameter("param2",adDate, adParamInput, , ConverterDataParaSQL(Date()))
										
										' Cria o recordset e posiciona a páginação do recordset.
										Set vobj_rsRegistro = vobj_command.Execute
										
										
										' Verifica se registros foram encontrados.
										If Not vobj_rsRegistro.EOF Then
											
											Dim vint_AuxTipo
											Dim vint_Minutos
											Dim vint_MinutosTotal
											
											vint_MinutosTotal = 0
											
											'Contador pra fazer o efeito zebrado.
											Dim contadorClass
											contadorClass = 0
											
											' Loop de todos os registros cadastrados
											' no banco de dados.
											Do While Not vobj_rsRegistro.EOF
												
												%>
												
												
												<tr class="tr<%=contadorClass Mod 2 %>">
													<td>&nbsp;<%=vobj_rsRegistro("DT_DATA")%></td>
													<td>&nbsp;<%
														
														If Not IsNull(vobj_rsRegistro("HR_ENTRADA")) Then
															
															Response.Write DesencriptaString(vobj_rsRegistro("HR_ENTRADA"))
														Else
															
															Response.Write ""
															
														End If
														
													%></td>
													<td>&nbsp;<%
														
														If Not IsNull(vobj_rsRegistro("HR_SAIDA")) Then
															
															
															Response.Write DesencriptaString(vobj_rsRegistro("HR_SAIDA"))
															vint_AuxTipo = vobj_rsRegistro("FL_TIPO")
															
														Else
															
															Response.Write ""
															vint_AuxTipo = vobj_rsRegistro("FL_TIPO")
															
														End If
														
													%></td>
													<td>&nbsp;<%
														If Not IsNull(vobj_rsRegistro("HR_ENTRADA")) And Not IsNull(vobj_rsRegistro("HR_SAIDA")) Then
															
															vint_Minutos = DateDiff("n", CDate(DesencriptaString(vobj_rsRegistro("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rsRegistro("HR_SAIDA"))))
															Response.Write converterMinutoParaHora(vint_Minutos)
															vint_MinutosTotal = vint_MinutosTotal + vint_Minutos
															
														Else
															
															Response.Write ""
															
														End If	
														%></td>
													<td>&nbsp;<%=vobj_rsRegistro("DS_ATIVIDADE")%></td>
													<td>&nbsp;<%=vobj_rsRegistro("ID_PROJETO")%></td>
												</tr>
												
												<%
												
												contadorClass = contadorClass + 1
												
												' Move para o próximo registro do loop.
												vobj_rsRegistro.MoveNext
											Loop
											
											%>
											
											<tr class="Cabecalho">
												<td>&nbsp;TOTAL</td>
												<td>&nbsp;</td>
												<td>&nbsp;</td>
												<td>&nbsp;<%=converterMinutoParaHora(vint_MinutosTotal)%></td>
												<td>&nbsp;</td>
												<td>&nbsp;</td>
											</tr>
											
											<%
										Else
											
											vint_AuxTipo = "0"
											
										End If
										
										vobj_rsRegistro.Close
										Set vobj_rsRegistro = Nothing
										Set vobj_command = Nothing
									
										%>
										
										<input type="hidden" name="txtFlTipo" value="<%=vint_AuxTipo%>">
									</table>
									</td>
								</tr>
							</table>
						</fieldset>
						</td>
					</tr>
					<tr>
						<td height="30"></td>
					</tr>
					<tr>
						<td>
							<fieldset style="LEFT: 0px; WIDTH: 590px;">
								<legend>
								   <b>Próximos Aniversariantes</b>
								</legend>
								<table width="590" height="30" border="0" cellpadding="0" cellspacing="0" class="Tabela">
									<tr>
										<td>
											<table class=font border="0" cellSpacing="1" cellPadding="1" name="tabCabec" id="tabCabec">
												<tr>
													<td><%
														
														' Declaração de variáveis locais.
														Dim vobj_commandAniversario
														Dim vobj_rsRegistroAniversario
														
														
														' ---------------------------------------------------------------------
														' Selecionando todos os registros cadastrados na tabela.
														' ---------------------------------------------------------------------
														Set vobj_commandAniversario = Server.CreateObject("ADODB.Command")
														Set vobj_commandAniversario.ActiveConnection = vobj_conexao
														
														vobj_commandAniversario.CommandType					= adCmdStoredProc
														vobj_commandAniversario.CommandText					= "consultaAniversariantes"
														
														vobj_commandAniversario.Parameters.Append vobj_commandAniversario.CreateParameter("param1",adDate, adParamInput,, converterDataParaSQL(DateSerial(2000, Month(Date()),Day(Date()))))
														vobj_commandAniversario.Parameters.Append vobj_commandAniversario.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(DateAdd("d", 15, DateSerial(2000, Month(Date()),Day(Date())))))
														
														' Cria o recordset e posiciona a páginação do recordset.
														Set vobj_rsRegistroAniversario = vobj_commandAniversario.Execute
														
														
														' Verifica se registros foram encontrados.
														If Not vobj_rsRegistroAniversario.EOF Then
															
															Dim vint_Dia
															Dim vint_Mes
															Dim vint_ContAux
															
															vint_ContAux = 0
															
															' Loop de todos os registros cadastrados
															' no banco de dados.
															Do While Not vobj_rsRegistroAniversario.EOF
																
																
																vint_Dia = Day(vobj_rsRegistroAniversario("DT_NASCIMENTO"))
																vint_Mes = Month(vobj_rsRegistroAniversario("DT_NASCIMENTO"))
																
																If Len(vint_Dia) = 1 Then
																	
																	vint_Dia = "0" & vint_Dia
																End if
																
																If Len(vint_Mes) = 1 Then
																	
																	vint_Mes = "0" & vint_Mes
																End if
																
																If vint_ContAux = 0 Then
																	
																	Response.Write Trim(vobj_rsRegistroAniversario("DS_NOME")) & " - " & "<strong>" & vint_Dia & "/" & vint_Mes & "</strong>"
																	
																Else
																	
																	Response.Write " | " & Trim(vobj_rsRegistroAniversario("DS_NOME")) & " - " & "<strong>" & vint_Dia & "/" & vint_Mes & "</strong>"
																	
																End If
																
																vint_ContAux = vint_ContAux + 1
																
																' Move para o próximo registro do loop.
																vobj_rsRegistroAniversario.MoveNext
															Loop
														Else
															
															%>
															
															Nenhum Aniversariantes nos próximos dias!
															
															<%
															
														End If
														
														vobj_rsRegistroAniversario.Close
														Set vobj_rsRegistroAniversario = Nothing
														Set vobj_commandAniversario = Nothing
														%>
															</br>
														<a class=font href="<%=getBaseLink("/manutencao/aniversariomesfiltro.asp")%>">
															<strong>Consulte os Aniversariantes</strong>
														</a>
													</td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</fieldset>
						</td>
					</tr>
				</TABLE>
			</form>
		</td>
	</tr>
</TABLE>

<!-- #include file = "../includes/LayoutEnd.asp" -->

<%

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
								' COMBO DE PROJETO USUARIO
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub CriarComboProjetoUsuario(pstr_Nome, pstr_ValorDefault, pstr_onChange, pstr_ValorAux)
	
	
	' Declaração de variáveis locais.
	Dim vobj_command
	Dim vobj_rs
	
	
	' ---------------------------------------------------------------------
	' Selecionando todos os Registros
	' ---------------------------------------------------------------------
	Set vobj_command = Server.CreateObject("ADODB.Command")
	Set vobj_command.ActiveConnection = vobj_conexao
	
	
	vobj_command.CommandType				= adCmdStoredProc
	vobj_command.CommandText				= "consultaProjetoUsuario"
	vobj_command.Parameters.Refresh
	
	vobj_command.Parameters.Append vobj_command.CreateParameter("param1", adChar, adParamInput, 10, pstr_ValorAux)
	
	Set vobj_rs = vobj_command.Execute
	' ---------------------------------------------------------------------
	
	%>
	&nbsp;
	<SELECT name="<%=pstr_Nome%>" onChange="<%=pstr_onChange%>" class="TextBox" <%=IIF(vboo_FlHoraSaidaLancada = True, "disabled", "")%> onKeyPress="javascript:fSubmitEnter(event,thisForm);">
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
	&nbsp;&nbsp;
	&nbsp;&nbsp;
	<%
		
			
	vobj_rs.Close
	Set vobj_rs = Nothing
	Set vobj_command = Nothing
	
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
								' COMBO DE ATIVIDADE
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Sub CriarComboAtividade(pstr_Nome, pstr_ValorDefault, pstr_onChange, pstr_ValorAux)
	
	
	' Declaração de variáveis locais.
	Dim vobj_command
	Dim vobj_rs
	
	
	' ---------------------------------------------------------------------
	' Selecionando todos os Registros
	' ---------------------------------------------------------------------
	Set vobj_command = Server.CreateObject("ADODB.Command")
	Set vobj_command.ActiveConnection = vobj_conexao
	
	
	vobj_command.CommandType				= adCmdStoredProc
	vobj_command.CommandText				= "consultaAtividadesAtivo"
	vobj_command.Parameters.Refresh
	
	
	Set vobj_rs = vobj_command.Execute
	' ---------------------------------------------------------------------
	
	%>
	&nbsp;
	<SELECT name="<%=pstr_Nome%>" onChange="<%=pstr_onChange%>" class="TextBox"  <%=IIF(vboo_FlHoraSaidaLancada = True, "disabled", "")%> onKeyPress="javascript:fSubmitEnter(event,thisForm);">
		<option value="<%=Empty%>">Selecione</option>
		<%
		
		
		If Not vobj_rs.EOF Then
			
			' Loop de todos os registros encontrados.
			Do While Not vobj_rs.EOF
				
				
				' Verificando se o registro
				' é o Registro default que deve ser selecionado.
				If Trim(pstr_ValorDefault) = Trim(vobj_rs("ID_ATIVIDADE")) Then
					%><OPTION selected value="<%=vobj_rs("ID_ATIVIDADE")%>"><%=vobj_rs("DS_ATIVIDADE")%></OPTION><%
				Else
					%><OPTION value="<%=vobj_rs("ID_ATIVIDADE")%>"><%=vobj_rs("DS_ATIVIDADE")%></OPTION><%
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
Private Function VerificarHoraSaidaLancada()
	
	' Declaração de variáveis locais.
	Dim vobj_command
	Dim vobj_rs
			
			
	' ---------------------------------------------------------------------
	' Selecionando todos os Registros
	' ---------------------------------------------------------------------
	Set vobj_command = Server.CreateObject("ADODB.Command")
	Set vobj_command.ActiveConnection = vobj_conexao
			
			
	vobj_command.CommandType				= adCmdStoredProc
	vobj_command.CommandText				= "consultaHoraSaidaLancada"
	vobj_command.Parameters.Refresh
			
	vobj_command.Parameters.Append vobj_command.CreateParameter("param1",adChar, adParamInput, 10, Session("sstr_IdUsuario"))
	vobj_command.Parameters.Append vobj_command.CreateParameter("param2",adDate, adParamInput, , ConverterDataParaSQL(Date()))
	
	Set vobj_rs = vobj_command.Execute
	' ---------------------------------------------------------------------
			
	If Not vobj_rs.EOF Then
				
		If vobj_rs("FL_SAIDA") = True Or vobj_rs("FL_SAIDA") = "on" Then
			
			' True neste caso, esta sendo para se há uma registro faltando data de saída.
			VerificarHoraSaidaLancada = True
		Else
			' False neste caso, esta sendo usado verificar se não tem registro pra date atual Date(), ou
			' se não tem resgitro faltando hora saida.
			VerificarHoraSaidaLancada = False
		End If
				
	Else
		' False neste caso, esta sendo usado verificar se não tem registro pra date atual Date(), ou
		' se não tem resgitro faltando hora saida.
		VerificarHoraSaidaLancada = False
				
	End If
			
			
	vobj_rs.Close
	Set vobj_rs = Nothing
	Set vobj_command = Nothing
	
	
	
End Function

' Função desenvolvida para fazer o tratamento do
' formulário de dados.
Private Function ValidarForm()
	
	Dim vstr_HoraEntrada
	Dim vstr_MinutoEntrada
	Dim vstr_HoraSaida
	Dim vstr_MinutoSaida
	
	
	If vstr_Executar = "REGISTRAR" Then
		
		' Tratamento de campos do formulário. =============================
		
		If Trim(vstr_IdProjetoUsuario) = "" Then
			
			Call AddErro("Projeto", "Favor, selecionar um projeto.")
		End If
		
		If Trim(vint_IdAtividade) = "" Then
			
			Call AddErro("Atividade", "Favor, selecionar uma atividade.")
		End If
		
		
		Dim vobj_commandFaixaHora
		Dim vobj_rsFaixaHora
		
		Set vobj_commandFaixaHora = Server.CreateObject("ADODB.Command")
		Set vobj_commandFaixaHora.ActiveConnection = vobj_conexao
		
		
		vobj_commandFaixaHora.CommandType				= adCmdStoredProc
		vobj_commandFaixaHora.CommandText				= "consultaFaixaHora"
		vobj_commandFaixaHora.Parameters.Refresh
		
		vobj_commandFaixaHora.Parameters.Append vobj_commandFaixaHora.CreateParameter("param1",adChar, adParamInput, 10, Session("sstr_IdUsuario"))
		vobj_commandFaixaHora.Parameters.Append vobj_commandFaixaHora.CreateParameter("param2",adDate, adParamInput, , ConverterDataParaSQL(Date()))
		
		Set vobj_rsFaixaHora = vobj_commandFaixaHora.Execute
		' ---------------------------------------------------------------------
		
		If Not vobj_rsFaixaHora.EOF Then
			
			Dim vstr_HoraEntradaAtual
			Dim vstr_MinutoEntradaAtual
			
			vstr_HoraEntradaAtual = CInt(Hour(vstr_HrEntrada))
			vstr_MinutoEntradaAtual = CInt(Minute(vstr_HrEntrada))
			
			Do While Not vobj_rsFaixaHora.EOF
				
				vstr_HoraEntrada = CInt(Hour(DesencriptaString(vobj_rsFaixaHora("HR_ENTRADA"))))
				vstr_MinutoEntrada = CInt(Minute(DesencriptaString(vobj_rsFaixaHora("HR_ENTRADA"))))
				
				vstr_HoraSaida = CInt(Hour(DesencriptaString(vobj_rsFaixaHora("HR_SAIDA"))))
				vstr_MinutoSaida = CInt(Minute(DesencriptaString(vobj_rsFaixaHora("HR_SAIDA"))))
				
				If vstr_HoraEntradaAtual >= vstr_HoraEntrada And vstr_HoraEntradaAtual <= vstr_HoraSaida Then
					
					If vstr_MinutoEntradaAtual >= vstr_MinutoEntrada And vstr_MinutoEntradaAtual <= vstr_MinutoSaida Then
						
						Call AddErro("HoraEntrada", "Não foi possivel o cadastramento de hora, aguarde um minuto e tente novamente, ou entre em contato com o Administrador.")
						
					End If
					
				End If
				
				vobj_rsFaixaHora.MoveNext
			Loop
		End If
		
		vobj_rsFaixaHora.Close
		Set vobj_rsFaixaHora = Nothing
		Set vobj_commandFaixaHora = Nothing
		
	End If
	
	If vstr_Executar = "REGISTRAR_SAIDA" Then
		
		Dim vobj_commandFaixaHoraSaida
		Dim vobj_rsFaixaHoraSaida
		
		Set vobj_commandFaixaHoraSaida = Server.CreateObject("ADODB.Command")
		Set vobj_commandFaixaHoraSaida.ActiveConnection = vobj_conexao
		
		
		vobj_commandFaixaHoraSaida.CommandType				= adCmdStoredProc
		vobj_commandFaixaHoraSaida.CommandText				= "consultaFaixaHora"
		vobj_commandFaixaHoraSaida.Parameters.Refresh
		
		vobj_commandFaixaHoraSaida.Parameters.Append vobj_commandFaixaHoraSaida.CreateParameter("param1",adChar, adParamInput, 10, Session("sstr_IdUsuario"))
		vobj_commandFaixaHoraSaida.Parameters.Append vobj_commandFaixaHoraSaida.CreateParameter("param2",adDate, adParamInput, , ConverterDataParaSQL(Date()))
		
		Set vobj_rsFaixaHoraSaida = vobj_commandFaixaHoraSaida.Execute
		' ---------------------------------------------------------------------
		
		If Not vobj_rsFaixaHoraSaida.EOF Then
			
			Dim vstr_HoraSaidaAtual
			Dim vstr_MinutoSaidaAtual
			
			vstr_HoraSaidaAtual = CInt(Hour(vstr_HrSaida))
			vstr_MinutoSaidaAtual = CInt(Minute(vstr_HrSaida))
			
			Do While Not vobj_rsFaixaHoraSaida.EOF
				
				vstr_HoraEntrada = CInt(Hour(DesencriptaString(vobj_rsFaixaHoraSaida("HR_ENTRADA"))))
				vstr_MinutoEntrada = CInt(Minute(DesencriptaString(vobj_rsFaixaHoraSaida("HR_ENTRADA"))))
				
				If Not vint_FlTipo = vobj_rsFaixaHoraSaida("FL_TIPO") Then
					
					vstr_HoraSaida = CInt(Hour(DesencriptaString(vobj_rsFaixaHoraSaida("HR_SAIDA"))))
					vstr_MinutoSaida = CInt(Minute(DesencriptaString(vobj_rsFaixaHoraSaida("HR_SAIDA"))))
					
					If vstr_HoraSaidaAtual >= vstr_HoraEntrada And vstr_HoraSaidaAtual <= vstr_HoraSaida Then
						
						If vstr_MinutoSaidaAtual >= vstr_MinutoEntrada And vstr_MinutoSaidaAtual <= vstr_MinutoSaida Then
							
							Call AddErro("HoraEntrada", "Não foi possivel o cadastramento de hora, entre em contato com o Administrador.")
							
						End If
						
					End If
					
				Else
					
					If vstr_HoraSaidaAtual < vstr_HoraEntrada Then
						
						Call AddErro("HoraEntrada", "Não foi possivel o cadastramento de hora, entre em contato com o Administrador.")
						
					ElseIf vstr_HoraSaidaAtual = vstr_HoraEntrada Then
						
						If vstr_MinutoSaidaAtual < vstr_MinutoEntrada Then
							
							Call AddErro("HoraEntrada", "Não foi possivel o cadastramento de hora, entre em contato com o Administrador.")
							
						ElseIf vstr_MinutoSaidaAtual = vstr_MinutoEntrada Then
							
							Call AddErro("HoraEntrada", "Não foi possivel o cadastramento de hora, aguarde um minuto e tente novamente.")
							
						End If
						
					End If
					
				End If
				
				vobj_rsFaixaHoraSaida.MoveNext
			Loop
		End If
		
		vobj_rsFaixaHoraSaida.Close
		Set vobj_rsFaixaHoraSaida = Nothing
		Set vobj_commandFaixaHoraSaida = Nothing
		
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

<script language="javascript">
	if(document.all.hdnExecutar.value=="REGISTRAR"){
		document.all.cmbComboProjetoUsuario.focus();
	}
	else{
		document.all.Registrar.focus();
	}
</script>

<!-- #include file = "../includes/CloseConnection.asp" -->