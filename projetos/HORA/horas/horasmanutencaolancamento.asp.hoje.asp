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
Dim vstr_IdUsuario

' Declaração de variáveis utilizadas para armazenar os
' valores dos campos da tela.
Dim vint_NmMes
Dim vstr_DsAno
Dim vstr_DtData
Dim vstr_IdProjetoUsuario
Dim vint_IdAtividade
Dim vint_FlTipo
Dim vstr_HrEntrada
Dim vstr_HrSaida



Dim vint_FlBloqueaData
Dim vboo_ValidaAlteracao

' para está página.
vstr_Operacao		= Request.Form("pstr_Operacao")
vstr_Processar		= Request.Form("hdnProcessar")
vstr_Executar		= Request.Form("hdnExecutar")

vstr_IdUsuario		= Request.Form("cmbComboUsuario")
vint_NmMes			= Cint(Request.Form("cmbComboMes"))
vstr_DsAno			= Cint(Request.Form("txtDsAno"))



vint_FlBloqueaData	= Request.Form("hdnBloqueaData")


If vstr_IdUsuario = "" Then
	
	Response.Redirect("horasmanutencaofiltro.asp")
	
ElseIf vint_NmMes = "" Then
	
	Response.Redirect("horasmanutencaofiltro.asp")

ElseIf vstr_DsAno = "" Or Not Len(vstr_DsAno) = 4 Or Not IsNumeric(vstr_DsAno) Then

	Response.Redirect("horasmanutencaofiltro.asp")
	
End If

' Verifica se o parametro que defini o tipo
' de operação a ser executado na página é
' igual a branco(vazio).
If Trim(vstr_Operacao) = "" Then
	
	' ... neste caso a operação
	' padrão de a de visualização apenas
	' do registro.
	vstr_Operacao = "V"
End If

' *******************************************************
' INICIO DA ROTINA QUE CONSEGUE OS DADOS DO REGISTRO
' *******************************************************

' Veririfa se a operação a ser executada nesta página é a 
' operação de Alteração ou Visualização e se a página não
' foi processada ainda.
If (vstr_Operacao = "A") And vstr_Processar <> "S" Then
	
	
	' ... neste caso deve ser solicitado o código do registro
	' e encontrar suas informações no banco de dados para exibir para
	' as informações do registro na tela.
	' Conseguindo o código do registro.
	vstr_DtData = Request.Form("txtDsData")
	
	
Else
	
	
	' Verifica se a operação a ser executada nesta página é
	' a operação de inclusão e verifica se a página não foi
	' processada ainda.
	If vstr_Operacao = "AI" And vstr_Processar <> "S" Then
		
		' Neste caso todas as variáveis devem ser vazias
		' para o usuário poder preencher seu novo cadastro
		' do registro.
		
		vstr_DtData = Request.Form("hdnDtData")
		
	ElseIf vstr_Operacao = "I" And vstr_Processar = "S" Then
		
		' ... está opção acontecerá quando o usuário processar
		' a página, neste caso todas os dados da tela serão
		' submetidos e devem ser pegos neste lugar.
		
		vstr_DtData = Request.Form("txtDsData")
		
		If isDate(vstr_DtData) Then
			
			vint_FlBloqueaData = "1"
			
		Else
			
			vint_FlBloqueaData = ""
			
		End If
		
	End If
End If
' *******************************************************
' FINAL DA ROTINA QUE CONSEGUE OS DADOS DO REGISTRO
' *******************************************************

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
	Dim vobj_commandProc1
	
	
	' Analiza a operação a ser executada na página
	' para descobrir o processamento que deve ser feito.
	Select Case vstr_Operacao
		
		
		Case "A"						' Operação de alteração do registro.
			
			Dim varr_IdProjetoUsuario
			Dim varr_IdAtividade
			Dim varr_FlTipo
			Dim varr_HrEntrada
			Dim varr_HrSaida
			Dim obs1
			
			varr_IdProjetoUsuario = Split(Request.Form("cmbComboProjetoUsuarioAltera"), ",")
			varr_IdAtividade = Split(Request.Form("cmbComboAtividadeAltera"), ",")
			varr_FlTipo = Split(Request.Form("hdnFlTipoAltera"), ",")
			varr_HrEntrada = Split(Request.Form("txtHrEntradaAltera"), ",")
			obs1 = Request.Form("txtObs")
			
			
			If Trim(Request.Form("txtHrSaidaAltera")) = "" Then
			
				varr_HrSaida = Split(" ", ",")
				
			Else
				
				varr_HrSaida = Split(Request.Form("txtHrSaidaAltera"), ",")
				
			End If
			vstr_DtData = Request.Form("txtDsData")
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
				
				Dim vint_Contadoralteracao
				
				' ---------------------------------------------------------------------
				' Incuindo dados no banco de dados.
				' ---------------------------------------------------------------------
				Set vobj_commandProc = Server.CreateObject("ADODB.Command")
				Set vobj_commandProc1 = Server.CreateObject("ADODB.Command")
				
				Set vobj_commandProc.ActiveConnection = vobj_conexao
				Set vobj_commandProc1.ActiveConnection = vobj_conexao
				
				For vint_Contadoralteracao = LBound(varr_FlTipo) To UBound(varr_FlTipo)
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc1.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "alteraRegistroHoraManutencao"
					vobj_commandProc1.CommandText					= "alteraRegistroObs"
					
					
					
					
					vobj_commandProc.Parameters.Refresh
					vobj_commandProc1.Parameters.Refresh
					
					
					
					vobj_commandProc1.Parameters.Append vobj_commandProc1.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_IdUsuario))
					vobj_commandProc1.Parameters.Append vobj_commandProc1.CreateParameter("param2",adDate, adParamInput,, Trim(vstr_DtData))
					vobj_commandProc1.Parameters.Append vobj_commandProc1.CreateParameter("param3",adChar, adParamInput,255, obs1)
					
					
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_IdUsuario))
					' vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(Trim(vstr_DtData)))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adDate, adParamInput,, Trim(vstr_DtData))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adInteger, adParamInput,, cInt(varr_FlTipo(vint_Contadoralteracao)))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adChar, adParamInput, 10, Trim(varr_IdProjetoUsuario(vint_Contadoralteracao)))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adInteger, adParamInput,, cInt(varr_IdAtividade(vint_Contadoralteracao)))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adChar, adParamInput, 10, EncriptaString(converterHoraParaSQL(varr_HrEntrada(vint_Contadoralteracao))))
					
					
					If Not Trim(varr_HrSaida(vint_Contadoralteracao)) = "" Then
						
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param7",adChar, adParamInput, 10, EncriptaString(converterHoraParaSQL(varr_HrSaida(vint_Contadoralteracao))))
						
					Else
						
						vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param7",adChar, adParamInput, 10, Null)
						
					End If
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param8",adDate, adParamInput, 10, Now())
					
					Call vobj_commandProc.Execute
					Call vobj_commandProc1.Execute
					
					
					
					
					' ---------------------------------------------------------------------
				
				Next
				
				Set vobj_commandProc = Nothing
				Set vobj_commandProc1 = Nothing
				
				Call OrganizarDiaHora(vstr_IdUsuario, vstr_DtData)
				
				vstr_IdProjetoUsuario = Empty
				vint_IdAtividade = Empty
				vint_FlTipo = Empty
				vstr_HrEntrada = Empty
				vstr_HrSaida = Empty
				
				vint_FlBloqueaData = "1"
				vstr_Operacao = "I"
			
			Else
				
				vint_FlBloqueaData = "1"
				
				vboo_ValidaAlteracao = True
				
			End If
			
		Case "I"						' Operação de inclusão do registro.
			
			
			vstr_IdProjetoUsuario = Request.Form("cmbComboProjetoUsuario")
			vint_IdAtividade = Request.Form("cmbComboAtividade")
			vint_FlTipo = 1
			vstr_HrEntrada = CStr(Request.Form("txtHrEntrada"))
			vstr_HrSaida = CStr(Request.Form("txtHrSaida"))
			vstr_DtData = Request.Form("txtDsData")
			
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
				
				' ---------------------------------------------------------------------
				' Incuindo dados no banco de dados.
				' ---------------------------------------------------------------------
				Set vobj_commandProc = Server.CreateObject("ADODB.Command")
				Set vobj_commandProc.ActiveConnection = vobj_conexao
				
				
				vobj_commandProc.CommandType					= adCmdStoredProc
				vobj_commandProc.CommandText					= "incluiRegistroHoraManutencao"
				
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, vstr_IdUsuario)
				' vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(vstr_DtData))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adDate, adParamInput,, vstr_DtData)
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adInteger, adParamInput,, vint_FlTipo)
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adChar, adParamInput, 10, vstr_IdProjetoUsuario)
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adInteger, adParamInput,, vint_IdAtividade)
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adChar, adParamInput, 10, EncriptaString(converterHoraParaSQL(vstr_HrEntrada)))
				
				
				If Not Trim(vstr_HrSaida) = "" Then
						
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adChar, adParamInput, 10, EncriptaString(converterHoraParaSQL(vstr_HrSaida)))
						
				Else
						
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adChar, adParamInput, 10, Null)
						
				End If
				
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param7",adDate, adParamInput, 10, Now())
				
				Call vobj_commandProc.Execute
				Set vobj_commandProc = Nothing
				' ---------------------------------------------------------------------
				
				Call OrganizarDiaHora(vstr_IdUsuario, vstr_DtData)
				
				vstr_IdProjetoUsuario = Empty
				vint_IdAtividade = Empty
				vint_FlTipo = Empty
				vstr_HrEntrada = Empty
				vstr_HrSaida = Empty
				vint_FlBloqueaData = "1"
				
				vstr_Operacao = "I"
				
			End If
			
		
		Case "E"
			
			vstr_DtData = Request.Form("txtDsData")
			vint_FlTipo = Request.Form("hdnFlTipo")
			
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
				
				
				' ---------------------------------------------------------------------
				' Excluindo dados no banco de dados.
				' ---------------------------------------------------------------------
				Set vobj_commandProc = Server.CreateObject("ADODB.Command")
				Set vobj_commandProc.ActiveConnection = vobj_conexao
				
				
				vobj_commandProc.CommandType					= adCmdStoredProc
				vobj_commandProc.CommandText					= "excluiRegistroHoraManutencao"
				
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, vstr_IdUsuario)
				' vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(vstr_DtData))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adDate, adParamInput,, vstr_DtData)
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adInteger, adParamInput,, vint_FlTipo)
				
				Call vobj_commandProc.Execute
				Set vobj_commandProc = Nothing
				' ---------------------------------------------------------------------
				
				Call OrganizarDiaHora(vstr_IdUsuario, vstr_DtData)
				
				vstr_IdProjetoUsuario = Empty
				vint_IdAtividade = Empty
				vint_FlTipo = Empty
				vstr_HrEntrada = Empty
				vstr_HrSaida = Empty
				vint_FlBloqueaData = "1"
				
				vstr_Operacao = "I"
				
			End If
				
	End Select
	
End If
' *******************************************************
' FINAL DA ROTINA QUE FAZ O PROCESSAMENTO DOS DADOS
' DO REGISTRO.
' *******************************************************


%>


<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/horasmanutencaolancamento.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<i><b class="TituloPagina">Manutenção de Horas</b></i><br>
			<br>
			<form name="thisForm" action="horasmanutencaolancamento.asp" method="post">
						
				<input type="hidden" name="hdnProcessar" value="S">
				<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>">
				<input type="hidden" name="hdnExecutar" />
				<input type="hidden" name="hdnFlTipo" value="<%=vint_FlTipo%>" />
				<input type="hidden" name="cmbComboUsuario" value="<%=vstr_IdUsuario%>" />
				<input type="hidden" name="cmbComboMes" value="<%=vint_NmMes%>" />
				<input type="hidden" name="txtDsAno" value="<%=vstr_DsAno%>" />
				<input type="hidden" name="hdnBloqueaData" />
				
						
				<TABLE class="font" BORDER="0" CELLSPACING="1" CELLPADDING="1" align="center">
					<tr>
						<td>
							<fieldset style="LEFT: 0px; WIDTH: 600px; HEIGHT: 140px">
								<legend>
									<strong>Registrar Horário para <%=vstr_IdUsuario%></strong>
								</legend>
								<TABLE align="center" class="font" BORDER="0" CELLSPACING="1" CELLPADDING="1">
									<tr>
										<td colspan="6"><%=ExibirErros()%></td>
									</tr>
									<tr>
										<td align="left">Data:</td>
										<td align="left"><input class="TextBox" type="text" name="txtDsData" size="11" maxlength="10" value="<%=vstr_DtData%>" <%
													
											If vint_FlBloqueaData = "1" Then
														
												Response.Write "ReadOnly"
												
											End If
													
											%> /></td>
										<td align="left">Entrada:</td>
										<td align="left"><input class="TextBox" type="text" name="txtHrEntrada" size="6" maxlength="5" value="<%=vstr_HrEntrada%>" /></td>
										<td align="left">Saída:</td>
										<td align="left"><input class="TextBox" type="text" name="txtHrSaida" size="6" maxlength="5" value="<%=vstr_HrSaida%>" /></td>
									</tr>
									<tr>
										<td align="left">Projeto</td>
										<td align="left"><%Call CriarComboProjetoUsuario("cmbComboProjetoUsuario", vstr_IdProjetoUsuario, Empty, vstr_IdUsuario)%></td>
										<td align="left">Atividade</td>
										<td align="left"><%Call CriarComboAtividade("cmbComboAtividade", vint_IdAtividade, Empty, Empty)%></td>
									</tr>
									<tr>
										<td colspan="6" align="center">
											<br>
											<input type="button" value="Inserir Novo" onclick="registrar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';">
											&nbsp;<input id="btnVoltar" type="button" value="Voltar" onclick='voltar(thisForm);' class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';">
										</td>
									</tr>
									<tr>
										<td colspan="6" align="center">&nbsp;<input type="hidden" name="Tipo" id="Tipo"></td>
									</tr>
								</TABLE>
							</fieldset>
						</td>
					</tr>
					<tr>
						<td>
							<fieldset style="LEFT: 0px; WIDTH: 600px">
								<legend>
									<b>Manutenção do Horário Lançado</b>
								</legend>
								<table width="590" height="30" border="0" cellpadding="0" cellspacing="0" class="font">
									<tr>
										<td>
											<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabCabec" id="tabCabec">
												<COLGROUP />
													<col align="center" width="25" />
													<col align="center" width="70" />
													<col align="center" width="70" />
													<col align="center" width="70" />
													<col align="center" width="80" />
													<col align="center" width="190" />
													<col align="center" width="70" />
													<TR class="Cabecalho">
														<TD>&nbsp;</TD>
														<TD>Data</TD>
														<TD>Entrada</TD>
														<TD>Saída</TD>
														<TD>Total</TD>
														<TD>Projeto</TD>
														<TD>Atividade</TD>
													</TR>
												<COLGROUP />
													<col align="center" width="25" />
													<col align="center" width="70" />
													<col align="center" width="70" />
													<col align="center" width="70" />
													<col align="center" width="80" />
													<col align="center" width="190" />
													<col align="center" width="70" />
													<%
															
													' Declaração de variáveis locais.
													Dim vobj_command
													Dim vobj_rs
													
															
													' ---------------------------------------------------------------------
													' Selecionando todos os Registros
													' ---------------------------------------------------------------------
													Set vobj_command = Server.CreateObject("ADODB.Command")
													Set vobj_command.ActiveConnection = vobj_conexao
															
													If IsDate(vstr_DtData) And Not Trim(vstr_DtData) = "" Then
														
														'Contador pra fazer o efeito zebrado.
														Dim contadorClass
														contadorClass = 0

														If vboo_ValidaAlteracao = True Then
														
															Dim vint_ContadorFalso
															
															For vint_ContadorFalso = LBound(varr_FlTipo) To UBound(varr_FlTipo)
																
																%>
																
																<tr class="tr<%=contadorClass Mod 2 %>" id="trLinhaRegistro" name="trLinhaRegistro">
																	<td style="cursor: hand" onclick="excluir('<%=varr_FlTipo(vint_ContadorFalso)%>');" title="Exclui Hora">
																		<img src="../images/star_off.gif" />
																	</td>
																	<td><%=vstr_DtData%>
																		<input type="hidden" name="hdnFlTipoAltera" value="<%=varr_FlTipo(vint_ContadorFalso)%>" />
																	</td>
																	<td><input class="TextBox" type="text" name="txtHrEntradaAltera" size="3" maxlength="5" value="<%=Trim(varr_HrEntrada(vint_ContadorFalso))%>" /></td>
																	<td><input class="TextBox" type="text" name="txtHrSaidaAltera" size="3" maxlength="5" value="<%=Trim(varr_HrSaida(vint_ContadorFalso))%>" /></td>
																	<td>&nbsp;</td>
																	<td><%Call CriarComboProjetoUsuario("cmbComboProjetoUsuarioAltera", Trim(varr_IdProjetoUsuario(vint_Contadoralteracao)), Empty, vstr_IdUsuario)%></td>
																	<td><%Call CriarComboAtividade("cmbComboAtividadeAltera", cInt(varr_IdAtividade(vint_Contadoralteracao)), Empty, Empty)%></td>
																</tr>
																<%
																
																contadorClass = contadorClass + 1
															Next
															
															%>
															
															<tr>
																<td colspan="2"><br /><input type="button" value="Alterar" onclick="alterar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" id=button1 name=button1></td>
															<tr>
															
															<%
															
														Else
															
															vobj_command.CommandType				= adCmdStoredProc
															vobj_command.CommandText				= "consultaManutencaoLancamento"
															vobj_command.Parameters.Refresh
															vobj_command.Parameters.Append vobj_command.CreateParameter("param1",adChar, adParamInput, 10, vstr_IdUsuario)
															vobj_command.Parameters.Append vobj_command.CreateParameter("param2",adDate, adParamInput,, vstr_DtData)
															
															
															

															
															Set vobj_rs = vobj_command.Execute
																	
															If Not vobj_rs.EOF Then
																															
																Dim vint_Minutos
																Dim vint_MinutosTotal
																		
																
																vint_MinutosTotal = 0
																
																
																' Loop de todos os registros encontrados.
																Do While Not vobj_rs.EOF
																																					
																	%>
																	<!--
																	<tr style="cursor: hand" id="atualAlterar_<%=vobj_rs("FL_TIPO")%>">
																	-->
																	<tr class="tr<%=contadorClass Mod 2 %>" id="trLinhaRegistro" name="trLinhaRegistro">
																		<td style="cursor: hand" onclick="excluir('<%=vobj_rs("FL_TIPO")%>');" title="Exclui Hora">
																			<img src="../images/star_off.gif" />
																		</td>
																		<td><%=converterDataParaHtml(vobj_rs("DT_DATA"))%>
																			<input type="hidden" name="hdnFlTipoAltera" value="<%=vobj_rs("FL_TIPO")%>" />
																		</td>
																		<td><input class="TextBox" type="text" name="txtHrEntradaAltera" size="3" maxlength="5" value="<%=DesencriptaString(vobj_rs("HR_ENTRADA"))%>" /></td>
																		<td><input class="TextBox" type="text" name="txtHrSaidaAltera" size="3" maxlength="5" value="<%
																			
																			If Not IsNull(vobj_rs("HR_SAIDA")) Then
																				
																				Response.Write DesencriptaString(vobj_rs("HR_SAIDA"))
																				
																			End If
																			
																		%>" /></td>
																		<td>&nbsp;<%
																					
																			If Not IsNull(vobj_rs("HR_ENTRADA")) And Not IsNull(vobj_rs("HR_SAIDA")) Then
																						
																				vint_Minutos = DateDiff("n", CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rs("HR_SAIDA"))))
																				Response.Write converterMinutoParaHora(vint_Minutos)
																				vint_MinutosTotal = vint_MinutosTotal + vint_Minutos
																						
																			End If
																			
																			%></td>
																		<td><%Call CriarComboProjetoUsuario("cmbComboProjetoUsuarioAltera", vobj_rs("ID_PROJETO"), Empty, vstr_IdUsuario)%></td>
																		<td><%Call CriarComboAtividade("cmbComboAtividadeAltera", vobj_rs("ID_ATIVIDADE"), Empty, Empty)%></td>
																		
																	</tr>
																	
																	
																	
																	<%
																	
																	contadorClass = contadorClass + 1
																	
																	vobj_rs.MoveNext
																Loop
																%>
																	
															<COLGROUP />
																<col align="center" width="25" />
																<col align="center" width="70" />
																<col align="center" width="70" />
																<col align="center" width="70" />
																<col align="center" width="80" />
																<col align="center" width="190" />
																<col align="center" width="70" />
																			
																<TR class="Cabecalho">
																	<th>&nbsp;</th>
																	<th>Total</th>
																	<th>&nbsp;</th>
																	<th>&nbsp;</th>
																	<th>&nbsp;<%=converterMinutoParaHora(vint_MinutosTotal)%></th>
																	<th>&nbsp;</th>
																	<th>&nbsp;</th>
																</TR>
																<tr>
																	<td colspan="2"><br /><input type="button" value="Alterar" onclick="alterar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" id=button2 name=button2></td>
																</tr>
																
																<%
																			
															End If
																	
															vobj_rs.Close
															Set vobj_rs = Nothing
															Set vobj_command = Nothing
														
														End If
													End If
															
												%>
											</table>
										</td>
									</tr>
								</table>
							</fieldset>
						</td>
					</tr>
				</TABLE>
			

<fieldset style="LEFT: 0px; WIDTH: 600px">
<legend> <b>Manutenção das Observações</b></legend>
 <%
 
 Dim vobj_rss
 Dim vobj_commandd
 Dim vobj_conexaoo
 Dim vstr_Obs
 vstr_Obs            = Request.Form("txtObs")
 
 Set vobj_commandd = Server.CreateObject("ADODB.Command")
 Set vobj_commandd.ActiveConnection = vobj_conexao
 
 vobj_commandd.CommandType				= adCmdStoredProc
 vobj_commandd.CommandText				= "Consultaobs"
 vobj_commandd.Parameters.Refresh
 vobj_commandd.Parameters.Append vobj_commandd.CreateParameter("param1",adChar, adParamInput, 10, vstr_IdUsuario)
 vobj_commandd.Parameters.Append vobj_commandd.CreateParameter("param2",adDate, adParamInput,, vstr_DtData)
 
 Set vobj_rss = vobj_commandd.Execute
 
													
Do While Not vobj_rss.EOF															
 %>

Obs&nbsp;:<input class="TextBox" type="text" name="txtObs" size="95" maxlength="140" id="Text" value="<%= vobj_rss("Obs") %>" />
<br /><p></p>


<input type="button" value="Alterar" onclick="alterar2();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" id=button3 name=button1>

<%

vobj_rss.MoveNext
Loop 
vobj_rss.Close

Set vobj_rss = Nothing
Set vobj_commandd = Nothing
 
 
 %>
</fieldset>

</form>
		</td>
	</tr>
</table>
<!-- #include file = "../includes/LayoutEnd.asp" -->


<%


' Função desenvolvida para fazer o tratamento do
' formulário de dados.
Private Function ValidarForm()
	
	Dim vobj_commandFaixaHora
	Dim vobj_rsFaixaHora
									
	Dim vstr_HoraEntrada
	Dim vstr_MinutoEntrada
	Dim vstr_HoraSaida
	Dim vstr_MinutoSaida
	
	Dim vstr_HoraEntradaAtual
	Dim vstr_MinutoEntradaAtual
	
	Dim vstr_HoraSaidaAtual
	Dim vstr_MinutoSaidaAtual
	
	Dim vstr_DtIncio
	Dim vstr_DtFinal
	
	If vstr_Executar = "REGISTRAR" Then
		
		' Tratamento de campos do formulário. =============================
		
		If Trim(vstr_IdProjetoUsuario) = "" Then
			
			Call AddErro("Projeto", "Favor, selecionar um projeto.")
		End If
		
		If Trim(vint_IdAtividade) = "" Then
			
			Call AddErro("Atividade", "Favor, selecionar uma atividade.")
		End If
		
		If Trim(vstr_DtData) = "" Then
			
			Call AddErro("Data", "Favor, digite uma data.")
			
		Else
			
			If Not isDate(vstr_DtData) Then
				
				Call AddErro("Data", "Favor, digite uma data válida. Ex 01/01/1900")
				
			Else

				' vstr_DtIncio = DateSerial(vstr_DsAno, vint_NmMes, 1)
				vstr_DtIncio = "01/" & vint_NmMes & "/" & vstr_DsAno
				' vstr_DtFinal = DateSerial(vstr_DsAno, vint_NmMes, Cint(GetUltimoDiaMes(vint_NmMes, vstr_DsAno)))
				vstr_DtFinal = Cint(GetUltimoDiaMes(vint_NmMes, vstr_DsAno)) & "/" & vint_NmMes & "/" & vstr_DsAno
				

				If Not Len(vstr_DtData) = 10 then
					
					Call AddErro("Data", "Favor, digite uma data válida. Ex 01/01/1900")
					
				Else

					'If cDate(vstr_DtData) < cDate(vstr_DtIncio) Or cDate(vstr_DtData) > cDate(vstr_DtFinal) Then
					If cDate(vstr_DtData) < cDate(vstr_DtIncio) Or cDate(vstr_DtData) > vstr_DtFinal Then
						
						Call AddErro("Data", "Favor, digite uma data que seja do mês de " & DescricaoMes(vint_NmMes) & " de " & vstr_DsAno)
						
					End If
					
				End If
				
				If Trim(vstr_HrEntrada) = "" Then
					
					Call AddErro("HrEntrada", "Favor, digite uma hora de entrada.")
					
				Else
					
					If Not isDate(vstr_HrEntrada) Then
						
						Call AddErro("HrEntrada", "Favor, digite uma hora de entrada válida. Ex 03:15.")
						
					Else
						
						If Not isDate(vstr_HrSaida) And Not Trim(vstr_HrSaida) = ""  Then
								
							Call AddErro("HrSaida", "Favor, digite uma hora de saída válida. Ex 03:15.")
								
						Else
							
							If Not Trim(vstr_HrSaida) = "" Then
								
								If cDate(vstr_HrSaida) < cDate(vstr_HrEntrada) Then
										
									Call AddErro("HoraSaidaMenorEntrada", "Erro. Favor, digitar uma hora de saída que seja maior que a de entrada.")
										
								End If
								
								If cDate(vstr_HrEntrada) = cDate(vstr_HrSaida)Then
									
									Call AddErro("HoraSaidaMenorEntrada", "Erro. Favor, digitar uma hora de entrada diferente da hora de saída.")
									
								End If
							
							End If
								
							Set vobj_commandFaixaHora = Server.CreateObject("ADODB.Command")
							Set vobj_commandFaixaHora.ActiveConnection = vobj_conexao
								
								
							vobj_commandFaixaHora.CommandType				= adCmdStoredProc
							vobj_commandFaixaHora.CommandText				= "consultaFaixaHora"
							vobj_commandFaixaHora.Parameters.Refresh
								
							vobj_commandFaixaHora.Parameters.Append vobj_commandFaixaHora.CreateParameter("param1",adChar, adParamInput, 10, vstr_IdUsuario)
							' vobj_commandFaixaHora.Parameters.Append vobj_commandFaixaHora.CreateParameter("param2",adDate, adParamInput, , ConverterDataParaSQL(vstr_DtData))
							vobj_commandFaixaHora.Parameters.Append vobj_commandFaixaHora.CreateParameter("param2",adDate, adParamInput, , vstr_DtData)
							
							Set vobj_rsFaixaHora = vobj_commandFaixaHora.Execute
							' ---------------------------------------------------------------------
								
							If Not vobj_rsFaixaHora.EOF Then
									
									
								vstr_HoraEntradaAtual = cInt(Hour(vstr_HrEntrada))
								vstr_MinutoEntradaAtual = cInt(Minute(vstr_HrEntrada))
									
								'vstr_HoraSaidaAtual =  cInt(Hour(vstr_HrSaida))
								'vstr_MinutoSaidaAtual = cInt(Minute(vstr_HrSaida))
									
								Do While Not vobj_rsFaixaHora.EOF
										
									If IsNull(vobj_rsFaixaHora("HR_SAIDA")) And Trim(vstr_HrSaida) = "" Then
											
										Call AddErro("HoraEntradaSemSaida", "Não foi possível a inclusão, pois já há um registro faltando hora de saída, sendo permitido somente um.")
											
										vobj_rsFaixaHora.MoveLast
											
									Else
											
										'vstr_HoraEntrada = cInt(Hour(DesencriptaString(vobj_rsFaixaHora("HR_ENTRADA"))))
										'vstr_MinutoEntrada = cInt(Minute(DesencriptaString(vobj_rsFaixaHora("HR_ENTRADA"))))
											
										'vstr_HoraSaida = cInt(Hour(DesencriptaString(vobj_rsFaixaHora("HR_SAIDA"))))
										'vstr_MinutoSaida = cInt(Minute(DesencriptaString(vobj_rsFaixaHora("HR_SAIDA"))))
										
										If IsNull(vobj_rsFaixaHora("HR_SAIDA")) Then
											
											If cDate(vstr_HrEntrada) <= cDate(DesencriptaString(vobj_rsFaixaHora("HR_ENTRADA"))) And cDate(vstr_HrSaida) >= cDate(DesencriptaString(vobj_rsFaixaHora("HR_ENTRADA"))) Then
												
												Call AddErro("HoraEntradaFaixaHora", "Não foi possível a inclusão, essa hora já foi inclusa no sistema.")
														
												vobj_rsFaixaHora.MoveLast
											Else
												
												If cDate(vstr_HrEntrada) > cDate(DesencriptaString(vobj_rsFaixaHora("HR_ENTRADA"))) Then
													
													Call AddErro("HoraEntradaFaixaHora", "Não foi possível a inclusão, há um registro faltando hora de saída que é menor que a hora de saída inserida.")
													
													vobj_rsFaixaHora.MoveLast
													
												End If
												
											End IF
										Else
											If cDate(vstr_HrEntrada) >= cDate(DesencriptaString(vobj_rsFaixaHora("HR_ENTRADA"))) And cDate(vstr_HrEntrada) <= cDate(DesencriptaString(vobj_rsFaixaHora("HR_SAIDA"))) Then
												
												Call AddErro("HoraEntradaFaixaHora", "Não foi possível a inclusão, essa hora já foi inclusa no sistema.")
												
												vobj_rsFaixaHora.MoveLast
														
											Else
												
												If Not Trim(vstr_HrSaida) = "" Then
													
													If cDate(vstr_HrSaida) >= cDate(DesencriptaString(vobj_rsFaixaHora("HR_ENTRADA"))) And cDate(vstr_HrSaida) <= cDate(DesencriptaString(vobj_rsFaixaHora("HR_SAIDA"))) Then
														
														Call AddErro("HoraEntradaFaixaHora", "Não foi possível a inclusão, essa hora já foi inclusa no sistema.")
														
													End If
													
												Else
													
													If cDate(vstr_HrEntrada) < cDate(DesencriptaString(vobj_rsFaixaHora("HR_ENTRADA"))) Then
														
														Call AddErro("HoraEntradaFaixaHora", "Não foi possível a inclusão, hora entrada não pode ser menor que as horas no sistema quando não há hora de saída")
														
														vobj_rsFaixaHora.MoveLast
														
													End If
													
												End If
											End If
										End if
									End If
										
									vint_FlTipo = cInt(vobj_rsFaixaHora("FL_TIPO")) + 1
										
									vobj_rsFaixaHora.MoveNext
								Loop
							Else
									
								vint_FlTipo = 1
									
							End If
						
							vobj_rsFaixaHora.Close
							Set vobj_rsFaixaHora = Nothing
							Set vobj_commandFaixaHora = Nothing
								
						End If
					End If
				End If
			End If
		End If
	End If
	
	
	If vstr_Executar = "ALTERAR" Then
		
		
		Dim varr_FlTipoAux
		Dim varr_HrEntradaAux
		Dim varr_HrSaidaAux
		
		Dim vint_ContAux
		Dim vint_ContAux2
		
		varr_FlTipoAux = Split(Request.Form("hdnFlTipoAltera"), ",")
		varr_HrEntradaAux = Split(Request.Form("txtHrEntradaAltera"), ",")
		varr_HrSaidaAux = Split(Request.Form("txtHrSaidaAltera"), ",")
		
		
		For vint_ContAux = LBound(varr_FlTipoAux) To UBound(varr_FlTipoAux)
			
			For vint_ContAux2 = LBound(varr_FlTipo) To UBound(varr_FlTipo)
				
				If Not varr_FlTipoAux(vint_ContAux) = varr_FlTipo(vint_ContAux2) Then
					
					If Trim(varr_HrEntradaAux(vint_ContAux)) = "" Then
						
						Call AddErro("HoraEntrada", "Erro, é obrigatório a hora de entrada em todos os campos de alteração")
						
						vint_ContAux = UBound(varr_FlTipoAux)
						vint_ContAux2 = UBound(varr_FlTipoAux)
					Else
						If Trim(varr_HrEntrada(vint_ContAux2)) = "" Then
							
							Call AddErro("HoraEntrada", "Erro, é obrigatório a hora de entrada em todos os campos de alteração")
							
							vint_ContAux = UBound(varr_FlTipoAux)
							vint_ContAux2 = UBound(varr_FlTipoAux)
							
						Else
							
							If Not isDate(varr_HrEntradaAux(vint_ContAux)) Then
								
								Call AddErro("HoraEntrada", "Erro, Favor, digite uma hora válida. Ex 01:00")
								
								vint_ContAux = UBound(varr_FlTipoAux)
								vint_ContAux2 = UBound(varr_FlTipoAux)
							Else
								
								If Not isDate(varr_HrEntrada(vint_ContAux2)) Then
									
									Call AddErro("HoraEntrada", "Erro, Favor, digite uma hora válida. Ex 01:00")
									
									vint_ContAux = UBound(varr_FlTipoAux)
									vint_ContAux2 = UBound(varr_FlTipoAux)
									
								Else
									
									If Not isDate(varr_HrSaidaAux(vint_ContAux)) And Not Trim(varr_HrSaidaAux(vint_ContAux)) = "" Then
												
										Call AddErro("HoraEntrada", "Erro, Favor, digite uma hora válida. Ex 01:00")
												
										vint_ContAux = UBound(varr_FlTipoAux)
										vint_ContAux2 = UBound(varr_FlTipoAux)
												
									Else
										
										If Not isDate(varr_HrSaida(vint_ContAux2)) And Not Trim(varr_HrSaida(vint_ContAux2)) = ""  Then
											
											Call AddErro("HoraEntrada", "Erro, Favor, digite uma hora válida. Ex 01:00")
											
											vint_ContAux = UBound(varr_FlTipoAux)
											vint_ContAux2 = UBound(varr_FlTipoAux)
										Else
											
											If Trim(varr_IdProjetoUsuario(vint_ContAux)) = "" Then
												
												Call AddErro("ProjetoUsuario", "Favor, selecione um projeto.")
												
												vint_ContAux = UBound(varr_FlTipoAux)
												vint_ContAux2 = UBound(varr_FlTipoAux)
												
											Else
												
												If Trim(varr_IdAtividade(vint_ContAux)) = "" Then
													
													Call AddErro("Atividade", "Favor, selecione uma Atividade")
													
													vint_ContAux = UBound(varr_FlTipoAux)
													vint_ContAux2 = UBound(varr_FlTipoAux)
													
												Else
													
													If Not Trim(varr_HrSaida(vint_ContAux2)) = "" Then
															
														If cDate(varr_HrEntradaAux(vint_ContAux)) >= cDate(varr_HrEntrada(vint_ContAux2)) And cDate(varr_HrEntradaAux(vint_ContAux)) <= cDate(varr_HrSaida(vint_ContAux2)) Then
																
															Call AddErro("HoraEntrada", "Não foi possível a inclusão, digite faixa de horas diferentes.")
																
															vint_ContAux = UBound(varr_FlTipoAux)
															vint_ContAux2 = UBound(varr_FlTipoAux)
																
														Else
																
															If Not Trim(varr_HrSaidaAux(vint_ContAux)) = "" Then
																	
																If cDate(varr_HrSaidaAux(vint_ContAux)) >= cDate(varr_HrEntrada(vint_ContAux2)) And cDate(varr_HrSaidaAux(vint_ContAux)) <= cDate(varr_HrSaida(vint_ContAux2)) Then
																		
																	Call AddErro("HoraEntrada", "Não foi possível a inclusão, digite faixa de horas diferentes.")
																		
																	vint_ContAux = UBound(varr_FlTipoAux)
																	vint_ContAux2 = UBound(varr_FlTipoAux)
																		
																End If
															End if
														End If
															
													Else
															
														If Trim(varr_HrSaidaAux(vint_ContAux)) = "" Then
																
															Call AddErro("HoraEntrada", "Não foi possível a inclusão, somente a ultima hora registrada no dia não é obrigatória.\nFavor, preencher as demais.")
																
															vint_ContAux = UBound(varr_FlTipoAux)
															vint_ContAux2 = UBound(varr_FlTipoAux)
																
														Else
																
															If cDate(varr_HrEntradaAux(vint_ContAux)) <= cDate(varr_HrEntrada(vint_ContAux2)) And cDate(varr_HrSaidaAux(vint_ContAux)) >= cDate(varr_HrEntrada(vint_ContAux2)) Then
																	
																Call AddErro("HoraEntrada", "Não foi possível a inclusão, digite faixas de horas diferentes.")
																	
																vint_ContAux = UBound(varr_FlTipoAux)
																vint_ContAux2 = UBound(varr_FlTipoAux)
																	
															Else
																
																If cDate(varr_HrEntradaAux(vint_ContAux)) >= cDate(varr_HrEntrada(vint_ContAux2)) Then
																		
																	Call AddErro("HoraEntrada", "Não foi possível a inclusão, somente a ultima hora registrada no dia não é obrigatória.\nFavor, preencher as demais.2")
																		
																	vint_ContAux = UBound(varr_FlTipoAux)
																	vint_ContAux2 = UBound(varr_FlTipoAux)
																		
																Else
																		
																	If cDate(varr_HrSaidaAux(vint_ContAux)) >= cDate(varr_HrEntrada(vint_ContAux2)) Then
																			
																		Call AddErro("HoraEntrada", "Não foi possível a inclusão, somente a ultima hora registrada no dia não é obrigatória.\nFavor, preencher as demais.3")
																			
																		vint_ContAux = UBound(varr_FlTipoAux)
																		vint_ContAux2 = UBound(varr_FlTipoAux)
																			
																	End If
																End If
															End If
														End If
													End If
												End If
											End If
										End If
									End If
										
								End If
							End If
						End If
					End If
				End If
			Next
		Next
		
	End If
	
	
	If vstr_Executar = "EXCLUIR" Then
		
		vstr_DtData = Request.Form("txtDsData")
		vint_FlTipo = Request.Form("hdnFlTipo")
		
		If Trim(vstr_DtData) = "" Then
			
			Call AddErro("DtData", "Erro inesperado, não foi possível realizar a exclusão.")
			
		Else
			
			If Not IsDate(vstr_DtData) Then
				
				Call AddErro("DtData", "Erro inesperado, não foi possível realizar a exclusão.")
				
			End if
			
		End If
		
		
		If Trim(vstr_IdUsuario) = "" Then
			
			Call AddErro("IdUsuario", "Erro inesperado, não foi possível realizar a exclusão.")
			
		End If
		
		If Trim(vint_FlTipo) = "" Then
			
			Call AddErro("Tipo", "Erro inesperado, não foi possível realizar a exclusão.")
			
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
<SELECT name="<%=pstr_Nome%>" onChange="<%=pstr_onChange%>" class="TextBox">
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
	<SELECT name="<%=pstr_Nome%>" onChange="<%=pstr_onChange%>" class="TextBox" >
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



Private Sub OrganizarDiaHora(pstr_IdUsuario, pstr_DtData)
		
		If Not Trim(pstr_IdUsuario ) = "" And Not Trim(pstr_DtData) = "" Then
		
		Dim vint_UltimoTipo
		Dim varr_HoraLancado()
		Dim varr_AuxOrdenacao()
		Dim vint_ContadorRegistroRS
		
		vint_UltimoTipo = 0
		
		' Declaração de variáveis locais.
		Dim vobj_command
		Dim vobj_rs
		
		
		' ---------------------------------------------------------------------
		' Selecionando todos os Registros
		' ---------------------------------------------------------------------
		Set vobj_command = Server.CreateObject("ADODB.Command")
		Set vobj_command.ActiveConnection = vobj_conexao
	
	
		vobj_command.CommandType				= adCmdStoredProc
		vobj_command.CommandText				= "consultaHoraDiaOrganiza"
		vobj_command.Parameters.Refresh
	
		vobj_command.Parameters.Append vobj_command.CreateParameter("param1", adChar, adParamInput, 10, pstr_IdUsuario)
		' vobj_command.Parameters.Append vobj_command.CreateParameter("param2", adDate, adParamInput,, converterDataParaSQL(pstr_DtData))
		vobj_command.Parameters.Append vobj_command.CreateParameter("param2", adDate, adParamInput,, pstr_DtData)
	
		Set vobj_rs = vobj_command.Execute
			
		If Not vobj_rs.EOF Then
				
			Dim vint_Contador
			Dim vint_ContadorAux
				
			vint_ContadorRegistroRS = CInt(vobj_rs.RecordCount)
				
			Redim Preserve varr_HoraLancado(vint_ContadorRegistroRS - 1, 2)
			Redim Preserve varr_AuxOrdenacao(vint_ContadorRegistroRS, 2)
				
			' Loop de todos os registros encontrados.
			For vint_Contador = 0 To vint_ContadorRegistroRS - 1
					
				varr_HoraLancado(vint_Contador, 0) = vobj_rs("FL_TIPO")
				varr_HoraLancado(vint_Contador, 1) = DesencriptaString(vobj_rs("HR_ENTRADA"))
				varr_HoraLancado(vint_Contador, 2) = Replace(DesencriptaString(vobj_rs("HR_ENTRADA")),":","")
					
				If vint_UltimoTipo < cInt(varr_HoraLancado(vint_Contador, 0)) Then
						
					vint_UltimoTipo = cInt(varr_HoraLancado(vint_Contador, 0))
						
				End If
					
				vobj_rs.MoveNext
			Next
				
			For vint_Contador = 1 To vint_ContadorRegistroRS
					
				varr_AuxOrdenacao(vint_Contador - 1, 0) = ""
				varr_AuxOrdenacao(vint_Contador - 1, 1) = ""
				varr_AuxOrdenacao(vint_Contador - 1, 2) = 2400 + vint_Contador
					
			Next
				
				
			For vint_Contador = 0 To vint_ContadorRegistroRS - 1
				
				For vint_ContadorAux = vint_ContadorRegistroRS - 1 To 0 Step -1
						
					If cInt(varr_HoraLancado(vint_Contador , 2)) < cInt(varr_AuxOrdenacao(vint_ContadorAux, 2)) Then
							
							
						varr_AuxOrdenacao(vint_ContadorAux + 1, 0) = varr_AuxOrdenacao(vint_ContadorAux, 0)
						varr_AuxOrdenacao(vint_ContadorAux + 1, 1) = varr_AuxOrdenacao(vint_ContadorAux, 1)
						varr_AuxOrdenacao(vint_ContadorAux + 1, 2) = varr_AuxOrdenacao(vint_ContadorAux, 2)
							
						varr_AuxOrdenacao(vint_ContadorAux, 0) = varr_HoraLancado(vint_Contador, 0)
						varr_AuxOrdenacao(vint_ContadorAux, 1) = varr_HoraLancado(vint_Contador, 1)
						varr_AuxOrdenacao(vint_ContadorAux, 2) = varr_HoraLancado(vint_Contador, 2)
							
					Else
							
						vint_ContadorAux = 0
							
					End If
						
				Next
					
			Next
			
			vint_UltimoTipo = cInt(vint_UltimoTipo) + 1
			
			For vint_Contador = 0 To vint_ContadorRegistroRS - 1
				
				
				vobj_command.CommandType				= adCmdStoredProc
				vobj_command.CommandText				= "alteraHoraDiaOrganiza"
				vobj_command.Parameters.Refresh
				
				vobj_command.Parameters.Append vobj_command.CreateParameter("param1", adChar, adParamInput, 10, pstr_IdUsuario)
				' vobj_command.Parameters.Append vobj_command.CreateParameter("param2", adDate, adParamInput,, converterDataParaSQL(pstr_DtData))
				vobj_command.Parameters.Append vobj_command.CreateParameter("param2", adDate, adParamInput,, pstr_DtData)
				vobj_command.Parameters.Append vobj_command.CreateParameter("param3", adInteger, adParamInput,, cInt(varr_AuxOrdenacao(vint_Contador, 0)))
				varr_AuxOrdenacao(vint_Contador, 0) = cInt(vint_UltimoTipo) + vint_Contador
				vobj_command.Parameters.Append vobj_command.CreateParameter("param4", adInteger, adParamInput,, cInt(vint_UltimoTipo) + vint_Contador)
				
				Call vobj_command.Execute
				
			Next
			
			For vint_Contador = 0 To vint_ContadorRegistroRS - 1
				
				
				vobj_command.CommandType				= adCmdStoredProc
				vobj_command.CommandText				= "alteraHoraDiaOrganiza"
				vobj_command.Parameters.Refresh
				
				vobj_command.Parameters.Append vobj_command.CreateParameter("param1", adChar, adParamInput, 10, pstr_IdUsuario)
				' vobj_command.Parameters.Append vobj_command.CreateParameter("param2", adDate, adParamInput,, converterDataParaSQL(pstr_DtData))
				vobj_command.Parameters.Append vobj_command.CreateParameter("param2", adDate, adParamInput,, pstr_DtData)
				vobj_command.Parameters.Append vobj_command.CreateParameter("param3", adInteger, adParamInput,, cInt(varr_AuxOrdenacao(vint_Contador, 0)))
				vobj_command.Parameters.Append vobj_command.CreateParameter("param4", adInteger, adParamInput,, vint_Contador + 1)
				
				Call vobj_command.Execute
				
			Next
			
		End If
			
		vobj_rs.Close
		Set vobj_rs = Nothing
		Set vobj_command = Nothing
		
	End If
End Sub
'response.write "erro>:" & err.description

%>

<!-- #include file = "../includes/CloseConnection.asp" -->