<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- #include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->

<%

If	Not Session("sboo_fladministrador") = True AND Not Session("sboo_flmoderador") = True Then
	
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
Dim vint_FlTipo

Dim vstr_DtLote
Dim varrvar_DtLote

Dim vboo_NotValidaAlteracao

' para está página.
vstr_Operacao		= Request.Form("pstr_Operacao")
vstr_Processar		= Request.Form("hdnProcessar")
vstr_Executar		= Request.Form("hdnExecutar")

vstr_IdUsuario		= Request.Form("cmbComboUsuario")
vint_NmMes			= Cint(Request.Form("cmbComboMes"))
vstr_DsAno			= Cint(Request.Form("txtDsAno"))

If Trim(vstr_Operacao) = "V" Then
	
	vstr_DtLote			= Request.Form("chkAlteraLote")	
	varrvar_DtLote		= Split(vstr_DtLote, ",")
	
Else
	
	vstr_DtLote			= Request.Form("hdnDtLote")
	varrvar_DtLote		= Split(vstr_DtLote, ",")
	
End If

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
	
	
	
Else
	
	
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
	
	
	' Analiza a operação a ser executada na página
	' para descobrir o processamento que deve ser feito.
	Select Case vstr_Operacao
		
		
		Case "A"						' Operação de alteração do registro.
			
			Dim varr_IdProjetoUsuario
			Dim varr_IdAtividade
			Dim varr_FlTipo
			Dim varr_HrEntrada
			Dim varr_HrSaida
			Dim varr_DtData
			
			varr_IdProjetoUsuario = Split(Request.Form("cmbComboProjetoUsuarioAltera"), ",")
			varr_IdAtividade = Split(Request.Form("cmbComboAtividadeAltera"), ",")
			varr_FlTipo = Split(Request.Form("hdnFlTipoAltera"), ",")
			varr_HrEntrada = Split(Request.Form("txtHrEntradaAltera"), ",")
			
			If Trim(Request.Form("txtHrSaidaAltera")) = "" Then
			
				varr_HrSaida = Split(" ", ",")
				
			Else
				
				varr_HrSaida = Split(Request.Form("txtHrSaidaAltera"), ",")
				
			End If
			
			varr_DtData = Split(Request.Form("hdnDtData"), ",")
			
			' Verificando se o formulário foi
			' devidamente válidado pelo sistema.
			If ValidarForm = True Then
				
				
				Dim vint_Contadoralteracao
				
				' ---------------------------------------------------------------------
				' Incuindo dados no banco de dados.
				' ---------------------------------------------------------------------
				Set vobj_commandProc = Server.CreateObject("ADODB.Command")
				Set vobj_commandProc.ActiveConnection = vobj_conexao
				
				For vint_Contadoralteracao = LBound(varr_FlTipo) To UBound(varr_FlTipo)
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "alteraRegistroHoraManutencao"
					vobj_commandProc.Parameters.Refresh
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_IdUsuario))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(Trim(varr_DtData(vint_Contadoralteracao))))
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
					' ---------------------------------------------------------------------
					
				Next
				
				Dim  vint_ContadorOrganizaDiaHora
				
				For vint_ContadorOrganizaDiaHora = LBound(varrvar_DtLote) To UBound(varrvar_DtLote)
					
					Call OrganizarDiaHora(vstr_IdUsuario, varrvar_DtLote(vint_ContadorOrganizaDiaHora))
					
				Next
				
				Set vobj_commandProc = Nothing
				
				vstr_Operacao = "A"
			
			Else
				
				' Alteracao não validada, sera mostrado novamete os campos recuperados da Array
				' e não banco.
				vboo_NotValidaAlteracao = True
				
			End If
			
		Case "E"
			
			vstr_DtData = Request.Form("hdnDsData")
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
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(vstr_DtData))
				vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adInteger, adParamInput,, vint_FlTipo)
				
				Call vobj_commandProc.Execute
				Set vobj_commandProc = Nothing
				' ---------------------------------------------------------------------
				
				Call OrganizarDiaHora(vstr_IdUsuario, vstr_DtData)
				
				vstr_Operacao = "A"
				
			End If
				
	End Select
	
End If
' *******************************************************
' FINAL DA ROTINA QUE FAZ O PROCESSAMENTO DOS DADOS
' DO REGISTRO.
' *******************************************************


%>


<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/horasmanutencaoalteracao.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<i><b class="TituloPagina">Manutenção de Horas</b></i><br>
			<br>
			<form name="thisForm" action="horasmanutencaoalteracao.asp" method="post">
						
				<input type="hidden" name="hdnProcessar" value="S">
				<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>">
				<input type="hidden" name="hdnExecutar" />
				<input type="hidden" name="hdnFlTipo" />
				<input type="hidden" name="hdnDsData" />
				<input type="hidden" name="hdnDtLote" value="<%=vstr_DtLote%>" />
				<input type="hidden" name="cmbComboUsuario" value="<%=vstr_IdUsuario%>" />
				<input type="hidden" name="cmbComboMes" value="<%=vint_NmMes%>" />
				<input type="hidden" name="txtDsAno" value="<%=vstr_DsAno%>" />
				
				<TABLE class="font" BORDER="0" CELLSPACING="1" CELLPADDING="1" align="center">
					<tr>
						<td><%=ExibirErros()%></td>
					</tr>
					</tr>
					
					<%
					
					Dim vint_ContDataLote
					
					For	vint_ContDataLote = LBound(varrvar_DtLote) To UBound(varrvar_DtLote)
						
						%>
						
						<tr>	
							<td>
								<fieldset style="LEFT: 0px; WIDTH: 600px">
									<legend>
										<b><%=Trim(varrvar_DtLote(vint_ContDataLote))%></b>
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
														
														If vboo_NotValidaAlteracao = True Then
															
															Dim vint_ContadorValidaAlteracao
															
															contadorClass = 0
															
															For vint_ContadorValidaAlteracao = LBound(varr_FlTipo) To UBound(varr_FlTipo)
																
																If Trim(varrvar_DtLote(vint_ContDataLote)) = Trim(varr_DtData(vint_ContadorValidaAlteracao)) Then
																	
																	%>
																	
																	<tr class="tr<%=contadorClass Mod 2 %>" id="trLinhaRegistro" name="trLinhaRegistro">
																		<td style="cursor: hand" onclick="excluir('<%=Trim(varr_FlTipo(vint_ContadorValidaAlteracao))%>', '<%=Trim(varr_DtData(vint_ContadorValidaAlteracao))%>');" title="Exclui Hora"><img src="../images/star_off.gif" /></td>
																		<td><%=Trim(varr_DtData(vint_ContadorValidaAlteracao))%>
																			<input type="hidden" name="hdnFlTipoAltera" value="<%=Trim(varr_FlTipo(vint_ContadorValidaAlteracao))%>" />
																			<input type="hidden" name="hdnDtData" value="<%=Trim(varr_DtData(vint_ContadorValidaAlteracao))%>" />
																		</td>
																		<td><input class="TextBox" type="text" name="txtHrEntradaAltera" size="3" maxlength="5" value="<%=Trim(varr_HrEntrada(vint_ContadorValidaAlteracao))%>" /></td>
																		<td><input class="TextBox" type="text" name="txtHrSaidaAltera" size="3" maxlength="5" value="<%=Trim(varr_HrSaida(vint_ContadorValidaAlteracao))%>" /></td>
																		<td>&nbsp;</td>
																		<td><%Call CriarComboProjetoUsuario("cmbComboProjetoUsuarioAltera", Trim(varr_IdProjetoUsuario(vint_ContadorValidaAlteracao)), Empty, vstr_IdUsuario)%></td>
																		<td><%Call CriarComboAtividade("cmbComboAtividadeAltera", cInt(varr_IdAtividade(vint_ContadorValidaAlteracao)), Empty, Empty)%></td>
																	</tr>
																	
																	<%
																	
																	contadorClass = contadorClass + 1
																End If
															Next
														Else
															
															' Declaração de variáveis locais.
															Dim vobj_command
															Dim vobj_rs
														
																	
															' ---------------------------------------------------------------------
															' Selecionando todos os Registros
															' ---------------------------------------------------------------------
															Set vobj_command = Server.CreateObject("ADODB.Command")
															Set vobj_command.ActiveConnection = vobj_conexao
																	
															'Contador pra fazer o efeito zebrado.
															Dim contadorClass
															contadorClass = 0
																
																	
															vobj_command.CommandType				= adCmdStoredProc
															vobj_command.CommandText				= "consultaManutencaoLancamento"
															vobj_command.Parameters.Refresh
															vobj_command.Parameters.Append vobj_command.CreateParameter("param1",adChar, adParamInput, 10, vstr_IdUsuario)
															vobj_command.Parameters.Append vobj_command.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(varrvar_DtLote(vint_ContDataLote)))
																	
																	
															Set vobj_rs = vobj_command.Execute
																			
															If Not vobj_rs.EOF Then
																		
																Dim vint_Minutos
																Dim vint_MinutosTotal
																				
																		
																vint_MinutosTotal = 0
																		
																		
																' Loop de todos os registros encontrados.
																Do While Not vobj_rs.EOF
																					
																	%>
																	
																	<tr class="tr<%=contadorClass Mod 2 %>" id="trLinhaRegistro" name="trLinhaRegistro">
																		<td style="cursor: hand" onclick="excluir('<%=vobj_rs("FL_TIPO")%>', '<%=converterDataParaHtml(vobj_rs("DT_DATA"))%>');" title="Exclui Hora"><img src="../images/star_off.gif" /></td>
																		<td><%=converterDataParaHtml(vobj_rs("DT_DATA"))%>
																			<input type="hidden" name="hdnFlTipoAltera" value="<%=vobj_rs("FL_TIPO")%>" />
																			<input type="hidden" name="hdnDtData" value="<%=converterDataParaHtml(vobj_rs("DT_DATA"))%>" />
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
																
															End If
															
															vobj_rs.Close
															Set vobj_rs = Nothing
															Set vobj_command = Nothing
															
														End If
														
														%>
														
												</table>
											</td>
										</tr>
									</table>
								</fieldset>
							</td>
						</tr>
						
						<%
						
					Next
					
					%>
					<tr>
						<td><br /><input type="button" value="Alterar" onclick="alterarLote();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';">
						<input id="btnVoltar" type="button" value="Voltar" onclick='voltar(thisForm);' class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';"></td>
					</tr>
				</TABLE>
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
	
	
	If vstr_Executar = "ALTERAR" Then
		
		
		Dim varr_FlTipoAux
		Dim varr_HrEntradaAux
		Dim varr_HrSaidaAux
		Dim varr_DtDataAux
		
		Dim vint_ContAux
		Dim vint_ContAux2
		
		varr_FlTipoAux = Split(Request.Form("hdnFlTipoAltera"), ",")
		varr_HrEntradaAux = Split(Request.Form("txtHrEntradaAltera"), ",")
		varr_HrSaidaAux = Split(Request.Form("txtHrSaidaAltera"), ",")
		varr_DtDataAux = Split(Request.Form("hdnDtData"), ",")
		
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
													
													If Trim(varr_DtData(vint_ContAux2)) = Trim(varr_DtDataAux(vint_ContAux))Then
														
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
				End If
			Next
		Next
		
	End If
	
	
	If vstr_Executar = "EXCLUIR" Then
		
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
		vobj_command.Parameters.Append vobj_command.CreateParameter("param2", adDate, adParamInput,, converterDataParaSQL(pstr_DtData))
	
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
				vobj_command.Parameters.Append vobj_command.CreateParameter("param2", adDate, adParamInput,, converterDataParaSQL(pstr_DtData))
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
				vobj_command.Parameters.Append vobj_command.CreateParameter("param2", adDate, adParamInput,, converterDataParaSQL(pstr_DtData))
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


%>

<!-- #include file = "../includes/CloseConnection.asp" -->