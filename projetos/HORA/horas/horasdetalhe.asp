<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- #include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->

<%

If	Not Session("sint_TipoUsuario") = "1" And Not Session("sint_TipoUsuario") = "2" And Not Session("sint_TipoUsuario") = "3"  Then
	
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

vstr_DtData			= Request.Form("hdnDtData")

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

%>


<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/horasdetalhe.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<i><b class="TituloPagina">Detalhe de Horas</b></i><br>
			<br>
			<form name="thisForm" action="horasdetalhe.asp" method="post">
						
				<input type="hidden" name="cmbComboUsuario" value="<%=vstr_IdUsuario%>" />
				<input type="hidden" name="cmbComboMes" value="<%=vint_NmMes%>" />
				<input type="hidden" name="txtDsAno" value="<%=vstr_DsAno%>" />
				
						
				<TABLE class="font" BORDER="0" CELLSPACING="1" CELLPADDING="1" align="center">
					<tr>
						<td>
							<fieldset style="LEFT: 0px; WIDTH: 600px">
								<legend>
									<b>Detalhe do Horário Lançado</b>
								</legend>
								<table width="590" height="30" border="0" cellpadding="0" cellspacing="0" class="font">
									<tr>
										<td colspan="6">
											<strong>Colaborador: <%
												
													Response.Write NomeUsuario(vstr_IdUsuario)
												
												%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Mês: <%
													
													Response.Write DescricaoMes(vint_NmMes) & "/" & vstr_DsAno
											
											%></strong>
										</td>
									</tr>
									<tr>
										<td>
											<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabCabec" id="tabCabec">
												<COLGROUP />
													<col align="center" width="70" />
													<col align="center" width="70" />
													<col align="center" width="70" />
													<col align="center" width="80" />
													<col align="center" width="190" />
													<col align="center" width="70" />
													<TR class="Cabecalho">
														<TD>Data</TD>
														<TD>Entrada</TD>
														<TD>Saída</TD>
														<TD>Total</TD>
														<TD>Projeto</TD>
														<TD>Atividade</TD>
													</TR>
												<COLGROUP />
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
													
													Dim contadorClass
													contadorClass = 0
													
													' ---------------------------------------------------------------------
													' Selecionando todos os Registros
													' ---------------------------------------------------------------------
													Set vobj_command = Server.CreateObject("ADODB.Command")
													Set vobj_command.ActiveConnection = vobj_conexao
															
															
													vobj_command.CommandType				= adCmdStoredProc
													vobj_command.CommandText				= "consultaRelatorioLancamento"
													vobj_command.Parameters.Refresh
													
													
													vobj_command.Parameters.Append vobj_command.CreateParameter("param1",adChar, adParamInput, 10, vstr_IdUsuario)
													vobj_command.Parameters.Append vobj_command.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(vstr_DtData))
															
															
													Set vobj_rs = vobj_command.Execute
																	
													If Not vobj_rs.EOF Then
																
														Dim vint_Minutos
														Dim vint_MinutosTotal
																		
																
														vint_MinutosTotal = 0
																
																
														' Loop de todos os registros encontrados.
														Do While Not vobj_rs.EOF
																			
															%>
															
															<tr class="tr<%=contadorClass Mod 2 %>" id="trLinhaRegistro" name="trLinhaRegistro">
																<td><%=converterDataParaHtml(vobj_rs("DT_DATA"))%>
																	<input type="hidden" name="hdnFlTipoAltera" value="<%=vobj_rs("FL_TIPO")%>" />
																</td>
																<td><%=DesencriptaString(vobj_rs("HR_ENTRADA"))%></td>
																<td><%
																			
																	If Not IsNull(vobj_rs("HR_SAIDA")) Then
																				
																		Response.Write DesencriptaString(vobj_rs("HR_SAIDA"))
																				
																	End If
																			
																%></td>
																<td>&nbsp;<%
																					
																	If Not IsNull(vobj_rs("HR_ENTRADA")) And Not IsNull(vobj_rs("HR_SAIDA")) Then
																						
																		vint_Minutos = DateDiff("n", CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rs("HR_SAIDA"))))
																		Response.Write converterMinutoParaHora(vint_Minutos)
																		vint_MinutosTotal = vint_MinutosTotal + vint_Minutos
																						
																	End If
																			
																	%></td>
																<td><%=vobj_rs("ID_PROJETO")%></td>
																<td><%=vobj_rs("DS_ATIVIDADE")%></td>
															</tr>
																	
															<%
																	
															contadorClass = contadorClass + 1
																	
															vobj_rs.MoveNext
														Loop
																
														%>
																	
													<COLGROUP />
														<col align="center" width="70" />
														<col align="center" width="70" />
														<col align="center" width="70" />
														<col align="center" width="80" />
														<col align="center" width="190" />
														<col align="center" width="70" />
																			
														<TR class="Cabecalho">
															<th>Total</th>
															<th>&nbsp;</th>
															<th>&nbsp;</th>
															<th>&nbsp;<%=converterMinutoParaHora(vint_MinutosTotal)%></th>
															<th>&nbsp;</th>
															<th>&nbsp;</th>
														</TR>
														<%
																			
													End If
																	
													vobj_rs.Close
													Set vobj_rs = Nothing
													Set vobj_command = Nothing
														
												%>
											</table>
										</td>
									</tr>
								</table>
							</fieldset>
						</td>
					</tr>
					<tr>
						<td align="middle">
							<table ALIGN="center" BORDER="0" CELLSPACING="1" CELLPADDING="1">
								<tr>
									<td><input type="button" value="Retornar" onclick="voltar(thisForm);" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Retornar a tela anterior"></td>
								</tr>
							</table>
						</td>
					</tr>
				</TABLE>
			</form>
		</td>
	</tr>
</table>

<!-- #include file = "../includes/LayoutEnd.asp" -->

<%

Private Function NomeUsuario(pstr_IdUsuario)
		
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
	vobj_commandRegistro.CommandText					= "consultaUsuario"
	
	vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("param1", adChar, adParamInput, 10, vstr_IdUsuario)
	' ---------------------------------------------------------------------
	
	
	' Cria o objeto recordset com as informações do registro.	
	Set vobj_rsRegistro = vobj_commandRegistro.Execute
	
	If Not vobj_rsRegistro.EOF Then
		
		NomeUsuario = Trim(vobj_rsRegistro("ID_USUARIO")) & "  -  " & Trim(vobj_rsRegistro("DS_USUARIO"))
		
	End If
	
	vobj_rsRegistro.Close
	Set vobj_rsRegistro = Nothing
	Set vobj_commandRegistro = Nothing
		
		
End Function

%>

<!-- #include file = "../includes/CloseConnection.asp" -->