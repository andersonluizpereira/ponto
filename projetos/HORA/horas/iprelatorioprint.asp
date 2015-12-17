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

' Declara��o de vari�veis locais. ==============================================

' Guarda a opera��o que ser� executa nesta tela.
' Obs.: Seus valores podem ser A = Altera��o, I = Inclus�o, V = Visualiza��o.
Dim vstr_Operacao

' Vari�vel flag que indica se a p�gina deve ser 
' processada, apenas disponivel para as opera��es de
' A e I.
Dim vstr_Processar

'Vari�vel de controle e fluxo de acoes
Dim vstr_Executar

' Declara��o de vari�veis utilizadas para armazenar os
' valores dos campos da tela.
Dim vint_NmMes
Dim vstr_DsAno


' para est� p�gina.
vstr_Operacao		= Request.Form("pstr_Operacao")
vstr_Processar		= Request.Form("hdnProcessar")
vstr_Executar		= Request.Form("hdnExecutar")



vint_NmMes				= Cint(Request.Form("cmbComboMes"))
vstr_DsAno				= Cint(Request.Form("txtDsAno"))

If vint_NmMes = "" Then
	
	Response.Redirect("horasrelatoriofiltro.asp")

ElseIf vstr_DsAno = "" Or Not Len(vstr_DsAno) = 4 Or Not IsNumeric(vstr_DsAno) Then

	Response.Redirect("horasrelatoriofiltro.asp")
	
End If

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Stefanini - Controle de Horas</TITLE>
<LINK rel="stylesheet" type="text/css" href="../css/chs.css">
<script type="text/javascript" src="js/iprelatorioprint.js"></script>
</HEAD>
<BODY>
<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<form name="thisForm" action="iprelatorio.asp" method="post">
				
				<input type="hidden" name="cmbComboMes" value="<%=vint_NmMes%>" />
				<input type="hidden" name="txtDsAno" value="<%=vstr_DsAno%>" />
				
				<i><b class="TituloPagina">Horas no M�s</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 595px;">
							<legend>
							   <b>Hor�rio IP</b>
							</legend>
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabResultado" id="tabResultado">
								<tr>
									<td colspan="6">
										<strong>M�s: <%
												
												Response.Write DescricaoMes(vint_NmMes) & "/" & vstr_DsAno
										
										%></strong>
									</td>
								</tr>
								<COLGROUP />
								<col align="middle" width="80" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="120" />
								<col align="middle" width="120" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<tr class="Cabecalho">
									<th>Usu�rio</th>
									<th>Projeto</th>
									<th>Data</th>
									<th>Hora Entrada</th>
									<th>Hora Sa�da</th>
									<th>IP</th>
									<th>Tipo</th>
								</tr>

								
								<%
								
								Dim contadorClass
								contadorClass = 0
									
								Dim vint_ContadorRegistroRS
									
								' Declara��o de vari�veis locais.
								Dim vobj_command
								Dim vobj_rs
								
								Dim vint_NmUltimoDia
								
								vint_NmUltimoDia = Cint(GetUltimoDiaMes(vint_NmMes, vstr_DsAno))	
									
								' ---------------------------------------------------------------------
								' Selecionando todos os Registros
								' ---------------------------------------------------------------------
								Set vobj_command = Server.CreateObject("ADODB.Command")
								Set vobj_command.ActiveConnection = vobj_conexao
									
									
								vobj_command.CommandType				= adCmdStoredProc
								vobj_command.CommandText				= "consultaRelatorioIP"
								vobj_command.Parameters.Refresh
								
								vobj_command.Parameters.Append vobj_command.CreateParameter("param1",adDate, adParamInput,, converterDataParaSQL("01/" & vint_NmMes & "/" & vstr_DsAno))
								vobj_command.Parameters.Append vobj_command.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(vint_NmUltimoDia & "/" & vint_NmMes & "/" & vstr_DsAno))
								
								
								Set vobj_rs = vobj_command.Execute
								
								If Not vobj_rs.EOF Then
									
									vint_ContadorRegistroRS = CInt(vobj_rs.RecordCount)
									
									' Loop de todos os registros encontrados.
									Do While Not vobj_rs.EOF
										
										If Not IsNull(vobj_rs("DS_IP_ENTRADA")) Then
											
											%>
													
											<tr class="tr<%=contadorClass Mod 2 %>">
												<td>&nbsp;<%=Trim(vobj_rs("DS_USUARIO"))%></td>
												<td>&nbsp;<%=Trim(vobj_rs("ID_PROJETO"))%></td>
												<td>&nbsp;<%=converterDataParaHtml(vobj_rs("DT_DATA"))%></td>
												<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_ENTRADA"))%></td>
												<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_SAIDA"))%></td>
												<td>&nbsp;<%=Trim(vobj_rs("DS_IP_ENTRADA"))%></td>
												<td>&nbsp;Entrada</td>
											</tr>
													
											<%
											
											contadorClass = contadorClass + 1
											
											If Not IsNull(vobj_rs("DS_IP_SAIDA")) Then
												
												%>
												
												<tr class="tr<%=contadorClass Mod 2 %>">
													<td>&nbsp;<%=Trim(vobj_rs("DS_USUARIO"))%></td>
													<td>&nbsp;<%=Trim(vobj_rs("ID_PROJETO"))%></td>
													<td>&nbsp;<%=converterDataParaHtml(vobj_rs("DT_DATA"))%></td>
													<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_ENTRADA"))%></td>
													<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_SAIDA"))%></td>
													<td>&nbsp;<%=Trim(vobj_rs("DS_IP_SAIDA"))%></td>
													<td>&nbsp;Sa�da</td>
												</tr>
												
												<%
												
												contadorClass = contadorClass + 1
												
											End If
											
										ElseIF Not IsNull(vobj_rs("DS_IP_SAIDA")) Then
											
											%>
											
											<tr class="tr<%=contadorClass Mod 2 %>">
												<td>&nbsp;<%=Trim(vobj_rs("DS_USUARIO"))%></td>
												<td>&nbsp;<%=Trim(vobj_rs("ID_PROJETO"))%></td>
												<td>&nbsp;<%=converterDataParaHtml(vobj_rs("DT_DATA"))%></td>
												<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_ENTRADA"))%></td>
												<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_SAIDA"))%></td>
												<td>&nbsp;<%=Trim(vobj_rs("DS_IP_SAIDA"))%></td>
												<td>&nbsp;Sa�da</td>
											</tr>
											
											<%
											
											contadorClass = contadorClass + 1
											
										Else
											
											%>
											
											<tr class="tr<%=contadorClass Mod 2 %>">
												<td>&nbsp;<%=Trim(vobj_rs("DS_USUARIO"))%></td>
												<td>&nbsp;<%=Trim(vobj_rs("ID_PROJETO"))%></td>
												<td>&nbsp;<%=converterDataParaHtml(vobj_rs("DT_DATA"))%></td>
												<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_ENTRADA"))%></td>
												<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_SAIDA"))%></td>
												<td>&nbsp;</td>
												<td>&nbsp;</td>
											</tr>
											
											<%
											
											contadorClass = contadorClass + 1
										End If
										
										vobj_rs.MoveNext
									Loop
									
								End If
								
								vobj_rs.Close
								Set vobj_rs = Nothing
								Set vobj_command = Nothing
									
								%>
								
							</table>
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
									<td><input type="button" value="Volta" onclick="voltar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Retonar a tela anterior"></td>
									<td><input type="button" value="Imprimir" onclick="imprimir();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Imprime relat�rio"></td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</TABLE>
</BODY>
</HTML>

<!-- #include file = "../includes/CloseConnection.asp" -->