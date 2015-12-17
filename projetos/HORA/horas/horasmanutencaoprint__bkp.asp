<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- #include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->

<%

If	Not Session("sint_TipoUsuario") = "1" And Not Session("sint_TipoUsuario") = "2"  Then
	
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


' para está página.
vstr_Operacao		= Request.Form("pstr_Operacao")
vstr_Processar		= Request.Form("hdnProcessar")
vstr_Executar		= Request.Form("hdnExecutar")


If Session("sint_TipoUsuario") = "2" Then
	
	vstr_IdUsuario			= Session("sstr_IdUsuario")
	
Else
	
	vstr_IdUsuario			= Request.Form("cmbComboUsuario")
	
End If

vint_NmMes				= Cint(Request.Form("cmbComboMes"))
vstr_DsAno				= Cint(Request.Form("txtDsAno"))

If vstr_IdUsuario = "" Then
	
	Response.Redirect("horasrelatoriofiltro.asp")
	
ElseIf vint_NmMes = "" Then
	
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
<script type="text/javascript" src="js/horasmanutencaoprint.js"></script>
</HEAD>
<BODY>
<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<form name="thisForm" action="horasmanutencao.asp" method="post">
				
				<input type="hidden" name="cmbComboUsuario" value="<%=vstr_IdUsuario%>" />
				<input type="hidden" name="cmbComboMes" value="<%=vint_NmMes%>" />
				<input type="hidden" name="txtDsAno" value="<%=vstr_DsAno%>" />
				
				<i><b class="TituloPagina">Horas no Mês</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 595px;">
							<legend>
							   <b>Horário</b>
							</legend>
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabResultado" id="tabResultado">
								<tr>
									<td colspan="6">
										<strong>Colaborador: <%
											
												Response.Write NomeUsuario(vstr_IdUsuario)
											
											%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Mês: <%
												
												Response.Write DescricaoMes(vint_NmMes) & "/" & vstr_DsAno
										
										%></strong>
									</td>
								</tr>
								<COLGROUP />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<tr class="Cabecalho">
									<th>Data</th>
									<th>Entrada</th>
									<th>Saída Almoço</th>
									<th>Entrada Almoço</th>
									<th>Saída</th>
									<th>Total</th>
									<th>Horas Acum.</th>
									<th>Horário entrada</th>
									<th>Atrasos</th>
								
								</tr>
							</table>
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabResultado" id="tabResultado">
								<COLGROUP />
     							<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<%
								Dim vstr_DataConsulta
								Dim vint_NmUltimoDia
								
								vint_NmUltimoDia = Cint(GetUltimoDiaMes(vint_NmMes, vstr_DsAno))
								
								Dim vint_ContadorDia
								Dim vint_Minutos
								Dim vint_MinutosTotal
								Dim	vint_tot
								Dim res 
								Dim soma
								Dim difer
								Dim conv
								Dim conv1
								Dim x
								Dim atrasos
								Dim ttdev
								
								vint_tot = 480 
								vint_MinutosTotal = 0
								
								Dim contadorClass
								contadorClass = 0
								
								For vint_ContadorDia = 1 To vint_NmUltimoDia
								
									Dim vint_ContadorRegistroRS
									
									vstr_DataConsulta = DateSerial(vstr_DsAno,vint_NmMes,vint_ContadorDia)
									
									' Declaração de variáveis locais.
									Dim vobj_command
									Dim vobj_rs
									
									
									' ---------------------------------------------------------------------
									' Selecionando todos os Registros
									' ---------------------------------------------------------------------
									Set vobj_command = Server.CreateObject("ADODB.Command")
									Set vobj_command.ActiveConnection = vobj_conexao
									
									
									vobj_command.CommandType				= adCmdStoredProc
									vobj_command.CommandText				= "consultaRelatorioHora"
									vobj_command.Parameters.Refresh
									
									vobj_command.Parameters.Append vobj_command.CreateParameter("param1",adChar, adParamInput, 10, vstr_IdUsuario)
									vobj_command.Parameters.Append vobj_command.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(vstr_DataConsulta))
									
									
									
									Set vobj_rs = vobj_command.Execute
									
									If Not vobj_rs.EOF Then
										
										vint_Minutos = 0
										
										vint_ContadorRegistroRS = CInt(vobj_rs.RecordCount)
										
										x = DesencriptaString(vobj_rs("HR_ENTRADA"))
										If vint_ContadorRegistroRS = 2 Then
												
											%>
												
											<tr class="tr<%=contadorClass Mod 2 %>">
												<td>&nbsp;<%=converterDataParaHtml(vobj_rs("DT_DATA"))%></td>
												<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_ENTRADA"))%></td>
												<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_SAIDA"))%></td>
												<%
													
												vint_Minutos = DateDiff("n", CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rs("HR_SAIDA"))))
													
												vobj_rs.MoveNext
													
												%>
												<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_ENTRADA"))%></td>
												<td>&nbsp;<%
												
													If Not IsNull(vobj_rs("HR_SAIDA")) Then
														
														Response.Write DesencriptaString(vobj_rs("HR_SAIDA"))
														
														vint_Minutos = vint_Minutos + DateDiff("n", CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rs("HR_SAIDA"))))
														
													Else
														
														Response.Write ""
														
													End If
													
												   vint_MinutosTotal = vint_MinutosTotal + vint_Minutos
                       					           res = vint_Minutos - vint_tot
			                  			           soma = soma + res
			                  			           
			                  			           
                                                 conv1 = Minute(converterHoraParaSQL(vobj_rs("horasen")))
                                                 conv =  Minute(converterHoraParaSQL(x))
                                                 difer = conv - conv1
                                                 ttdev = vint_Minutos - difer
                                                 
                                                 
                                                 
                                                 atrasos = atrasos + difer          
												 	
												
												
												%></td>
												<td>&nbsp;<% If (vint_Minutos < vint_tot) Then
												
												Response.Write "<font color='red'>"      
					                     		Response.Write converterMinutoParaHora(vint_Minutos)
					                     		
					                     		Else
					                     		   
					                     		Response.Write converterMinutoParaHora(vint_Minutos)
					                     		  
					                     		    End If
												
												
												
												%></td>
					                     		
					                     		
					                     		
					                     		<td>&nbsp;<%=converterMinutoParaHora(res)%></td>
					                     		<td>&nbsp;<%= converterHoraParaSQL(vobj_rs("horasen"))  %></td>
					                     		<td>&nbsp;<%= converterMinutoParaHora(difer)  %></td>
					                    
					                    
					                    
					                    
					                     	<!--	<td>&nbsp;<% If (ttdev < vint_tot) Then
					                     		Response.Write "<font color='red'>"      
					                     		Response.Write  converterMinutoParaHora(ttdev)
					                     		   
					                     		   Else
					                     		   
					                     		   Response.Write converterMinutoParaHora(ttdev)
					                     		  
					                     		    End If
					                     		 
					                     		  %></td> -->
											</tr>
												
											<%
											
											contadorClass = contadorClass + 1
												
										ElseIf vint_ContadorRegistroRS = 1 Then
											
											%>
											
											<tr class="tr<%=contadorClass Mod 2 %>">
												<td>&nbsp;<%=converterDataParaHtml(vobj_rs("DT_DATA"))%></td>
												<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_ENTRADA"))%></td>
												<td>&nbsp;<%
												
													If Not IsNull(vobj_rs("HR_SAIDA")) Then
														
														Response.Write DesencriptaString(vobj_rs("HR_SAIDA"))
														
														vint_Minutos = DateDiff("n", CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rs("HR_SAIDA"))))
														
													Else
														
														Response.Write "-"
														
													End If
													
												%></td>
												<td>&nbsp;</td>
												<td>&nbsp;</td>
												<td>&nbsp;<%
													
													vint_MinutosTotal = vint_MinutosTotal + vint_Minutos
													
													If vint_Minutos = 0 Then
														
														Response.Write ""
														
													Else
														Response.Write converterMinutoParaHora(vint_Minutos)
														
													End If
													
												%></td>
											</tr>
											
											<%
											
											contadorClass = contadorClass + 1
											
										Else
											
											%>
											
											<tr class="tr<%=contadorClass Mod 2 %>">
												<td>&nbsp;<%=converterDataParaHtml(vobj_rs("DT_DATA"))%></td>
												<td>&nbsp;<%=DesencriptaString(vobj_rs("HR_ENTRADA"))%></td>
												<td>&nbsp;</td>
												<td>&nbsp;</td>
												<%
												
												Dim vint_ContadorAux
												
												For vint_ContadorAux = 1 To vint_ContadorRegistroRS
													
													If Not IsNull(vobj_rs("HR_SAIDA"))Then
														
														vint_Minutos = vint_Minutos + DateDiff("n", CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rs("HR_SAIDA"))))
														
													End If
													
													If Not vint_ContadorAux = vint_ContadorRegistroRS Then
														
														vobj_rs.MoveNext
														
													End If
													
												Next
												
												%>
												<td>&nbsp;<%
												
													If Not IsNull(vobj_rs("HR_SAIDA")) Then
														
														Response.Write DesencriptaString(vobj_rs("HR_SAIDA"))
														
													Else
														
														Response.Write ""
														
													End If
													
													vint_MinutosTotal = vint_MinutosTotal + vint_Minutos
													
												%></td>
												<td>&nbsp;<%=converterMinutoParaHora(vint_Minutos)%></td>
											</tr>
											
											<%
											
											contadorClass = contadorClass + 1
											
										End If
									End If
								
									vobj_rs.Close
									Set vobj_rs = Nothing
									Set vobj_command = Nothing
									
								Next
								
								%>
								<tr class="Cabecalho">
									<th>TOTAL</th>
									<th>&nbsp;</th>
									<th>&nbsp;</th>
									<th>&nbsp;Dias trabalh.</th>
									<th><%= contadorClass  %></th>
									<th>&nbsp;<%=converterMinutoParaHora(vint_MinutosTotal)%></th>
									<th>&nbsp;<%=converterMinutoParaHora(soma)%></th>
									<th>Total Atrasos</th>
									<th>&nbsp;<%=converterMinutoParaHora(atrasos)%></th>
									
								</tr>
								
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
									<td><input type="button" value="Imprimir" onclick="imprimir();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Imprime relatório"></td>
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