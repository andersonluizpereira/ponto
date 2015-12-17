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

' para está página.
vstr_Operacao		= Request.Form("pstr_Operacao")
vstr_Executar		= Request.Form("hdnExecutar")



vstr_IdUsuario			= Request.Form("cmbComboUsuario")
vint_NmMes				= Cint(Request.Form("cmbComboMes"))
vstr_DsAno				= Cint(Request.Form("txtDsAno"))

If vstr_IdUsuario = "" Then
	
	Response.Redirect("horasmanutencaofiltro.asp")
	
ElseIf vint_NmMes = "" Then
	
	Response.Redirect("horasmanutencaofiltro.asp")

ElseIf vstr_DsAno = "" Or Not Len(vstr_DsAno) = 4 Or Not IsNumeric(vstr_DsAno) Then

	Response.Redirect("horasmanutencaofiltro.asp")
	
End If

' Ajusta hora do almoço para casos que tem mais de 7 horas trabalhas e não marcou pausa obrigatoria no sistema.
' Obs. Sabado não inclu essa regra.
If vstr_Executar = "AJUSTAR_HORA" Then
	
	vstr_DtData = Request.Form("hdnDtData")
	
	'Verufucando se a data não final de semana, se for final de semana a regra não é necessária.
	If Not DiaDaSemana(vstr_DtData) = "Domingo" And Not DiaDaSemana(vstr_DtData) = "Sábado" Then
		
		Dim vint_MinutoDiferencaHora
		Dim vint_MinutoDiferencaAlmoco
		
		Dim vint_ContadorRegistroRSAjustarHora
		
		vint_MinutoDiferencaHora = 0
		vint_MinutoDiferencaAlmoco = 0
		
		Dim vobj_commandAjustarHora
		Dim vobj_rsAjustarHora
		
		
		' ---------------------------------------------------------------------
		' Selecionando todos os Registros
		' ---------------------------------------------------------------------
		Set vobj_commandAjustarHora = Server.CreateObject("ADODB.Command")
		Set vobj_commandAjustarHora.ActiveConnection = vobj_conexao
		
		
		vobj_commandAjustarHora.CommandType				= adCmdStoredProc
		vobj_commandAjustarHora.CommandText				= "consultaRelatorioHora"
		vobj_commandAjustarHora.Parameters.Refresh
		
		vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param1",adChar, adParamInput, 10, vstr_IdUsuario)
		vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param2",adDate, adParamInput,, converterDataParaSQL(vstr_DtData))
		
		
		Set vobj_rsAjustarHora = vobj_commandAjustarHora.Execute
		
		If Not vobj_rsAjustarHora.EOF Then
			
			vint_ContadorRegistroRSAjustarHora = CInt(vobj_rsAjustarHora.RecordCount)
			
			Dim vstr_DtEntradaAux
			Dim vstr_DtSaidaAux
			Dim vstr_DtEntradaAux2
			Dim vstr_DtSaidaAux2
			
			' Recebe nova hora, podendo ser na ida pro almoço ou volta do almoço.
			' Função principal, é manter 1 hora de alomoço entre entrada e volta do almoço
			DIm vstr_HrNova
			
			Dim vint_NmTipoAux1
			Dim vint_NmTipoAux2
			
			Dim vint_IdAtividadesAux
			Dim vstr_idProjetoAux
			
			' Verificando se tem dois registros de pontos. 2 é o mais normal de haver.
			If vint_ContadorRegistroRSAjustarHora = 2 Then
				
				vstr_DtEntradaAux = DesencriptaString(vobj_rsAjustarHora("HR_ENTRADA"))
				vstr_DtSaidaAux = DesencriptaString(vobj_rsAjustarHora("HR_SAIDA"))
				vint_NmTipoAux1 = vobj_rsAjustarHora("FL_TIPO")
				
				vint_MinutoDiferencaHora = DateDiff("n", CDate(vstr_DtEntradaAux), CDate(vstr_DtSaidaAux))
				
				vobj_rsAjustarHora.MoveNext
				
				vstr_DtEntradaAux2 = DesencriptaString(vobj_rsAjustarHora("HR_ENTRADA"))
				vstr_DtSaidaAux2 = DesencriptaString(vobj_rsAjustarHora("HR_SAIDA"))
				vint_NmTipoAux2 = vobj_rsAjustarHora("FL_TIPO")
				
				If Not IsNull(vobj_rsAjustarHora("HR_SAIDA")) Then
					
					vint_MinutoDiferencaHora = vint_MinutoDiferencaHora + DateDiff("n", CDate(vstr_DtEntradaAux2), CDate(vstr_DtSaidaAux2))
					
				End If
				
				vint_MinutoDiferencaAlmoco = DateDiff("n", CDate(vstr_DtSaidaAux), CDate(vstr_DtEntradaAux2))
				
				If Cint(vint_MinutoDiferencaHora) >= 420 And Cint(vint_MinutoDiferencaAlmoco) < 60 Then
					
					If DateDiff("n", CDate(vstr_DtEntradaAux), CDate(vstr_DtSaidaAux)) > 240 Then
						
						
						vstr_HrNova = DateAdd("n", cInt(vint_MinutoDiferencaAlmoco) - 60, cDate(vstr_DtSaidaAux))
						
						vobj_commandAjustarHora.CommandType				= adCmdStoredProc
						vobj_commandAjustarHora.CommandText				= "alteraHoraDiaAjusteAlmocoSaida"
						vobj_commandAjustarHora.Parameters.Refresh
						
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param1", adChar, adParamInput, 10, vstr_IdUsuario)
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param2", adDate, adParamInput,, converterDataParaSQL(vstr_DtData))
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param3", adInteger, adParamInput,, vint_NmTipoAux1)
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param4", adChar, adParamInput, 5, EncriptaString(converterHoraParaSQL(cDate(vstr_HrNova))))
						
						Call vobj_commandAjustarHora.Execute
						
					Else
						
						vstr_HrNova = DateAdd("n", 60 - cInt(vint_MinutoDiferencaAlmoco), cDate(vstr_DtEntradaAux2))
						
						vobj_commandAjustarHora.CommandType				= adCmdStoredProc
						vobj_commandAjustarHora.CommandText				= "alteraHoraDiaAjusteAlmocoEntrada"
						vobj_commandAjustarHora.Parameters.Refresh
						
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param1", adChar, adParamInput, 10, vstr_IdUsuario)
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param2", adDate, adParamInput,, converterDataParaSQL(vstr_DtData))
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param3", adInteger, adParamInput,, vint_NmTipoAux2)
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param4", adChar, adParamInput, 5, EncriptaString(converterHoraParaSQL(cDate(vstr_HrNova))))
						
						Call vobj_commandAjustarHora.Execute
						
					End If
					
				End If
				
			ElseIf vint_ContadorRegistroRSAjustarHora = 1 Then
				
				vstr_DtEntradaAux = DesencriptaString(vobj_rsAjustarHora("HR_ENTRADA"))
				vstr_DtSaidaAux = DesencriptaString(vobj_rsAjustarHora("HR_SAIDA"))
				vint_NmTipoAux1 = vobj_rsAjustarHora("FL_TIPO")
				
				vint_IdAtividadesAux = vobj_rsAjustarHora("ID_ATIVIDADE")
				vstr_idProjetoAux = vobj_rsAjustarHora("ID_PROJETO")
				
				If Not IsNull(vobj_rsAjustarHora("HR_SAIDA")) Then
					
					vint_MinutoDiferencaHora = DateDiff("n", CDate(vstr_DtEntradaAux), CDate(vstr_DtSaidaAux))
					
					If Cint(vint_MinutoDiferencaHora) >= 420 Then
						
						Dim vint_MinutoMetadeDia
						
						vint_MinutoMetadeDia = 0 
						
							
						' Pegando quantidade de minutos equivalentes a metade do dia
						vint_MinutoMetadeDia = cInt(vint_MinutoDiferencaHora) \ 2
							
						' Atribuindo nova saida, mesma da metade do da e retirando tempo pausa.
						vstr_DtSaidaAux = DateAdd("n", cInt(vint_MinutoMetadeDia) - 60, cDate(vstr_DtEntradaAux))
							
						' Nova hora de entrada após almoço, com uma pausa de uma hora da saida
						vstr_DtEntradaAux2 = DateAdd("n", 60, cDate(vstr_DtSaidaAux))
							
						' Nova hora de saida após almoço, com um possível acrecimo de um minuto que pode sobrar na divisao do dia.
						vstr_DtSaidaAux2 = DateAdd("n", cInt(vint_MinutoMetadeDia)+ (cInt(vint_MinutoDiferencaHora) Mod 2), cDate(vstr_DtEntradaAux2))
							
							
						' alterando para metade do dia antes pausa de uma hora.
						vobj_commandAjustarHora.CommandType				= adCmdStoredProc
						vobj_commandAjustarHora.CommandText				= "alteraHoraDiaAjusteAlmocoEntradaSaida"
						vobj_commandAjustarHora.Parameters.Refresh
							
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param1", adChar, adParamInput, 10, vstr_IdUsuario)
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param2", adDate, adParamInput,, converterDataParaSQL(vstr_DtData))
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param3", adInteger, adParamInput,, vint_NmTipoAux1)
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param4", adChar, adParamInput, 5, EncriptaString(converterHoraParaSQL(cDate(vstr_DtEntradaAux))))
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param4", adChar, adParamInput, 5, EncriptaString(converterHoraParaSQL(cDate(vstr_DtSaidaAux))))
							
						Call vobj_commandAjustarHora.Execute
							
						' incluindo segunda parte do dia apos pausa de uma hora.
						vobj_commandAjustarHora.CommandType				= adCmdStoredProc
						vobj_commandAjustarHora.CommandText				= "incluiHoraDiaAjusteAposAlmoco"
						vobj_commandAjustarHora.Parameters.Refresh
							
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param1", adChar, adParamInput, 10, vstr_IdUsuario)
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param2", adDate, adParamInput,, converterDataParaSQL(vstr_DtData))
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param3", adInteger, adParamInput,, Cint(vint_NmTipoAux1)+1)
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param4", adChar, adParamInput, 10, vstr_idProjetoAux)
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param5", adInteger, adParamInput,, vint_IdAtividadesAux)
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param6", adChar, adParamInput, 5, EncriptaString(converterHoraParaSQL(cDate(vstr_DtEntradaAux2))))
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param7", adChar, adParamInput, 5, EncriptaString(converterHoraParaSQL(cDate(vstr_DtSaidaAux2))))
						vobj_commandAjustarHora.Parameters.Append vobj_commandAjustarHora.CreateParameter("param8", adDate, adParamInput,, converterDataParaSQL(Date()))
							
						Call vobj_commandAjustarHora.Execute
							
					End If
				End If
			End If
		End If							
		
		vobj_rsAjustarHora.Close
		Set vobj_rsAjustarHora = Nothing
		Set vobj_commandAjustarHora = Nothing
		
	End If
End If

%>

<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/horasmanutencao.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<form name="thisForm" action="horasmanutencaolancamento.asp" method="post">
				
				<input type="hidden" name="hdnProcessar" />
				<input type="hidden" name="pstr_Operacao" />
				<input type="hidden" name="hdnExecutar" />
				<input type="hidden" name="hdnDtData" />
				<input type="hidden" name="cmbComboUsuario" value="<%=vstr_IdUsuario%>" />
				<input type="hidden" name="cmbComboMes" value="<%=vint_NmMes%>" />
				<input type="hidden" name="txtDsAno" value="<%=vstr_DsAno%>" />
				<input type="hidden" name="hdnBloqueaData" />
				
				<i><b class="TituloPagina">Horas no Mês</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 1055px;">
							<legend>
							   <b>Horário</b>
							</legend>
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabResultado" id="tabResultado">
								<tr>
									<td colspan="8">
										<strong>Colaborador: <%
												
												Response.Write NomeUsuario(vstr_IdUsuario)
											
											%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Mês:<%
												
												Response.Write Ucase(DescricaoMes(vint_NmMes)) & "/" & vstr_DsAno 
										%></strong>
									</td>
								</tr>
								<COLGROUP />
								<col align="middle" width="30" />
								<col align="middle" width="25" />
								<col align="middle" width="100" />
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
									<th></th>
									<th></th>
									<th>Data</th>
									<th>Entrada</th>
									<th>Saída <br />Almoço</th>
									<th>Entrada Almoço</th>
									<th>Saída</th>
									<th>Total</th>
									<th>Horas <br /> Acum.</th>
									<th>Horário entrada</th>
									<th>Atrasos</th>
									<th>Observações</th>
								</tr>
							</table>
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabResultado" id="tabResultado">
								<COLGROUP />
								<col align="middle" width="30" />
								<col align="middle" width="25" />
								<col align="middle" width="100" />
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
								
								
							    Dim vint_Minutos
								Dim vint_MinutosAlmoco
								Dim vint_MinutosTotal
								Dim vint_ContadorDia
								Dim	vint_tot
								Dim res 
								Dim soma
								Dim difer
								Dim conv
								Dim conv1
								Dim x
								Dim atrasos
								Dim ttdev
								Dim tt
								Dim tt1
								Dim y
								Dim per1
								
								vint_tot = 480 
								vint_MinutosTotal = 0
								
								Dim contadorClass
								contadorClass = 0
								vint_MinutosTotal = 0
								
								
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
										
										'Contador pra fazer o efeito zebrado.
										
										
										vint_Minutos = 0
										
										vint_ContadorRegistroRS = CInt(vobj_rs.RecordCount)
										
										Dim vstr_DataAux
										Dim vstr_DataEntradaAux
										Dim vstr_DataSaidaAux
										Dim vstr_DataSaidaAux2
										x = DesencriptaString(vobj_rs("HR_ENTRADA"))	
										
										If vint_ContadorRegistroRS = 2 Then
											
											
											vstr_DataAux = converterDataParaHtml(vobj_rs("DT_DATA"))
											vstr_DataEntradaAux = DesencriptaString(vobj_rs("HR_ENTRADA"))
											vstr_DataSaidaAux = DesencriptaString(vobj_rs("HR_SAIDA"))
											
											vint_Minutos = DateDiff("n", CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rs("HR_SAIDA"))))
												
											vobj_rs.MoveNext
											
											If Not IsNull(vobj_rs("HR_SAIDA")) Then
														
												vint_MinutosAlmoco = DateDiff("n", CDate(vstr_DataSaidaAux), CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))))
												
												vint_Minutos = vint_Minutos + DateDiff("n", CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rs("HR_SAIDA"))))
												
											End If
											
											%>
											
											<tr style="cursor: hand" class="tr<%=contadorClass Mod 2 %>">
												<td><input type="checkbox" name="chkAlteraLote" class="TextBox" value="<%=vstr_DataAux%>" /></td>
												<%
												
												'Verificando se a data não final de semana, se for final de semana a regra não é necessária.
												If Not DiaDaSemana(vstr_DataConsulta) = "Domingo" And Not DiaDaSemana(vstr_DataConsulta) = "Sábado" Then
												
													If Cint(vint_Minutos) >= 420 And Cint(vint_MinutosAlmoco) < 60 Then
														
														%>
														
														<td onclick="ajustarAlmoco(thisForm, '<%=vstr_DataAux%>');">
															<img src="../images/warning.gif" title="Ajusta Hora de Almoço" />
														</td>
														
														<%
														
													ElseIf Cint(vint_Minutos) >= 360 And Cint(vint_MinutosAlmoco) < 60 Then
														
														%>
														
														<td onclick="alterar('<%=vstr_DataAux%>');">
															<img src="../images/warning.gif" title="Rever Horas" />
														</td>
														
														<%
														
													Else
														
														%>
														
														<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;</td>
														
														<%
														
													End If 
													
												Else
													
													%>
													
													<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;</td>
													
													<%
													
												End If
												
												%>
												
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%=vstr_DataAux%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">1&nbsp;<%=vstr_DataEntradaAux%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">1&nbsp;<%=vstr_DataSaidaAux%></td>
												
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%=DesencriptaString(vobj_rs("HR_ENTRADA"))%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%
												
													If Not IsNull(vobj_rs("HR_SAIDA")) Then
														
														Response.Write DesencriptaString(vobj_rs("HR_SAIDA"))
														
													Else
														
														Response.Write ""
														
													End If
													
												vint_MinutosTotal = vint_MinutosTotal + vint_Minutos
												
												 conv1 = Minute(converterHoraParaSQL(vobj_rs("horasen")))
                                                 conv =  Minute(converterHoraParaSQL(x))
                                                 
                                                 tt = Hour(converterHoraParaSQL(vobj_rs("horasen")))
                                                 y = Hour(converterHoraParaSQL(x))
                                                
                                                 tt1 = Hour(converterHoraParaSQL(x))
                                                                                                
                                           
                                                If((y-tt)=-5) Then
                                                    
                                                 difer = conv-(60*5)
                                                
                                                 Else If((y-tt)=-4) Then
                                                    
                                                 difer = conv-(60*4)
                                                 
                                                 Else If((y-tt)=-3) Then
                                                    
                                                 difer = conv-(60*3)
                                                 
                                                 Else If((y-tt)=-2) Then
                                                    
                                                 difer = conv-(60*2)
                                                  
                                                  Else If((y-tt)=-1) Then
                                                    
                                                 difer = conv-60
                                                                                            
                                                  Else If ((y-tt)=5) Then                                            
                                                
                                                   difer = conv+(60*5)
                                                  
                                                  Else If ((y-tt)=4) Then                                            
                                                
                                                   difer = conv+(60*4)
                                                   
                                                  Else If ((y-tt)=3) Then                                            
                                                
                                                   difer = conv+(60*3)
                                                  
                                                  Else If ((y-tt)=2) Then                                            
                                                
                                                   difer = conv+(60*2)
                                                   
                                                  Else If ((y-tt)=1) Then                                            
                                                
                                                   difer = conv+(60*1)
                                                   
                                                 
                                                 Else 
                                                 difer = (conv - conv1)
                                                 
                                                 
                                                 
                                                 
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 
                                                 
                                                 
                                                 If(tt=6) Then
                                                 
                                                    vint_tot = 360
                                                 
                                                 
                                                Else If(tt=7) Then
                                                 
                                                    vint_tot = 420 
                                                 
                                                Else If(tt=8 ) Then
                                                 
                                                    vint_tot = 480 
                                                 
                                                 Else If(tt=9) Then
                                                 
                                                    vint_tot = 540
                                                    
                                                 Else If(tt=10) Then
                                                 
                                                    vint_tot = 540
                                                 
                                                 Else If(tt=11) Then
                                                 
                                                    vint_tot = 660
                                                    
                                                 Else If(tt=12) Then
                                                 
                                                    vint_tot = 720
                                                    
                                                 Else 
                                                    
                                                   vint_tot = 780
                                                 
                                                 End If
                                                 End If
                                                 End If
                                                 End If
                                                 End If
                                                 End If
                                                 End If
                                                 
                                                 
                                                 per1 = Hour(converterHoraParaSQL(vobj_rs("per")))
                                                 
                                                 If(per1=6) Then
                                                 
                                                 vint_tot = 360
                                                 
   else if (per1=8) Then
  vint_tot = 480
                                              
   End IF
   End IF
                                                 
                                                 res = vint_Minutos - vint_tot
                       					         
                       					         
                       					         soma = soma + res
                       					        
                                                 atrasos = atrasos + difer          
												
													
													
												%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;
												
												<% If (vint_Minutos < vint_tot) Then
												
												Response.Write "<font color='red'>"      
					                     		Response.Write converterMinutoParaHora(vint_Minutos)
					                     		
					                     		Else
					                     		   
					                     		Response.Write converterMinutoParaHora(vint_Minutos)
					                     		
					                     		
					                     		
					                     		  
					                     		    End If
												
												%></td>
												<td>&nbsp;<%= converterMinutoParaHora(res)%></td>
					                     		<td>&nbsp;<%= converterHoraParaSQL(vobj_rs("horasen"))  %></td>
					                     		<td>&nbsp;<%= converterMinutoParaHora(difer) %></td>
					                     		<td>&nbsp;<%= vobj_rs("Obs") %></td>
					                    
												
											
											
											</tr>
												
											<%
											
											contadorClass = contadorClass + 1
											
										ElseIf vint_ContadorRegistroRS = 1 Then
											
											vstr_DataAux = converterDataParaHtml(vobj_rs("DT_DATA"))
											vstr_DataEntradaAux = DesencriptaString(vobj_rs("HR_ENTRADA"))
											
											%>
											
											<tr style="cursor: hand" class="tr<%=contadorClass Mod 2 %>">
												<td><input type="checkbox" name="chkAlteraLote" class="TextBox" value="<%=vstr_DataAux%>" /></td>
												
												
												<%
													
													If Not IsNull(vobj_rs("HR_SAIDA")) Then
														
														vstr_DataSaidaAux = DesencriptaString(vobj_rs("HR_SAIDA"))
														
														vint_Minutos = DateDiff("n", CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rs("HR_SAIDA"))))
														
														'Verificando se a data não final de semana, se for final de semana a regra não é necessária.
														If Not DiaDaSemana(vstr_DataConsulta) = "Domingo" And Not DiaDaSemana(vstr_DataConsulta) = "Sábado" Then
															
															If Cint(vint_Minutos) >= 420 Then
																
																%>
																
																<td onclick="ajustarAlmoco(thisForm, '<%=vstr_DataAux%>');">
																	<img src="../images/warning.gif" title="Ajusta Hora de Almoço" />
																</td>
																
																<%
																
															ElseIf Cint(vint_Minutos) >= 360 Then
																
																%>
																
																<td onclick="alterar('<%=vstr_DataAux%>');">
																	<img src="../images/warning.gif" title="Rever Horas" />
																</td>
																
																<%
																
															Else
																
																%>
																
																<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;</td>
																
																<%
																
															End If
															
														Else
															
															%>
															
															<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;</td>
															
															<%
															
														End If
														
													Else
														
														%>
														
														<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;</td>
														
														<%
														
													End If
													
												%>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%=vstr_DataAux%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%=vstr_DataEntradaAux%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%
												
													If Not IsNull(vobj_rs("HR_SAIDA")) Then
														
														Response.Write vstr_DataSaidaAux
														
													Else
														
														Response.Write ""
														
													End If
													
												%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;</td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;</td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%
													
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
											
											vstr_DataAux = converterDataParaHtml(vobj_rs("DT_DATA"))
											vstr_DataEntradaAux = DesencriptaString(vobj_rs("HR_ENTRADA"))
											vstr_DataSaidaAux = DesencriptaString(vobj_rs("HR_SAIDA"))
											
											%>
											
											<tr style="cursor: hand" class="tr<%=contadorClass Mod 2 %>">
												<td><input type="checkbox" name="chkAlteraLote" class="TextBox" value="<%=vstr_DataAux%>" /></td>
												<td onclick="alterar('<%=vstr_DataAux%>');"><%
												
												Dim vint_ContadorAux
												
												For vint_ContadorAux = 1 To vint_ContadorRegistroRS
													
													If Not IsNull(vobj_rs("HR_SAIDA"))Then
														
														vint_Minutos = vint_Minutos + DateDiff("n", CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rs("HR_SAIDA"))))
														
													End If
													
													If Not vint_ContadorAux = vint_ContadorRegistroRS Then
														
														vstr_DataSaidaAux2 = DesencriptaString(vobj_rs("HR_SAIDA"))
														
														vobj_rs.MoveNext
																																																								
														vint_MinutosAlmoco = vint_MinutosAlmoco + DateDiff("n", CDate(vstr_DataSaidaAux2), CDate(DesencriptaString(vobj_rs("HR_ENTRADA"))))
														
													End If
													
												Next
												
												
												If vint_Minutos => 360 And vint_MinutosAlmoco < 60 Then
													
													%>
													
													<img src="../images/warning.gif" title="Rever Horas" />
													
													<%
													
												End If
												
												%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%=vstr_DataAux%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%=vstr_DataEntradaAux%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;</td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;</td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%
												
													If Not IsNull(vobj_rs("HR_SAIDA")) Then
														
														Response.Write DesencriptaString(vobj_rs("HR_SAIDA"))
														
													Else
														
														Response.Write ""
														
													End If
													
													vint_MinutosTotal = vint_MinutosTotal + vint_Minutos
													
												%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%=converterMinutoParaHora(vint_Minutos)%></td>
												
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
                                    <th>&nbsp;</th>
									<th>&nbsp;</th>
									<th>TOTAL&nbsp;</th>
									<th>&nbsp;</th>
									<th>&nbsp;Dias trabalh.</th>
									<th>&nbsp;<%= contadorClass  %></th>
									<th>&nbsp;</th>
									<th>&nbsp;<%=converterMinutoParaHora(vint_MinutosTotal)%></th>
									<th>&nbsp;<%=converterMinutoParaHora(soma)%></th>
									<th>&nbsp;Total Atrasos</th>
									<th>&nbsp;<% If(atrasos<0) Then
									
									atrasos=0 
									Response.Write converterMinutoParaHora(atrasos)
									
									Else
									
									Response.Write converterMinutoParaHora(atrasos)
									
									End If
									
									%></th>
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
									<td><input type="button" value="Tela Impressão" onclick="imprimir();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Envia para a tela de impressão"></td>
									<td><input type="button" value="Incluir" onclick="incluir();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Inclui uma nova data"></td>
									<td><input type="button" value="Retornar" onclick="voltar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Retornar a tela anterior"></td>
									<td><input type="button" value="Alterar Lote" onclick='alterarLote(thisForm);' class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Alterar em Lote"></td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</TABLE>

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