<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- S#include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->

<%

If	Not Session("sboo_fladministrador") = True Then
	
	Response.Redirect getBaseLink("/horas/horaslancamento.asp")
	
End If
Dim vint_Dia
Dim vint_NmMes
Dim vstr_DsAno
Dim vstr_DtData
Dim vstr_IdUsuario

vstr_IdUsuario = Request.Form("hdnIdRegistro")
vint_Dia= Request.Form("Data")

%>


<!-- #include file = "../includes/LayoutBegin.asp" -->
<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<form name="thisForm" method="post">
				
				
				<i><b class="TituloPagina">Horas no M�s</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 840px;">
							<legend>
							   <b>Hor�rio</b>
							</legend>
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabResultado" id="tabResultado">
								<tr>
									<td colspan="8">
										<strong>Dia: <%= vstr_DsAno
												
																							
											%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
										</strong>
									</td>
								</tr>
								<COLGROUP />
								<col align="middle" width="50" />
								<col align="middle" width="50" />
								<col align="middle" width="200" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<tr class="Cabecalho">
									<th>Data</th>
									<th>User ID</th>
									<th>Nome</th>
									<th>Entrada</th>
									<th>Sa�da</th>
									<th>Total</th>
									<th>Horas <br /> Acum.</th>
									<th>Hor�rio entrada</th>
									<th>Atrasos</th>
								<%
								
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
																
								Dim contadorClass
								contadorClass = 0
								vint_MinutosTotal = 0
								
								
								Dim vobj_rsRegistro
	                            Dim vobj_commandRegistro
	                            Dim vstr_DataAux
								Dim vstr_DataEntradaAux
								Dim vstr_DataSaidaAux
							    Dim vstr_DataSaidaAux2							                
	                            Dim vstr_DataConsulta
								Dim vint_NmUltimoDia
								
								vint_NmUltimoDia = Cint(GetUltimoDiaMes(vint_NmMes, vstr_DsAno))
								    
	                             For vint_ContadorDia = 1 To vint_NmUltimoDia
								
									Dim vint_ContadorRegistroRS
									
									vstr_DataConsulta = DateSerial(vstr_DsAno,vint_NmMes,vint_ContadorDia)
									
									' Declara��o de vari�veis locais.
									Dim vobj_command
								    Dim vobj_rs
									
									
									' ---------------------------------------------------------------------
									' Selecionando todos os Registros
									' ---------------------------------------------------------------------
									Set vobj_command = Server.CreateObject("ADODB.Command")
									Set vobj_command.ActiveConnection = vobj_conexao
									
	                                vobj_command.CommandType = adCmdStoredProc
  	                           	    vobj_command.CommandText = "ConsultaHoraDiaria"
			                        vobj_command.Parameters.Append vobj_command.CreateParameter("paramters",adDate, adParamInput,, converterDataParaSQL(vint_Dia))
	                                ' ---------------------------------------------------------------------
	                                Set vobj_rs = vobj_command.Execute
	                                
	                                If Not vobj_rs.EOF Then
										
										'Contador pra fazer o efeito zebrado.
										
										
										vint_Minutos = 0
										
										vint_ContadorRegistroRS = CInt(vobj_rs.RecordCount)
										
										
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
												
												'Verificando se a data n�o final de semana, se for final de semana a regra n�o � necess�ria.
												If Not DiaDaSemana(vstr_DataConsulta) = "Domingo" And Not DiaDaSemana(vstr_DataConsulta) = "S�bado" Then
												
													If Cint(vint_Minutos) >= 420 And Cint(vint_MinutosAlmoco) < 60 Then
														
														%>
														
														<td onclick="ajustarAlmoco(thisForm, '<%=vstr_DataAux%>');">
															<img src="../images/warning.gif" title="Ajusta Hora de Almo�o" />
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
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%=vstr_DataEntradaAux%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<%=vstr_DataSaidaAux%></td>
												
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
                                                 
                                                 res =    vint_Minutos - vint_tot
                       					         
                       					         
                       					         soma = soma + res
                       					        
                                                 atrasos = atrasos + difer          
												
													
													
												%></td>
												<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;<% If (vint_Minutos < vint_tot) Then
												
												Response.Write "<font color='red'>"      
					                     		Response.Write converterMinutoParaHora(vint_Minutos)
					                     		
					                     		Else
					                     		   
					                     		Response.Write converterMinutoParaHora(vint_Minutos)
					                     		  
					                     		 End If
					                     		    
												
												%></td>
											
											</td>
												<td>&nbsp;<%= converterMinutoParaHora(res)%></td>
					                     		<td>&nbsp;<%= converterHoraParaSQL(vobj_rs("horasen"))  %></td>
					                     		<td>&nbsp;<%= converterMinutoParaHora(difer) %></td>
					                    
												
											
											
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
														
														'Verificando se a data n�o final de semana, se for final de semana a regra n�o � necess�ria.
														If Not DiaDaSemana(vstr_DataConsulta) = "Domingo" And Not DiaDaSemana(vstr_DataConsulta) = "S�bado" Then
																															
																%>
																
												
																
														
														<td onclick="alterar('<%=vstr_DataAux%>');">&nbsp;</td>
														
													
												
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
									<td><input type="button" value="Tela Impress�o" onclick="imprimir();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Envia para a tela de impress�o"></td>
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
<!-- #include file = "../includes/CloseConnection.asp" -->

