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

Dim vstr_DsAno 
Dim vobj_commandRegistro

vstr_DsAno = request("Data")


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
				
				
				<i><b class="TituloPagina">Horas no Mês</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 840px;">
							<legend>
							   <b>Horário</b>
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
									<th>Saída</th>
									<th>Total</th>
									<th>Horas <br /> Acum.</th>
									<th>Horário entrada</th>
									<th>Atrasos</th>
									
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
							vobj_command.CommandText					= "ConsultaHoraDiaria"
							
							//vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("param1", adChar, adParamInput, 10, vstr_DsAno)
							
							' Cria o recordset e posiciona a páginação do recordset.
							Set vobj_rsRegistro = vobj_command.Execute
							
							
							' Verifica se registros foram encontrados.
							If Not vobj_rsRegistro.EOF Then
								
								Dim contadorClass
								contadorClass = 0
								
								' Loop de todos os registros cadastrados
								' no banco de dados.
								Do While Not vobj_rsRegistro.EOF
									
									%>
									
									
									<tr style="cursor: hand" class="tr<%=contadorClass Mod 2 %>">
										
										<%
											
											If CBool(vobj_rsRegistro("FL_ATIVO")) = True Then
												
												%>
												
												<td onclick="desativar('<%=vobj_rsRegistro("ID_USUARIO")%>');" title="Desativar Usuário"><img src="<%=getBaseLink("/images/star_on.gif")%>"></td>
												
												<%
											Else
												
												%>
												
												<td onclick="ativar('<%=vobj_rsRegistro("ID_USUARIO")%>');" title="Ativar Usuário"><img src="<%=getBaseLink("/images/star_off.gif")%>"></td>
												
												<%
													
											End If
												
											%>
										
										<td onclick="alterar('<%=Trim(vobj_rsRegistro("ID_USUARIO"))%>');"><%=Trim(vobj_rsRegistro("ID_USUARIO"))%></td>
										<td onclick="alterar('<%=Trim(vobj_rsRegistro("ID_USUARIO"))%>');"><%=Trim(vobj_rsRegistro("DS_USUARIO"))%></td>
										<td onclick="alterar('<%=Trim(vobj_rsRegistro("ID_USUARIO"))%>');"><%=vobj_rsRegistro("DS_CPF")%></td>
										<td onclick="alterar('<%=Trim(vobj_rsRegistro("ID_USUARIO"))%>');"><%=vobj_rsRegistro("DS_RG")%></td>
										<td onclick="alterar('<%=Trim(vobj_rsRegistro("ID_USUARIO"))%>');"><%
										
											If vobj_rsRegistro("FL_PERFIL") = 1 Then
												
												Response.Write "Administrador"
											Else
												
												Response.Write "Colaborador"
											End If
											
										%></td>
									</tr>
									
									
									<%
									
									contadorClass = contadorClass + 1
									
									' Move para o próximo registro do loop.
									vobj_rsRegistro.MoveNext
								Loop
							Else
								%>
								
								
								<tr>
									<td colspan="6">Nenhum registro foi encontrado!!!</td>
								</tr>
								
								
								<%
							End If
							
							vobj_rsRegistro.Close
							Set vobj_rsRegistro = Nothing
							Set vobj_command = Nothing
							
							%>
      
                                    </table>
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabResultado" id="tabResultado">
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
												
								
								</td>
									
												<td>&nbsp;</td>
					                     		<td>&nbsp;</td>
					                     		<td>&nbsp;</td>
					                    
												
											
											
											</tr>
												
																				
											<tr style="cursor: hand" class="">
												
												
												</td>
											</tr>
											
											
											
											
												
												
												
									</td>
													
												</td>
												<td >&nbsp;<% %> <td>
												
											</tr>
											
											<%
											
											
									
									
								
								%>
								
								<tr class="Cabecalho">
                                    <th>TOTAL&nbsp;</th>
									<th>&nbsp;</th>
									<th>&nbsp;Dias trabalh.</th>
									<th>&nbsp;</th>
									<th>&nbsp;</th>
									<th>&nbsp;</th>
									<th>&nbsp;</th>
									<th>&nbsp;Total Atrasos</th>
									<th>&nbsp;</th>
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
<!-- #include file = "../includes/CloseConnection.asp" -->