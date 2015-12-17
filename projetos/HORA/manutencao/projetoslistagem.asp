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

%>

<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/projetoslistagem.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td valign="top">
			<table width="570" height="30" border="0" cellpadding="0" cellspacing="0" class="font">
				<tr>
					<td>
						<form name="thisForm" action="projetosmanutencao.asp" method="post">
									
							<input type="hidden" name="hdnExecutar" />
							<input type="hidden" name="pstr_Operacao" />
							<input type="hidden" name="hdnIdRegistro" />
									
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabPesquisa" id="tabPesquisa">
								<COLGROUP>
									<col align="center" width="20"> <!-- Exclusão -->
									<col align="center" width="60"> <!-- Area -->
									<col align="left" width="240"> <!-- Descricao -->
									<col align="center" width="80"> <!-- Equipe -->
									<col align="center" width="80"> <!-- Check Point -->
									<col align="center" width="80"> <!-- Volume -->
								</COLGROUP>
								<tr>
									<td align="left" colspan="6" class="TituloPagina">Projetos</td>
								</tr>
								<tr class="Cabecalho">
									<td>&nbsp;</td>
									<td>Projeto</td>
									<td>Descrição</td>
									<td>Início</td>
									<td>Fim</td>
									<td>Cadastro</td>
									<!--<td>Equipe</td>
									<td>Check Point</td>
									<td>Volume</td> -->
								</tr>
										
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
								vobj_command.CommandText					= "consultaProjetos"
										
										
								' Cria o recordset e posiciona a páginação do recordset.
								Set vobj_rsRegistro = vobj_command.Execute
										
										
								' Verifica se registros foram encontrados.
								If Not vobj_rsRegistro.EOF Then
									
									'Contador pra fazer o efeito zebrado.
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
												
												<td onclick="desativar('<%=vobj_rsRegistro("ID_PROJETO")%>');" title="Desativar Projeto"><img src="<%=getBaseLink("/images/star_on.gif")%>"></td>
												
												<%
											Else
												
												%>
												
												<td onclick="ativar('<%=vobj_rsRegistro("ID_PROJETO")%>');" title="Ativar Projeto"><img src="<%=getBaseLink("/images/star_off.gif")%>"></td>
												
												<%
													
											End If
												
											%>
											
											<td onclick="alterar('<%=vobj_rsRegistro("ID_PROJETO")%>');"><%=vobj_rsRegistro("ID_PROJETO")%></td>
											<td onclick="alterar('<%=vobj_rsRegistro("ID_PROJETO")%>');"><%=vobj_rsRegistro("DS_PROJETO")%></td>
											<td onclick="alterar('<%=vobj_rsRegistro("ID_PROJETO")%>');"><%=converterDataParaHtml(vobj_rsRegistro("DT_INICIO"))%></td>
											<td onclick="alterar('<%=vobj_rsRegistro("ID_PROJETO")%>');"><%=converterDataParaHtml(vobj_rsRegistro("DT_FINAL"))%></td>
											<td onclick="alterar('<%=vobj_rsRegistro("ID_PROJETO")%>');"><%=converterDataParaHtml(vobj_rsRegistro("DT_CADASTRO"))%></td>
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
						</form>
					</td>
				</tr>
				<tr>
					<td align="middle">&nbsp;</td>
				</tr>
				<tr>
					<td align="middle">
						<TABLE ALIGN="center" BORDER="0" CELLSPACING="1" CELLPADDING="1">
							<TR>
								<TD>
									<input type="button" name="cmdIncluir" value="Incluir" onclick="incluir();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Incluir Projeto">
								</TD>
							</TR>
						</TABLE>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
		
<!-- #include file = "../includes/LayoutEnd.asp" -->

<!-- #include file = "../includes/CloseConnection.asp" -->