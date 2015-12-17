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

<script type="text/javascript" src="js/usuarioslistagem.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td valign="top">
		<table width="570" height="30" border="0" cellpadding="0" cellspacing="0" class="font">
			<tr>
				<td>
					<form name="thisForm" action="usuariosmanutencao.asp" method="post">
						
						<input type="hidden" name="hdnExecutar" />
						<input type="hidden" name="pstr_Operacao" />
						<input type="hidden" name="hdnIdRegistro" />
						
						<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabPesquisa" id="tabPesquisa"> 
						<COLGROUP>
							<col align="middle" width="20"> <!-- Exclusão -->
							<col align="middle" width="60"> <!-- User_ID -->
							<col align="left"	width="200"> <!-- Nome -->
							<col align="middle" width="80"> <!-- CPF -->
							<col align="left"   width="80"> <!-- RG -->
							<col align="middle" width="80"> <!-- Perfil -->
						</COLGROUP>
							<tr>
								<td align="left" colspan="6" class="TituloPagina">Usuários&nbsp;<input type="button" name="cmdIncluir" value="Incluir" onclick="incluir();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Incluir Usuário"></td>
							</tr>
							<tr class="Cabecalho">
								<td>&nbsp;</td>
								<td>User</td>
								<td>Nome</td>
								<td>CPF</td>
								<td>RG</td>
								<td>Perfil</td>
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
							vobj_command.CommandText					= "consultaUsuarios"
							
							
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
					</form>
				</td>
			</tr>
			<tr>
				<td align="middle" >
					&nbsp;
				</td>
			</tr>
			<tr>
				<td align="middle" >
					<TABLE ALIGN=center BORDER=0 CELLSPACING=1 CELLPADDING=1>
						<TR>
							<TD>
								<input type="button" name="cmdIncluir" value="Incluir" onclick="incluir();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Incluir Usuário">
							</TD>
						</TR>
					</TABLE>
				</td>
			</tr>
		</table>
		</td>
	</tr>
</TABLE>

<!-- #include file = "../includes/LayoutEnd.asp" -->


<!-- #include file = "../includes/CloseConnection.asp" -->