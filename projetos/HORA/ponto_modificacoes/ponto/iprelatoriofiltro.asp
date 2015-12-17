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

' Declaração de variáveis utilizadas para armazenar os
' valores dos campos da tela.
Dim vint_IdMes
Dim vstr_DsAno

' para está página.
vstr_Operacao		= Request.Form("pstr_Operacao")
vstr_Processar		= Request.Form("hdnProcessar")
vstr_Executar		= Request.Form("hdnExecutar")

' Verifica se o parametro que defini o tipo
' de operação a ser executado na página é
' igual a branco(vazio).
If Trim(vstr_Operacao) = "" Then
	
	' ... neste caso a operação
	' padrão de a de visualização apenas
	' do registro.
	vstr_Operacao = "V"
End If

' Veririfa se a operação a ser executada nesta página é a 
' operação de Alteração ou Visualização e se a página não
' foi processada ainda.
If (vstr_Operacao = "A" or vstr_Operacao = "V") And vstr_Processar <> "S" Then
	
	
	' ... neste caso deve ser solicitado o código do registro
	' e encontrar suas informações no banco de dados para exibir para
	' as informações do registro na tela.
	' Conseguindo o código do registro.
	
	
Else
	
	
	' Verifica se a operação a ser executada nesta página é
	' a operação de inclusão e verifica se a página não foi
	' processada ainda.
	If vstr_Operacao = "I" And vstr_Processar <> "S" Then
		
		' Neste caso todas as variáveis devem ser vazias
		' para o usuário poder preencher seu novo cadastro
		' do registro.
		
		
		vint_IdMes				= Empty
		vstr_DsAno				= Empty
		
	Else
		
		' ... está opção acontecerá quando o usuário processar
		' a página, neste caso todas os dados da tela serão
		' submetidos e devem ser pegos neste lugar.
		
		vint_IdMes				= Request.Form("cmbComboMes")
		vstr_DsAno				= Request.Form("txtDsAno")
		
	End If
End If

%>

<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/iprelatoriofiltro.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<form name="thisForm" action="iprelatorio.asp" method="post">
				
				<input type="hidden" name="hdnProcessar" value="S" />
				<input type="hidden" name="pstr_Operacao" value="<%=vstr_Operacao%>" />
				<input type="hidden" name="hdnExecutar" />
				
				<i><b class="TituloPagina">Relatório por IP</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 595px; HEIGHT: 50px">
							<legend>
							   <b>Filtro Relatório</b>
							</legend>
							<TABLE valign="center" class="font" BORDER=0 CELLSPACING=1 CELLPADDING=1>
								<tr>
									<td>&nbsp;Mês&nbsp;</td>
									<td>
										<select name="cmbComboMes" id="Data" class="TextBox">
											<option value="" selected>Selecione</option>
											<%
											
											Dim vvar_ArrayMes(11)
											
											vvar_ArrayMes(0) = "Janeiro"
											vvar_ArrayMes(1) = "Fevereiro"
											vvar_ArrayMes(2) = "Março"
											vvar_ArrayMes(3) = "Abril"
											vvar_ArrayMes(4) = "Maio"
											vvar_ArrayMes(5) = "Junho"
											vvar_ArrayMes(6) = "Julho"
											vvar_ArrayMes(7) = "Agosto"
											vvar_ArrayMes(8) = "Setembro"
											vvar_ArrayMes(9) = "Outubro"
											vvar_ArrayMes(10) = "Novembro"
											vvar_ArrayMes(11) = "Dezembro"
											
											Dim vint_Contator
											
											For vint_Contator = 1 To 12
												
												If Month(Date) = vint_Contator Then
													
													%><option value="<%=vint_Contator%>" selected><%=vvar_ArrayMes(vint_Contator - 1)%></option><%
													
												Else
													
													%><option value="<%=vint_Contator%>"><%=vvar_ArrayMes(vint_Contator - 1)%></option><%
													
												End If
											Next
											
											%>
											
										</select>
									</td>
									<td>&nbsp;&nbsp;Ano&nbsp;</td>
									<td>
										<input class="TextBox" type="text" maxlength="4" size="5" name="txtDsAno" Value="<%=Year(Date())%>">
									</td>
								</tr>
							</TABLE>
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
									<td><input type="button" value="Listar" onclick="listar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Lista IPs do mês selecionado"></td>
									<td><input type="button" value="Retornar" onclick="voltar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Retorna a página anterior"></td>
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