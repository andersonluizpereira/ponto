<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- #include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->

<html>
	<head>
		<title>Stefanini - Controle de Horas</title>
		<script src="../js/menu.js" type="text/javascript"></script>
        <LINK rel="stylesheet" type="text/css" href="../css/chs.css">
    </head>
	<body>
		<form name="thisForm2" method="post">
		<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">

			<%
			' Declaração de variáveis locais.
			Dim vobj_command
			Dim vobj_rsRegistro


			' ---------------------------------------------------------------------
			' Selecionando todos os registros cadastrados na tabela.
			' ---------------------------------------------------------------------
			Set vobj_command = Server.CreateObject("ADODB.Command")
			Set vobj_command.ActiveConnection = vobj_conexao

			vobj_command.CommandType					= adCmdStoredProc
			vobj_command.CommandText					= "consultaMural"


			' Cria o recordset e posiciona a páginação do recordset.
			Set vobj_rsRegistro = vobj_command.Execute


			' Verifica se registros foram encontrados.
			If Not vobj_rsRegistro.EOF Then
					
				Response.Write Chr(13) & Chr(13)
					
				' Loop de todos os registros cadastrados
				' no banco de dados.
				Do While Not vobj_rsRegistro.EOF
					
					
						
					%>
						
					<tr>
						<td height="30"><%
						
							If	Session("sboo_fladministrador") = True Then
								
								%><input type="checkbox" name="chkExcluirMensagem" value="<%=vobj_rsRegistro("ID_MURAL")%>" /><%
								
							End If
							
							Response.Write Trim(vobj_rsRegistro("DS_MURAL"))%></td>
					</tr>
					<tr>
						<td>Mensagem enviada em: <%=converterDataParaHtml(vobj_rsRegistro("DT_CADASTRO"))%> ás <%=Trim(vobj_rsRegistro("HR_CADASTRO"))%> por: <%=Trim(vobj_rsRegistro("DS_USUARIO"))%>
						</td>
					</tr>
					<tr>
						<td height="30"></td>
					</tr>
					<%
						
					' Move para o próximo registro do loop.
					vobj_rsRegistro.MoveNext
				Loop
			End If

			vobj_rsRegistro.Close
			Set vobj_rsRegistro = Nothing
			Set vobj_command = Nothing
																
			%>
			</table>
			</form>	
		</body>
	</head>
</html>

<!-- #include file = "../includes/CloseConnection.asp" -->