<script language="javascript">

	function funAjuda() {
		document.location =  "<%=getBaseLink("/HELP.hlp")%>";
	}
	
	function funHome() {
		document.location = "<%=getBaseLink("/horas/horaslancamento.asp")%>";
	}

	function funLogoff() {
		document.location = "<%=getBaseLink("/login/logoff.asp")%>";
	}
</script>
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td><img SRC="<%=getBaseLink("/images/menu_inicio.gif")%>"></td>
		<td><img style="cursor:hand" onclick="funAjuda();" SRC="<%=getBaseLink("/images/menu_ajuda.jpg")%>"></td>
		<td><img SRC="<%=getBaseLink("/images/menu_divisor.jpg")%>"></td>
		<td><img style="cursor:hand" onclick="funHome();" SRC="<%=getBaseLink("/images/menu_home.jpg")%>"></td>
		<td><img SRC="<%=getBaseLink("/images/menu_divisor.jpg")%>"></td>
		<td><img style="cursor:hand" onclick="funLogoff();" SRC="<%=getBaseLink("/images/menu_logoff.jpg")%>"></td>
	</tr>
</table>