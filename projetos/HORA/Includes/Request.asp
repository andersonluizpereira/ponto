<%
' Função desenvolvida para imprimir na tela todos os 
' parametros do método GET.
Public Sub ExibirParametrosGET()
	
	
	' Declaração de variável utilizada para percorrer todos
	' os parametros da coleção.
	Dim vstr_ContaParametroGET
	%>
	
	
	<TABLE border="1">
		<TR>
			<TD colspan="2" align="center"><i><b>Metodo GET</b></i></TD>
		</TR>
		<TR>
			<TD><b>Parametro</b></TD>
			<TD><b>Valor</b></TD>
		</TR>
		
		
		<%
		' Loop por todos os elementos da coleção de parametros GET.
		For Each vstr_ContaParametroGET In Request.QueryString
			%>
			
			
			<TR>
				<TD>&nbsp;<%=vstr_ContaParametroGET%></TD>
				<TD>&nbsp;<%=Request.QueryString(vstr_ContaParametroGET)%></TD>
			</TR>
			
			
			<%
		Next
		%>
		
		
	</TABLE>
	
	
	<%
End Sub


' Função desenvolvida para imprimir na tela todos os 
' parametros do método POST.
Public Sub ExibirParametrosPOST()
	
	
	' Declaração de variável utilizada para percorrer todos
	' os parametros da coleção.
	Dim vstr_ContaParametroPOST
	%>
	
	
	<TABLE border="1">
		<TR>
			<TD colspan="2" align="center"><i><b>Metodo POST</b></i></TD>
		</TR>
		<TR>
			<TD><b>Parametro</b></TD>
			<TD><b>Valor</b></TD>
		</TR>
		
		
		<%
		' Loop por todos os elementos da coleção de parametros POST.
		For Each vstr_ContaParametroPOST In Request.Form
			%>
			
			
			<TR>
				<TD>&nbsp;<%=vstr_ContaParametroPOST%></TD>
				<TD>&nbsp;<%=Request.Form(vstr_ContaParametroPOST)%></TD>
			</TR>
			
			
			<%
		Next
		%>
		
		
	</TABLE>
	
	
	<%
End Sub


' Função desenvolvida para imprimir na tela todos os 
' parametros da Sessão.
Public Sub ExibirParametrosSession()
	
	
	' Declaração de variável utilizada para percorrer todos
	' os parametros da coleção.
	Dim vstr_ContaParametroSession
	%>
	
	
	<TABLE border="1">
		<TR>
			<TD colspan="2" align="center"><i><b>Parametros Session</b></i></TD>
		</TR>
		<TR>
			<TD><b>Parametro</b></TD>
			<TD><b>Valor</b></TD>
		</TR>
		
		
		<%
		' Loop por todos os elementos da coleção de parametros POST.
		For Each vstr_ContaParametroSession In Session.Contents
			%>
			
			
			<TR>
				<TD>&nbsp;<%=vstr_ContaParametroSession%></TD>
				<TD>&nbsp;<%=Session(vstr_ContaParametroSession)%></TD>
			</TR>
			
			
			<%
		Next
		%>
		
		
	</TABLE>
	
	
	<%
End Sub


Public Function getBaseLink(pstr_URL)
	
	Dim vstr_Server
	Dim vstr_SistemaWeb
	
	
	vstr_Server = Request.ServerVariables("HTTP_HOST")
	vstr_SistemaWeb = Request.ServerVariables("URL")
	
	vstr_SistemaWeb = Mid(vstr_SistemaWeb, 2, Instr(2, vstr_SistemaWeb, "/") - 2)
	
	getBaseLink = "http://" & vstr_Server & "/" & vstr_SistemaWeb & pstr_URL
End Function
%>