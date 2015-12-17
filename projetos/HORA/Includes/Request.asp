<%
' Fun��o desenvolvida para imprimir na tela todos os 
' parametros do m�todo GET.
Public Sub ExibirParametrosGET()
	
	
	' Declara��o de vari�vel utilizada para percorrer todos
	' os parametros da cole��o.
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
		' Loop por todos os elementos da cole��o de parametros GET.
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


' Fun��o desenvolvida para imprimir na tela todos os 
' parametros do m�todo POST.
Public Sub ExibirParametrosPOST()
	
	
	' Declara��o de vari�vel utilizada para percorrer todos
	' os parametros da cole��o.
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
		' Loop por todos os elementos da cole��o de parametros POST.
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


' Fun��o desenvolvida para imprimir na tela todos os 
' parametros da Sess�o.
Public Sub ExibirParametrosSession()
	
	
	' Declara��o de vari�vel utilizada para percorrer todos
	' os parametros da cole��o.
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
		' Loop por todos os elementos da cole��o de parametros POST.
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