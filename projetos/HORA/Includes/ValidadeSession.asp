<!-- #include file = "../includes/Request.asp" -->
<%
' Declara��o de todas as vari�veis utilizadas para
' validar a sess�o do usu�rio logado no sistema.
Dim vint_IdUsuarioLogado

' Conseguindo os parametros obrigat�rios para
' a utiliza��o do sistema.
vint_IdUsuarioLogado	= Session("sstr_IdUsuario")


' Verificando se os parametros est�o todos corretamentes
' especificados pela sess�o do usu�rio.
If Trim(vint_IdUsuarioLogado) = "" Then
	
	
	' Finaliza a sess�o do usu�rio.
	Session.Abandon
	Response.Redirect getBaseLink("/login/login.asp")
End If
%>