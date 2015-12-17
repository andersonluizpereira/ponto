<!-- #include file = "../includes/Request.asp" -->
<%
' Declaração de todas as variáveis utilizadas para
' validar a sessão do usuário logado no sistema.
Dim vint_IdUsuarioLogado

' Conseguindo os parametros obrigatórios para
' a utilização do sistema.
vint_IdUsuarioLogado	= Session("sstr_IdUsuario")


' Verificando se os parametros estão todos corretamentes
' especificados pela sessão do usuário.
If Trim(vint_IdUsuarioLogado) = "" Then
	
	
	' Finaliza a sessão do usuário.
	Session.Abandon
	Response.Redirect getBaseLink("/login/login.asp")
End If
%>