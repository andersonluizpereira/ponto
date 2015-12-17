<%
Public vobj_conexao
Public vboo_ConexaoAberta


'======================= Abre Conexao Com o Banco =================
Public Sub AbrirConexao
	Set vobj_conexao = Server.CreateObject("ADODB.Connection")
	vobj_conexao.ConnectionTimeout = 0
	vobj_conexao.CommandTimeout = 0
	vobj_conexao.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Projetos\Hora\DB\CHS.mdb;Persist Security Info=False;"
	vobj_conexao.CursorLocation = 3	
	
	' Seta a variável que indica que a conexão
	' está aberta para true.
	vboo_ConexaoAberta = True
End Sub

'====================== Fecha a Conexao Com o Banco ==============
Public Sub FecharConexao
	vobj_conexao.Close
	Set vobj_conexao = Nothing
	
	' Seta a variável para indicar que a conexão 
	' está fechada.
	vboo_ConexaoAberta = False
End Sub

%>


<!--METADATA
	TYPE="TypeLib"
	NAME="Microsoft ActiveX Data Objects 2.5 Library"
	UUID="{00000205-0000-0010-8000-00AA006D2EA4}"
	VERSION="2.5"
-->