<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- #include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->




<%
 
Dim vstr_Operacao

' Variável flag que indica se a página deve ser 
' processada, apenas disponivel para as operações de
' A e I.
Dim vstr_Processar

'Variável de controle e fluxo de acoes
Dim vstr_Executar

' Armazena o código de referência do registro que será alterado, Inclusso
' ou visualizado.
Dim vstr_IdUsuario

' Declaração de variáveis utilizadas para armazenar os
' valores dos campos da tela.
Dim vstr_CdUsuario
Dim vstr_DsUsuario
Dim vstr_DsPer
Dim vstr_DsHorasen
Dim vstr_DsCPF
Dim vstr_DsRG
Dim vint_IdFuncao
Dim vint_IdEquipe
Dim vint_FlPerfil
Dim vstr_DtNascimento
Dim vstr_DsTelefone
Dim vstr_DsRamal
Dim vstr_DsLocalAlocado


Dim vstr_DtAniversario

'Dim vint_FlAtivo
Dim vstr_CdSenha
Dim vstr_CdConfirmaSenha
Dim vobj_commandProc
Dim vobj_rs
Dim vobj_rsRegistroConsulta
Dim ValidarForm



        
        
        vstr_CdUsuario			= Request.Form("txtCdUsuario")
		vstr_DsUsuario			= Request.Form("txtDsUsuario")
		vstr_DsPer			    = Request.Form("txtDsPer")
		vstr_DsHorasen		    = Request.Form("txtDshorasen")
		vstr_DsCPF				= Request.Form("txtDsCPF")
		vstr_DsRG				= Request.Form("txtDsRG")
		vint_IdFuncao			= Request.Form("cmbComboFuncao")
		vint_IdEquipe			= Request.Form("txtDsEquipe")
		vint_FlPerfil			= Request.Form("cmbComboPerfil")
		'vint_FlAtivo			= Request.Form("txtFlAtivo")
		vstr_CdSenha			= Request.Form("txtCdSenha")
		vstr_CdConfirmaSenha	= Request.Form("txtCdConfirmaSenha")
		
		vstr_DtNascimento		= Request.Form("txtDtNascimento")
		vstr_DsTelefone			= Request.Form("txtDsTelefone")
		vstr_DsRamal			= Request.Form("txtDsRamal")
		vstr_DsLocalAlocado		= Request.Form("txtDsLocalAlocado")
		
			
			If ValidarForm = True Then
				
				' ---------------------------------------------------------------------
				' Procedimento desenvolvimento para tratar a entrada de umas mesma
				' area
				' ---------------------------------------------------------------------
				
				' Declaração de variáveis auxiliares
				' para obter as informações do registro.


				
				
				' ---------------------------------------------------------------------
				' Selecionando os dados do registro.
				' ---------------------------------------------------------------------
				Set vobj_commandRegistroConsulta = Server.CreateObject("ADODB.Command")
				Set vobj_commandRegistroConsulta.ActiveConnection = vobj_conexao
				
				
				vobj_commandRegistroConsulta.CommandType					= adCmdStoredProc
				vobj_commandRegistroConsulta.CommandText					= "consultaUsuario"
				
				
				vobj_commandRegistroConsulta.Parameters.Append vobj_commandRegistroConsulta.CreateParameter("param1",adChar, adParamInput, 10, vstr_CdUsuario)
			
			
			
				If Not vobj_rsRegistroConsulta.EOF Then
					
					Call AddErro("Erro", "Há um registro com o mesmo nome de Usuário.")
					
					
					' ---------------------------------------------------------------------
					' Incluindo os dados do registro no banco de dados.
					' ---------------------------------------------------------------------
					Set vobj_commandProc = Server.CreateObject("ADODB.Command")
					Set vobj_commandProc.ActiveConnection = vobj_conexao
					
					vobj_commandProc.CommandType					= adCmdStoredProc
					vobj_commandProc.CommandText					= "incluiUsuario"
					
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param1",adChar, adParamInput, 10, Trim(vstr_CdUsuario))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param2",adChar, adParamInput, 100, Trim(vstr_DsUsuario))
                    vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param3",adChar, adParamInput, 11, Trim(vstr_DsPer))
                    vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param4",adChar, adParamInput, 11, Trim(vstr_DsHorasen))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param5",adChar, adParamInput, 11, Trim(vstr_DsCPF))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param6",adChar, adParamInput, 15, Trim(vstr_DsRG))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param7",adInteger, adParamInput,, vint_IdFuncao)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param8",adChar, adParamInput, 25, vint_IdEquipe)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param9",adInteger, adParamInput,, vint_FlPerfil)
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param10",adChar, adParamInput, 15, EncriptaString(Trim(vstr_CdSenha)))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param11",adDate, adParamInput,, converterDataParaSQL(Date()))
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param12",adChar, adParamInput, 10, Trim(vstr_DtNascimento))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param13",adChar, adParamInput, 20, Trim(vstr_DsTelefone))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param14",adChar, adParamInput, 15, Trim(vstr_DsRamal))
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param15",adChar, adParamInput, 30, Trim(vstr_DsLocalAlocado))
					
					
					
					
					If Not Trim(vstr_DtNascimento) = "" Then
						
						vstr_DtAniversario = DateSerial(2000, Month(vstr_DtNascimento), Day(vstr_DtNascimento))
						
					End If
					
					vobj_commandProc.Parameters.Append vobj_commandProc.CreateParameter("param13",adChar, adParamInput, 10, converterDataParaSQL(vstr_DtAniversario))
					
					vobj_commandProc.Execute
					
					Response.write "resre"
					
					Response.Redirect("usuarioslistagem.asp")
					
					
					
					
					
					' Altera a variável que indica o tipo de
					' operação que é executada na página.
					
					
					
					' Redireciona para a página de listagem
					' dos registros.

		


End If
End If
 %>
