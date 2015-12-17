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

' Armazena o código de referência do registro que será alterado, Inclusso
' ou visualizado.
Dim vstr_IdUsuario

' Declaração de variáveis utilizadas para armazenar os
' valores dos campos da tela.
Dim vint_NmMes
Dim vstr_DsAno
Dim vstr_DtData


' para está página.
vstr_Operacao		= Request.Form("pstr_Operacao")
vstr_Executar		= Request.Form("hdnExecutar")



vstr_IdUsuario			= Request.Form("cmbComboUsuario")
vint_NmMes				= Cint(Request.Form("cmbComboMes"))
vstr_DsAno				= Cint(Request.Form("txtDsAno"))
vstr_DtData             = Cint(Request.Form("txtDsAno"))

If vstr_IdUsuario = "" Then
	
	Response.Redirect("horasmanutencaofiltro.asp")
	
ElseIf vint_NmMes = "" Then
	
	Response.Redirect("horasmanutencaofiltro.asp")

ElseIf vstr_DsAno = "" Or Not Len(vstr_DsAno) = 4 Or Not IsNumeric(vstr_DsAno) Then

	Response.Redirect("horasmanutencaofiltro.asp")
	
End If

' Ajusta hora do almoço para casos que tem mais de 7 horas trabalhas e não marcou pausa obrigatoria no sistema.
' Obs. Sabado não inclu essa regra.
If vstr_Executar = "AJUSTAR_HORA" Then
	
	vstr_DtData = Request.Form("hdnDtData")
	
	'Verufucando se a data não final de semana, se for final de semana a regra não é necessária.
	If Not DiaDaSemana(vstr_DtData) = "Domingo" And Not DiaDaSemana(vstr_DtData) = "Sábado" Then
		
		Dim vint_MinutoDiferencaHora
		Dim vint_MinutoDiferencaAlmoco
		
		Dim vint_ContadorRegistroRSAjustarHora
		
		vint_MinutoDiferencaHora = 0
		vint_MinutoDiferencaAlmoco = 0
		
		


End If		
End If
			
%>

<!-- #include file = "../includes/LayoutBegin.asp" -->

<script type="text/javascript" src="js/horasmanutencao.js"></script>

<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<form name="thisForm" action="horasmanutencaolancamento.asp" method="post">
				
				<input type="hidden" name="hdnProcessar" />
				<input type="hidden" name="pstr_Operacao" />
				<input type="hidden" name="hdnExecutar" />
				<input type="hidden" name="hdnDtData" />
				<input type="hidden" name="cmbComboUsuario" value="<%=vstr_IdUsuario%>" />
				<input type="hidden" name="cmbComboMes" value="<%=vint_NmMes%>" />
				<input type="hidden" name="txtDsAno" value="<%=vstr_DsAno%>" />
				<input type="hidden" name="hdnBloqueaData" />
				
				<i><b class="TituloPagina">Horas no Mês</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 1055px;">
							<legend>
							   <b>Horário</b>
							</legend>
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabResultado" id="tabResultado">
								<tr>
									<td colspan="8">
										<strong>Colaborador: <%
												
												Response.Write NomeUsuario(vstr_IdUsuario)
											
											%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Mês:<%
												
												Response.Write Ucase(DescricaoMes(vint_NmMes)) & "/" & vstr_DsAno 
										%></strong>
									</td>
								</tr>
								<COLGROUP />
								<col align="middle" width="30" />
								<col align="middle" width="25" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<tr class="Cabecalho">
									<th></th>
									<th></th>
									<th>Data</th>
									<th>Entrada</th>
									<th>Saída <br />Almoço</th>
									<th>Entrada Almoço</th>
									<th>Saída</th>
									<th>Total</th>
									<th>Horas <br /> Acum.</th>
									<th>Horário entrada</th>
									<th>Atrasos</th>
									<th>Observações</th>
								</tr>
							</table>
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabResultado" id="tabResultado">
								<COLGROUP />
								<col align="middle" width="30" />
								<col align="middle" width="25" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
								<col align="middle" width="100" />
												
								<%
			Dim vstr_DataConsulta
			Dim vint_NmUltimoDia
			
			vint_NmUltimoDia = Cint(GetUltimoDiaMes(vint_NmMes, vstr_DsAno))
													
			Dim vobj_command
			Dim vobj_commandRegistro
			Dim vobj_rsRegistro
			Dim vobj_rs
			Dim vobj_commandRegist
			Dim vobj_rsRegist
			
			Dim hora_ent
			Dim hora_sai_almo
			Dim hora_ent_almo
			Dim hora_sai
			Dim total_dia
			Dim almo
			Dim dif
			Dim total
			Dim x
			Dim y
			Dim z
			Dim w
			Dim per
			Dim horas
			Dim vint_tot
			Dim difer
			Dim conv
			Dim conv1
			Dim tt
			Dim tt1
			
			
			
			//Dim vobj_rs
			

					
	' ---------------------------------------------------------------------
	' Selecionando os dados do registro.
	' ---------------------------------------------------------------------
	Set vobj_commandRegistro = Server.CreateObject("ADODB.Command")
	Set vobj_commandRegistro.ActiveConnection = vobj_conexao
	
	 vobj_commandRegistro.CommandType					= adCmdStoredProc
  	 vobj_commandRegistro.CommandText					= "consultaRelatorioHora1"
  	 
  	 vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("ID_USUARIO", adChar, adParamInput, 10, vstr_IdUsuario)
  	 vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("DT_DATA", adDate, adParamInput, ,vstr_DtData)
	 
	Set vobj_commandRegist = Server.CreateObject("ADODB.Command")
	Set vobj_commandRegist.ActiveConnection = vobj_conexao
	
	 vobj_commandRegist.CommandType					= adCmdStoredProc
  	vobj_commandRegist.CommandText					= "consultaRelatorioHora2"
  	 
  	 vobj_commandRegist.Parameters.Append vobj_commandRegist.CreateParameter("ID_USUARIO", adChar, adParamInput, 10, vstr_IdUsuario)
  	 vobj_commandRegist.Parameters.Append vobj_commandRegist.CreateParameter("DT_DATA", adDate, adParamInput, , vstr_DtData)
	 
		
	 Set vobj_rsRegistro = vobj_commandRegistro.Execute
	 Set vobj_rsRegist = vobj_commandRegist.Execute
	
	        
	
	If Not vobj_rsRegistro.EOF and Not vobj_rsRegist.EOF Then
								
								Dim contadorClass
								contadorClass = 0
								
								' Loop de todos os registros cadastrados
								' no banco de dados.
								Do While Not vobj_rsRegistro.EOF AND Not vobj_rsRegist.EOF
            
            
            hora_ent = DesencriptaString(vobj_rsRegistro("HR_ENTRADA"))
			hora_sai_almo = DesencriptaString(vobj_rsRegistro("HR_SAIDA"))
			hora_ent_almo = DesencriptaString(vobj_rsRegist("HR_ENTRADA"))
			hora_sai = DesencriptaString(vobj_rsRegist("HR_SAIDA"))
	        
	        x = DesencriptaString(vobj_rsRegistro("HR_ENTRADA"))	
	        
	        //y = Minute(converterHoraParaSQL(hora_sai_almo))
	        z = Minute(converterHoraParaSQL(hora_ent_almo))
	        w = Minute(converterHoraParaSQL(hora_sai))
	        horas = Minute(vobj_rsRegistro("Horase"))
	        per = Hour(converterHoraParaSQL(vobj_rsRegistro("per")))
	        
	           tt = Hour(converterHoraParaSQL(vobj_rsRegistro("Horase")))
               y = Hour(converterHoraParaSQL(x))
                                                
                                                 tt1 = Hour(converterHoraParaSQL(x))
                                                   
                                                 conv1 = Minute(converterHoraParaSQL(vobj_rsRegistro("Horase")))
                                                 conv =  Minute(converterHoraParaSQL(x))                                             
                                           
                                                If((y-tt)=-5) Then
                                                    
                                                 difer = conv-(60*5)
                                                
                                                 Else If((y-tt)=-4) Then
                                                    
                                                 difer = conv-(60*4)
                                                 
                                                 Else If((y-tt)=-3) Then
                                                    
                                                 difer = conv-(60*3)
                                                 
                                                 Else If((y-tt)=-2) Then
                                                    
                                                 difer = conv-(60*2)
                                                  
                                                  Else If((y-tt)=-1) Then
                                                    
                                                 difer = conv-60
                                                                                            
                                                  Else If ((y-tt)=5) Then                                            
                                                
                                                   difer = conv+(60*5)
                                                  
                                                  Else If ((y-tt)=4) Then                                            
                                                
                                                   difer = conv+(60*4)
                                                   
                                                  Else If ((y-tt)=3) Then                                            
                                                
                                                   difer = conv+(60*3)
                                                  
                                                  Else If ((y-tt)=2) Then                                            
                                                
                                                   difer = conv+(60*2)
                                                   
                                                  Else If ((y-tt)=1) Then                                            
                                                
                                                   difer = conv+(60*1)
                                                   
                                                 
                                                 Else 
                                                 difer = (conv - conv1)
                                                 
                                                 
                                                 
                                                 
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                 End IF  
                                                  
	       
	       
	       
	       
	        If(per=6) Then
                                                 
                                                 vint_tot = 360
                                                 
   else if (per=8) Then
  vint_tot = 480
                                              
   End IF
   End IF
	        
	        
	        
	        almo = DateDiff("n", CDate(DesencriptaString(vobj_rsRegist("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rsRegistro("HR_SAIDA"))))
	        total_dia = DateDiff("n", CDate(DesencriptaString(vobj_rsRegistro("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rsRegist("HR_SAIDA"))))
	        
	        total = (total_dia + almo)
	        dif = (total_dia + almo) - 480
	        
	        
	                            							
								%>
								
			<tr><td><%= vobj_rsRegistro("DT_DATA") %></td>
			<td><%= vobj_rsRegistro("User")%></td>
			<td><%= vobj_rsRegistro("Nomee")%></td>
			<td><%=DesencriptaString(vobj_rsRegistro("HR_ENTRADA"))%></td>
			<td><%=DesencriptaString(vobj_rsRegistro("HR_SAIDA"))%></td>
		    <td><%=DesencriptaString(vobj_rsRegist("HR_ENTRADA"))%></td>
			<td><%=DesencriptaString(vobj_rsRegist("HR_SAIDA"))%></td>		
			<td><% Response.Write converterMinutoParaHora(dif)  %></td>	
			<td><% Response.Write converterMinutoParaHora(almo) %></td>
			<td><% 
			
			If total < vint_tot Then  
			Response.Write "<font color='red'>" 
			Response.Write converterMinutoParaHora(total)
			
			Else
			
			Response.Write converterMinutoParaHora(total)
			
			End If
			
			
			 %>
			
			
			
			
			
			
			</td>
			<td><%= vobj_rsRegistro("per")                      %></td>	
			<td><%= vobj_rsRegistro("Horase")                      %></td>	
			<td><%= converterMinutoParaHora(difer)  %>      </td> 		
				
				
				<%
				
				      contadorClass = contadorClass + 1
								
					' Move para o próximo registro do loop.
					vobj_rsRegistro.MoveNext
					vobj_rsRegist.MoveNext
					Loop
					Else
					 %>				
					 	<%
	
							End If
							
							
							
							vobj_rsRegistro.Close
							vobj_rsRegist.Close
							Set vobj_rsRegistro = Nothing
							Set vobj_rsRegist = Nothing
							Set vobj_command = Nothing
			
			
								
		
		
								%>
								
								<tr class="Cabecalho">
                                    <th>&nbsp;</th>
									<th>&nbsp;</th>
									<th>TOTAL&nbsp;</th>
									<th>&nbsp;</th>
									<th>&nbsp;Dias trabalh.</th>
									<th>&nbsp;</th>
									<th>&nbsp;</th>
									<th>&nbsp;</th>
									<th>&nbsp;</th>
									<th>&nbsp;Total Atrasos</th>
									<th>&nbsp;</th>
									<th>&nbsp;</th>
								</tr>
							</table>
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
									<td><input type="button" value="Tela Impressão" onclick="imprimir();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Envia para a tela de impressão"></td>
									<td><input type="button" value="Incluir" onclick="incluir();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Inclui uma nova data"></td>
									<td><input type="button" value="Retornar" onclick="voltar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Retornar a tela anterior"></td>
									<td><input type="button" value="Alterar Lote" onclick='alterarLote(thisForm);' class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Alterar em Lote"></td>
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
<%
Private Function NomeUsuario(pstr_IdUsuario)
		
	' Declaração de variáveis auxiliares
	' para obter as informações do registro.
	Dim vobj_rsRegistro
	Dim vobj_commandRegistro
	
	
	' ---------------------------------------------------------------------
	' Selecionando os dados do registro.
	' ---------------------------------------------------------------------
	Set vobj_commandRegistro = Server.CreateObject("ADODB.Command")
	Set vobj_commandRegistro.ActiveConnection = vobj_conexao
							
							
	vobj_commandRegistro.CommandType					= adCmdStoredProc
	vobj_commandRegistro.CommandText					= "consultaUsuario"
	
	vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("param1", adChar, adParamInput, 10, vstr_IdUsuario)
	' ---------------------------------------------------------------------
	
	
	' Cria o objeto recordset com as informações do registro.	
	Set vobj_rsRegistro = vobj_commandRegistro.Execute
	
	If Not vobj_rsRegistro.EOF Then
		
		NomeUsuario = Trim(vobj_rsRegistro("ID_USUARIO")) & "  -  " & Trim(vobj_rsRegistro("DS_USUARIO"))
		
	End If
	
	vobj_rsRegistro.Close
	Set vobj_rsRegistro = Nothing
	Set vobj_commandRegistro = Nothing
		
		
End Function
 %>
<!-- #include file = "../includes/CloseConnection.asp" -->