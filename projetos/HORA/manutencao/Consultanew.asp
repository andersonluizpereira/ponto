<!-- #include file = "../includes/Function.asp" -->
<!-- #include file = "../includes/BD.asp" -->
<!-- S#include file = "../includes/GetConnection.asp" -->
<!-- #include file = "../includes/Request.asp" -->
<!-- #include file = "../includes/Validade.asp" -->
<!-- #include file = "../includes/ValidadeSession.asp" -->

<%

If	Not Session("sboo_fladministrador") = True AND Not Session("sboo_flmoderador") = True AND Not Session("sboo_flconsulta") = True Then
	
	Response.Redirect getBaseLink("/horas/horaslancamento.asp")
	
End If

Dim vint_Tipo1
Dim vint_Tipo2
Dim vint_Dia
Dim vint_NmMes
Dim vstr_DsAno
Dim vstr_DtData
Dim vstr_IdUsuario
Dim vobj_command
Dim vobj_command1
Dim ttsaident
vstr_IdUsuario = Request.Form("hdnIdRegistro")
vint_Dia= Request.Form("Data")
vint_Tipo1 =1
vint_Tipo2 =2


%>


<!-- #include file = "../includes/LayoutBegin.asp" -->
<script type="text/javascript" src="js/usuariosmanutencaofiltro.js"></script>
<table class="font" width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20"></td>
	</tr>
	<tr>
		<td style="VERTICAL-ALIGN: top">
			<form name="thisForm" method="post">
				
				
				<i><b class="TituloPagina">Horas no M�s</b></i>
				<table border="0" class="font" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<fieldset style="LEFT: 0px; WIDTH: 1000px;">
							<legend>
							   <b>Hor�rio</b>
							</legend>
							<table class="font" border="0" cellSpacing="1" cellPadding="1" name="tabResultado" id="tabResultado">
								<tr>
									<td colspan="8">
										<strong>Dia: <%= vint_Dia
												
																							
											%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
										</strong>
									</td>
								</tr>
								<COLGROUP />
								<col align="middle" width="50" />
								<col align="middle" width="50" />
								<col align="middle" width="400" />
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
									<th>Data</th>
									<th>User ID</th>
									<th>Nome</th>
									<th>Entrada</th>
									<th>Sa�da Almo�o</th>
									<th>Entrada Almo�o</th>
									<th>Sa�da </th>
									<th>Horas acumuladas</th>
									<th>Tempo Almo�o</th>
									<th>Total </th>
									<th>Periodo de Trab.</th>
									<th>Hor�rio entrada</th>
									<th>Atrasos</th>
								<%     
								
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
	 
	   If Session("sboo_flmoderador") = True Then
  	 
  	 vobj_commandRegistro.CommandText					= "ConsultaHoraDiariatipo1vdi"
  	 
  	 Else
  	 
  	 vobj_commandRegistro.CommandText					= "ConsultaHoraDiariatipo1"
  	 
  	 End if
  	 
  	 vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("Data", adDate, adParamInput, 10, vint_Dia)
  	 'vobj_commandRegistro.Parameters.Append vobj_commandRegistro.CreateParameter("Tipo", adDate, adParamInput, 10, vint_Tipo1)
	 
	Set vobj_commandRegist = Server.CreateObject("ADODB.Command")
	Set vobj_commandRegist.ActiveConnection = vobj_conexao
	
	 vobj_commandRegist.CommandType					= adCmdStoredProc
  	
  	   If Session("sboo_flmoderador") = True Then
  	
  	vobj_commandRegist.CommandText					= "ConsultaHoraDiariatipo2vdi"
  	
  	  Else
  	  
  	  vobj_commandRegist.CommandText					= "ConsultaHoraDiariatipo2"
  	  
  	End if
  	 
  	 vobj_commandRegist.Parameters.Append vobj_commandRegist.CreateParameter("Data", adDate, adParamInput, 10, vint_Dia)
  	 'vobj_commandRegist.Parameters.Append vobj_commandRegist.CreateParameter("Tipo", adDate, adParamInput, 10, vint_Tipo2)
	 
		
	 Set vobj_rsRegistro = vobj_commandRegistro.Execute
	 Set vobj_rsRegist = vobj_commandRegist.Execute
	
	        
	
	If Not vobj_rsRegistro.EOF and Not vobj_rsRegist.EOF Then
								
								Dim contadorClass
								contadorClass = 0
								
								' Loop de todos os registros cadastrados
								' no banco de dados.
								Do While Not vobj_rsRegistro.EOF and Not vobj_rsRegist.EOF
            
            
            hora_ent =      DesencriptaString(vobj_rsRegistro("HR_ENTRADA"))
			hora_sai_almo = DesencriptaString(vobj_rsRegistro("HR_SAIDA"))
			hora_ent_almo = DesencriptaString(vobj_rsRegist("HR_ENTRADALMO"))
			hora_sai =      DesencriptaString(vobj_rsRegist("HR_SAIDAGERA"))
	        
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
	        
	        
	        
	        almo = DateDiff("n", CDate(DesencriptaString(vobj_rsRegist("HR_ENTRADALMO"))), CDate(DesencriptaString(vobj_rsRegistro("HR_SAIDA"))))
	        total_dia = DateDiff("n", CDate(DesencriptaString(vobj_rsRegistro("HR_ENTRADA"))), CDate(DesencriptaString(vobj_rsRegist("HR_SAIDAGERA"))))
	        
	        total = (total_dia + almo)
	        dif = (total_dia + almo) - per
	        
	        
	                            							
								%>
								
			<tr><td><%= vobj_rsRegistro("DT_DATA") %></td>
			<td><%= vobj_rsRegistro("User")%></td>
			<td><%= vobj_rsRegistro("Nomee")%></td>
			
			<td><%= hora_ent      %></td>
			<td><%= hora_sai_almo %></td>
		    
		 
		 
		 
		    
		    <td><%=    hora_ent_almo %>
		    
		     </td> 
			
			<td><%
			
			If Not IsNull(hora_sai) Then 
			
			Response.write hora_sai
			
			Else 
			
			Response.Write ""
			
			End If
			
			%> </td>		
			
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
								
					' Move para o pr�ximo registro do loop.
                   
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
									<th>&nbsp;Qtd. Pessoas que trab.</th>
									<th><%= contadorClass %>&nbsp;</th>
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
									
									
									<td><input type="button" value="Retornar" onclick="voltar();" class="BotaoOff" onmouseover="this.className='BotaoOn';" onmouseout="this.className='BotaoOff';" title="Retornar a tela anterior"></td>
									
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

