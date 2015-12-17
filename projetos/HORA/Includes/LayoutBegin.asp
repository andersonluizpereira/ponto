<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title>Stefanini - Controle de Horas</title>
		<link href="../css/arquivo.ie.css" rel="stylesheet" type="text/css" media="screen" />
        <script src="../js/menu.js" type="text/javascript"></script>
        <LINK rel="stylesheet" type="text/css" href="../css/chs.css">
    </head>

	<body>

		<div id="conteudo">
			<div id="topo">
		    	<div id="topo-info"><%=Session("sstr_DsUsuario")%> | <a href="<%=getBaseLink("/login/logoff.asp")%>">Sair</a></div>
		    	<div id="topo-logo">
		    	</div>
				<div id="topo-info2">
		    	</div>
		        <div id="topo-menu">
                    <div id="Master_Header_Menu">
                        <ul>
                            <%
                            
                            If Session("sboo_fladministrador") = True  Then
								
								%>
                            
								<li class="menuItem menuItemFirst">
								    <a href="#">Manutenção</a> 
								   
								    <ul class="menu1">
								    	<li class="menuItem">
								            <a href="#">Cadastro Básico &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ></a>
								            <ul class="menu9">
												<li class="menuItem">
												 	<a href="<%=getBaseLink("/manutencao/arealistagem.asp")%>">Áreas</a>
												</li>
												<li class="menuItem">
												 	<a href="<%=getBaseLink("/manutencao/atividadelistagem.asp")%>">Atividades</a>
												</li>
												<li class="menuItem">
												 	<a href="<%=getBaseLink("/manutencao/funcaolistagem.asp")%>">Funções/Cargos</a>
												</li>
												
												<!--<li class="menuItem">
												 	<a href="<%=getBaseLink("/manutencao/usuarioscad.asp")%>">Cadastrar Usuários</a>
												</li>-->
												
											</ul>
								        </li>
								        
								        <li class="menuItem">
								            <a href="#">Relação &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ></a> 
								            <ul class="menu9">
												<li class="menuItem">
												 	<a href="<%=getBaseLink("/manutencao/projetoslistagem.asp")%>">Projetos</a>
												</li>
											
											     <li class="menuItem">
												 	<a href="<%=getBaseLink("/manutencao/usuarioslistagem.asp")%>">Usuários</a>
											
												<li class="menuItem">
												 	<a href="<%=getBaseLink("/manutencao/usuariosprojetos.asp")%>">Usuários Projeto</a>
												</li>
																		
											</ul>
								        </li>
								        
								        
								        
								        
								   <%    ''''''''''''''''''''''''  %>
								         <li class="menuItem">    
								    
										 	<a href="<%=getBaseLink("/horas/horasmanutencaofiltro.asp")%>">Horas</a>
										</li>
										
										<li class="menuItem">
										 	<a href="<%=getBaseLink("/manutencao/usuariosmanutencaofiltro.asp")%>">Consultar Horas</a>
										</li>
										
								      
                                      
                                        <li class="menuItem">
										 	<a href="<%=getBaseLink("/horas/iprelatoriofiltro.asp")%>">Consulta IP</a>
										</li>
									
								   
								   <%    ''''''''''''''''''''''''  %>
								        
								        
								        
								        <%
								    
								    Else If Session("sboo_flmoderador")	= True  Then	
								       	
										 %>    
							
							     	<li class="menuItem menuItemFirst">
								    <a href="#">Manutenção</a> 
								   
								    <ul class="menu1">
								    
								       
								        <li class="menuItem">    
								    
										 	<a href="<%=getBaseLink("/horas/horasmanutencaofiltro.asp")%>">Horas</a>
										</li>
										
										<li class="menuItem">
										 	<a href="<%=getBaseLink("/manutencao/usuariosmanutencaofiltro.asp")%>">Consultar Horas</a>
										</li>
										
								      
                                      
                                        
										
               							  </ul>
								</li>
                                                                               
                                                                                 <%
								    
								    Else If Session("sboo_flconsulta")	= True  Then	
								       	
										 %>    
							
							     	<li class="menuItem menuItemFirst">
								    <a href="#">Manutenção</a> 
								   
								    <ul class="menu1">
								    
								       
								        <li class="menuItem">    
								    
										 	<a href="<%=getBaseLink("/horas/horasmanutencaofiltrocons.asp")%>">Horas</a>
										</li>
										
										<li class="menuItem">
										 	<a href="<%=getBaseLink("/manutencao/usuariosmanutencaofiltro.asp")%>">Consultar Horas</a>
										</li>
										
								      
                                      
                                        
										
               							  </ul>
								</li>
								
								<%
								
							End If
End If
							
							
							%>
								   
									<%    ''''''''''''''''''''''''  %>

								        <%    ''''''''''''''''''''''''  %>
								        
								        
								        
								   
									<%    ''''''''''''''''''''''''  %>


								        

							
							<li class="menuItem menu-center3">
                                <a href="#">Lançamento Horas</a> 
                                <ul class="menu2">
									<li class="menuItem">
									 	<a href="<%=getBaseLink("/horas/horaslancamento.asp")%>">Lançamento</a>
									</li>
							
									
									<%
									
									If Not Session("sint_TipoUsuario") = "0" Then
										
										%>
										
										<li class="menuItem">
										 	<a href="<%=getBaseLink("/horas/horasrelatoriofiltro.asp")%>">Consulta</a>
										</li>
										
										<%
										
									End If
									
									If Session("sboo_fladministrador") = True Then
										
										%>
										<!--
										<li class="menuItem">
										 	<a href="<%=getBaseLink("/horas/iprelatoriofiltro.asp")%>">Consulta IP</a>
										</li>
										-->
										<%
										
									End If
									End If
									
									
									%>
								
                               </ul>
                            </li>
                                   <%
                            	If Session("sboo_fladministrador") = True Then
                                   %>
                            
                            
							<li class="menuItem menu-center3">
                                <a href="#">Lançamento Horas</a> 
                                <ul class="menu2">
									<li class="menuItem">
									 	<a href="<%=getBaseLink("/horas/horaslancamento.asp")%>">Lançamento</a>
									</li>
							
									
									<%
									
									If Not Session("sint_TipoUsuario") = "0" Then
										
										%>
										
										<li class="menuItem">
										 	<a href="<%=getBaseLink("/horas/horasrelatoriofiltro.asp")%>">Consulta</a>
										</li>
										
										<%
										
									End If
									
									If Session("sboo_fladministrador") = True Then
										
										%>
										<!--
										<li class="menuItem">
										 	<a href="<%=getBaseLink("/horas/iprelatoriofiltro.asp")%>">Consulta IP</a>
										</li>
										-->
										<%
										
									End If
									'End If
									
									%>
								
                               </ul>
                            </li>
                            
                            
                            <%
                            
                           End If
                            
                             %>
                            
                            
                           
                           
                           
                            <li class="menuItem menu-center2">
                                <a href="<%=getBaseLink("/manutencao/alterardados.asp")%>">Alterar Dados</a>
                            </li>
                            <li class="menuItem menuItemLast">
                                <a href="<%=getBaseLink("/manutencao/muralRecado.asp")%>">Mural de Recado <%
									
									If VerificarNovaMensagem(Session("sstr_IdUsuario")) = True Then
										
										%><img src="../images/new.gif" title="Há uma nova mensagem" /><%
										
									End If
									
									
								%>
								
								 
                           
								
								</a> 
                            </li>
                        </ul>
                    </div>
                </div> 
		    </div>
            <script type=text/javascript>
				<!--
					new DropMenu("Master_Header_Menu");
				//-->
			</script>
		    <div style="height:300; text-align:center">