<SCRIPT LANGUAGE="javascript" src="<%=getBaseLink("/js/menu.js")%>"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
	function OnMouseOver(subGroup, direction){
		event.srcElement.className="MenuItemOver";
		if (document.readyState == 'complete') 
			aspnm_itemMsOver(event.srcElement.id, subGroup, direction, 5, 0, 0, null);
	}
	
	function OnMouseOut(superItem, subGroup){
		event.srcElement.className='MenuItem';
		if (document.readyState == 'complete') 
			aspnm_itemMsOut(event.srcElement.id, superItem, subGroup, 0, null);
	}
	
</SCRIPT>

<span id="Menu1">
	<!-- Tabela do Menu Principal -->
	<table id="tblMenu" class="TopMenuGroup" cellspacing="1" cellpadding="0" border="0" onmouseout="if (document.readyState == 'complete') aspnm_groupMsOut('tblMenu', null, null, 0);" style="z-index:999; width: 100%; display: none;">
		<tr>
			
			<%
			
			If Session("sboo_fladministrador")	= True Then
				
				%>
				
				<td id="tdMenu_CV"  class="MenuItem" onmouseover="OnMouseOver('SubMenu_CV', 'belowleft');" onmouseout="OnMouseOut('tblMenu', 'SubMenu_CV');" onmousemove="return false;" ondblclick="return false;" style="display:'none';" nowrap>Manutenção</td>
				
				<%
				
			End If
			
			%>
				
			<td id="tdMenu_HR"  class="MenuItem" onmouseover="OnMouseOver('SubMenu_HR', 'belowleft');" onmouseout="OnMouseOut('tblMenu', 'SubMenu_HR');" onmousemove="return false;" ondblclick="return false;" style="display:'none';" nowrap>Lançamento de Horas</td>
			<td id="tdMenu_TS"	class="MenuItem" onmouseover="this.className='MenuItemOver';"		   onmouseout="this.className='MenuItem';"			 onmousemove="return false;" ondblclick="return false;" onclick="javascript:fctChamaURL(1,'<%=getBaseLink("/manutencao/trocarsenha.asp")%>')" style="display:'';" nowrap>Trocar Senha</td>
			<td id="tdMenu_ML"	class="MenuItem" onmouseover="this.className='MenuItemOver';"		   onmouseout="this.className='MenuItem';"			 onmousemove="return false;" ondblclick="return false;" onclick="javascript:fctChamaURL(1,'<%=getBaseLink("/manutencao/amigosecreto.asp")%>')" style="display:'';" nowrap>Mural de Recados</td>
			<td width="100%" align="right" class="MenuItem">
			<%=Session("sstr_IdUsuario") & " - " & Session("sstr_DsUsuario") & "&nbsp;"%>
			</td>
		</tr>
	</table>

</span>
<span>
	<table id="SubMenu_HR" class="MenuGroup" cellspacing="1" cellpadding="0" border="0" onmouseover="aspnm_groupMsOver('SubMenu_HR')" onmouseout="if (document.readyState == 'complete') aspnm_groupMsOut('SubMenu_HR', 'tdMenu_HR', 'tblMenu', 0, null);" style="z-index:999;position:absolute;visibility:hidden;z-index:999;left:0px;top:0px;">
		<tr>
			<td id="tdSubmenu_Lancamento" class="MenuItem" onmouseover="this.className='MenuItemOver';" onmouseout="this.className='MenuItem';" onmousemove="return false;" ondblclick="return false;" onclick="fctChamaURL(1,'<%=getBaseLink("/horas/horaslancamento.asp")%>')" style="display:''; width:100px">Lançamento</td>
		</tr>
		<tr>
			
			<%
			
			If Session("sboo_fladministrador")	= True Or Session("sint_TipoUsuario") = "2" Then
				
				%>
				
				<td id="tdSubmenu_Consulta" class="MenuItem" onmouseover="this.className='MenuItemOver';" onmouseout="this.className='MenuItem';" onmousemove="return false;" ondblclick="return false;" onclick="fctChamaURL(1,'<%=getBaseLink("/horas/horasrelatoriofiltro.asp")%>')" style="display:'none'; width:100px">Consulta</td>
				
				<%
				
			End If
			
			
			%>
			
		</tr>
	</table>
	
	<%
	
	If Session("sboo_fladministrador")	= True Then
		
		%>
		
		<table id="SubMenu_CV" class="MenuGroup" cellspacing="1" cellpadding="0" border="0" onmouseover="aspnm_groupMsOver('SubMenu_CV')" onmouseout="if (document.readyState == 'complete') aspnm_groupMsOut('SubMenu_CV', 'tdMenu_CV', 'tblMenu', 0, null);" style="z-index:999;position:absolute;visibility:hidden;z-index:999;left:0px;top:0px;">
			<tr>
				<td>
					<table id="tdSubmenu_Despachante" class="MenuItem" cellpadding="0" cellspacing="0" border="0" width="100%" height="100%" onmouseover="if (document.readyState == 'complete') aspnm_updateCell('tdSubmenu_Despachante', 'MenuItemOver', null, '', 'Menu1_icon_1', '', 'over');if (document.readyState == 'complete') aspnm_itemMsOver('tdSubmenu_Despachante', 'tblGrupo_Despachante', 'rightdown', 0, 0, 0, null);" onmouseout="if (document.readyState == 'complete') aspnm_updateCell('tdSubmenu_Despachante', 'MenuItem', null, '', 'Menu1_icon_1', '/horas/images/arrow_black.gif', 'out');if (document.readyState == 'complete') aspnm_itemMsOut('tdSubmenu_Despachante', 'Menu1_group_2', 'tblGrupo_Despachante', 0, null);" onmousedown onmouseup onmousemove="return false;" ondblclick="return false;" style="display:none">
						<tr>
							<td>Cadastro Básico</td>
							<td align="right" style="padding:0;" width="0"><img id="Menu1_icon_1" src="<%=getBaseLink("/images/arrow_black.gif")%>" border="0" WIDTH="15" HEIGHT="10"></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td>
					<table id="tdSubmenu_Relacao" class="MenuItem" cellpadding="0" cellspacing="0" border="0" width="100%" height="100%" onmouseover="if (document.readyState == 'complete') aspnm_updateCell('tdSubmenu_Relacao', 'MenuItemOver', null, '', 'Menu1_icon_1', '', 'over');if (document.readyState == 'complete') aspnm_itemMsOver('tdSubmenu_Relacao', 'tblGrupo_Relacao', 'rightdown', 0, 0, 0, null);" onmouseout="if (document.readyState == 'complete') aspnm_updateCell('tdSubmenu_Relacao', 'MenuItem', null, '', 'Menu1_icon_1', '/horas/images/arrow_black.gif', 'out');if (document.readyState == 'complete') aspnm_itemMsOut('tdSubmenu_Relacao', 'Menu1_group_2', 'tblGrupo_Relacao', 0, null);" onmousedown onmouseup onmousemove="return false;" ondblclick="return false;" style="display:none">
						<tr>
							<td>Relação</td>
							<td align="right" style="padding:0;" width="0"><img id="Menu1_icon_1" src="<%=getBaseLink("/images/arrow_black.gif")%>" border="0" WIDTH="15" HEIGHT="10"></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td id="tdSubmenu_MHoras" class="MenuItem" onmouseover="this.className='MenuItemOver';" onmouseout="this.className='MenuItem';" onmousemove="return false;" ondblclick="return false;" onclick="fctChamaURL(1,'<%=getBaseLink("/horas/horasmanutencaofiltro.asp")%>')" style="display:''; width:100px">Horas</td>
			</tr>
			<!--<tr>
				<td id="tdSubmenu_MHoras" class="MenuItem" onmouseover="this.className='MenuItemOver';" onmouseout="this.className='MenuItem';" onmousemove="return false;" ondblclick="return false;" onclick="fctChamaURL(1,'<%=getBaseLink("/horas/horasalmoco.asp")%>')" style="display:''; width:100px">Horario de Almoço</td>
			</tr>-->
		</table>
		<table id="tblGrupo_Despachante" class="MenuGroup" cellspacing="1" cellpadding="0" border="0" onmouseover="aspnm_groupMsOver('tblGrupo_Despachante')" onmouseout="if (document.readyState == 'complete') aspnm_groupMsOut('tblGrupo_Despachante', 'tdSubmenu_Despachante', 'Menu1_group_2', 0, null);" style="z-index:999;position:absolute;visibility:hidden;z-index:999;left:0px;top:0px;">
			<tr>
				<td id="tdSubmenu_Area" class="MenuItem" onmouseover="this.className='MenuItemOver';" onmouseout="this.className='MenuItem';" onmousemove="return false;" ondblclick="return false;" onclick="fctChamaURL(1,'<%=getBaseLink("/manutencao/arealistagem.asp")%>')" style="display:''; width:100px">Áreas do Banco</td>
			</tr>
			<tr>
				<td id="tdSubmenu_Atividades" class="MenuItem" onmouseover="this.className='MenuItemOver';" onmouseout="this.className='MenuItem';" onmousemove="return false;" ondblclick="return false;" onclick="fctChamaURL(1,'<%=getBaseLink("/manutencao/atividadelistagem.asp")%>')" style="display:''; width:100px">Atividades</td>
			</tr>
			<tr>
				<td id="tdSubmenu_Funcoes" class="MenuItem" onmouseover="this.className='MenuItemOver';" onmouseout="this.className='MenuItem';" onmousemove="return false;" ondblclick="return false;" onclick="fctChamaURL(1,'<%=getBaseLink("/manutencao/funcaolistagem.asp")%>')" style="display:''; width:100px">Função/Cargo</td>
			</tr>
		</table>
		<table id="tblGrupo_Relacao" class="MenuGroup" cellspacing="1" cellpadding="0" border="0" onmouseover="aspnm_groupMsOver('tblGrupo_Relacao')" onmouseout="if (document.readyState == 'complete') aspnm_groupMsOut('tblGrupo_Relacao', 'tdSubmenu_Relacao', 'Menu1_group_2', 0, null);" style="z-index:999;position:absolute;visibility:hidden;z-index:999;left:0px;top:0px;">
			<tr>
				<td id="tdSubmenu_Projetos" class="MenuItem" onmouseover="this.className='MenuItemOver';" onmouseout="this.className='MenuItem';" onmousemove="return false;" ondblclick="return false;" onclick="fctChamaURL(1,'<%=getBaseLink("/manutencao/projetoslistagem.asp")%>')" style="display:''; width:150px">Projetos</td>
			</tr>
			<tr>
				<td id="tdSubmenu_Usuarios" class="MenuItem" onmouseover="this.className='MenuItemOver';" onmouseout="this.className='MenuItem';" onmousemove="return false;" ondblclick="return false;" onclick="fctChamaURL(1,'<%=getBaseLink("/manutencao/usuarioslistagem.asp")%>')" style="display:''; width:150px">Usuários</td>
			</tr>
			<tr>
				<td id="tdSubmenu_UsuariosProjetos" class="MenuItem" onmouseover="this.className='MenuItemOver';" onmouseout="this.className='MenuItem';" onmousemove="return false;" ondblclick="return false;" onclick="fctChamaURL(1,'<%=getBaseLink("/manutencao/usuariosprojetos.asp")%>')" style="display:''; width:150px">Usuários em Projetos</td>
			</tr>
		</table>
		
		<%
	
	End If
	
	%>

<%
On error resume next
%>
<script LANGUAGE="javascript">
<!--
	tblMenu.style.display = "";
	tdMenu_HR.style.display = "";
	tdMenu_TS.style.display = "";
	SubMenu_HR.style.display = "";
	

<%	if Session("sint_TipoUsuario") = "1" then %> 


		tdMenu_CV.style.display = "";
		SubMenu_CV.style.display = "";
		tdSubmenu_Despachante.style.display = "";
		tdSubmenu_Relacao.style.display = "";
		tdSubmenu_Consulta.style.display = "";
		
		
<%	end if %>


<%	if Session("sint_TipoUsuario") = "2" then %> 


		tdSubmenu_Consulta.style.display = "";

		
		
<%	end if %>

	
//-->
</script>
<%on error goto 0%>

</span>
