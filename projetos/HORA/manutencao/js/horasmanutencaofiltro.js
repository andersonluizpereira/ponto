function voltar(){	
	document.location = "usuariosmanutencaofiltro.asp"
}

function enviar(){
	document.thisForm.hdnExecutar.value = "LISTAR"
	//document.thisForm.vstr_Operacao.value = "V"
	document.thisForm.submit();
}