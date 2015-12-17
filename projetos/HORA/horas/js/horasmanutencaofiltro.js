function voltar(){	
	document.location = "horaslancamento.asp"
}

function enviar(){
	document.thisForm.hdnExecutar.value = "LISTAR"
	document.thisForm.submit();
}