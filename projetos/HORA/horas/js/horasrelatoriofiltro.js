function voltar(){	
	document.location = "horaslancamento.asp"
}

function listar(){
	document.thisForm.hdnExecutar.value = "LISTA"
	document.thisForm.submit();
}