function voltar(){	
	document.location = "../horas/horaslancamento.asp"
}

function listar(){
	document.thisForm.hdnExecutar.value = "LISTA"
	document.thisForm.submit();
}