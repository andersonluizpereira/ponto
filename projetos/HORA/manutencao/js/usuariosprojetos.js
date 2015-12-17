function voltar() {
	document.location = "../horas/horaslancamento.asp"
}
	
function atualizarComboProjeto(){
	
	document.thisForm.cmbComboUsuario.value = ""
	document.thisForm.cmbComboUsuarioProjeto.value = ""
	document.thisForm.submit();
}

function atualizarComboAssociados(){
	
	document.thisForm.submit();
}

function Associar(){
	
	document.thisForm.hdnExecutar.value = "ASSOCIAR"
	document.thisForm.submit();
}

function Remover(){
	
	document.thisForm.hdnExecutar.value = "REMOVER"
	document.thisForm.submit();
}

function AtivoOnOff(){
	document.thisForm.hdnExecutar.value = "MUDAR_ATIVO"
	document.thisForm.submit();
}