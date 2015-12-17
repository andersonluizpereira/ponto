// Procedimento desenvolvido para levar o usuário
// até a página de inclusão de registro.
function incluir() {
	document.thisForm.pstr_Operacao.value = "I";
	document.thisForm.submit();
}
		
// Desativa os registros do banco de dados.
function desativar(idRegistro){
	document.thisForm.hdnExecutar.value = 'DESATIVAR';
	document.thisForm.hdnIdRegistro.value = idRegistro;
	document.thisForm.submit();
}

// Ativa os registros do banco de dados.
function ativar(idRegistro){
	document.thisForm.hdnExecutar.value = 'ATIVAR';
	document.thisForm.hdnIdRegistro.value = idRegistro;
	document.thisForm.submit();
}
		
function alterar(idRegistro){
	document.thisForm.hdnIdRegistro.value = idRegistro;
	document.thisForm.pstr_Operacao.value = "A";
	document.thisForm.submit();
}