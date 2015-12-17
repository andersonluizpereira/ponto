function voltar(objForm) {	
	objForm.action = "horasmanutencao.asp";
	objForm.submit();
}
			
function registrar() {
				
	document.thisForm.hdnExecutar.value = "REGISTRAR";
	document.thisForm.pstr_Operacao.value = "I"
	document.thisForm.submit();
}
			
function alterar() {
				
	document.thisForm.hdnExecutar.value = "ALTERAR";
	document.thisForm.pstr_Operacao.value = "A"
	document.thisForm.submit();
}

function alterar2() {
				
	document.thisForm.hdnExecutar.value = "ALTERAR";
	document.thisForm.pstr_Operacao.value = "A"
	document.thisForm.submit();
}

			
function fMudaCorLinha(objTR){
	var arrayLinhas = document.getElementsByName("trLinhaRegistro");
	for(i=0;i<arrayLinhas.length;i++){
		arrayLinhas[i].style.color = "#00559A";
	}
	objTR.style.color = "red";
}
			
function excluir(pint_Tipo) {
				
	apagar = confirm("Deseja realmente apagar ?");
				
	if(apagar == true){
					
		document.thisForm.hdnExecutar.value = "EXCUIR";
		document.thisForm.pstr_Operacao.value = "E";
		document.thisForm.hdnFlTipo.value = pint_Tipo;
		document.thisForm.submit();
					
	}
}