function voltar(objForm) {	
	objForm.action = "horasmanutencao.asp";
	objForm.submit();
}
			
function alterarLote() {
				
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
			
function excluir(pint_Tipo, pstr_Data) {
				
	apagar = confirm("Deseja realmente apagar ?");
				
	if(apagar == true){
		document.thisForm.hdnExecutar.value = "EXCLUIR";
		document.thisForm.pstr_Operacao.value = "E";
		document.thisForm.hdnDsData.value = pstr_Data;
		document.thisForm.hdnFlTipo.value = pint_Tipo;
		document.thisForm.submit();
					
	}
}