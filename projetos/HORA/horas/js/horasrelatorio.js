function voltar(){	
	document.location = "horasrelatoriofiltro.asp"
}

function imprimir(){
	document.thisForm.submit();
}

function AbrirDetalhe(objForm,pstr_Data){
	objForm.hdnDtData.value = pstr_Data
	objForm.pstr_Operacao.value = "A"
	objForm.action = "horasdetalhe.asp";
	document.thisForm.submit();
}

function imprimirLote(objForm){
	
	objForm.pstr_Operacao.value = "V"
	objForm.action = "horasrelatorioprintlote.asp";	
	objForm.submit();
}