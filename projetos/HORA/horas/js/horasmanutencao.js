function voltar(){	
	document.location = "horasmanutencaofiltro.asp"
}

function alterar(pstr_Data){
	document.thisForm.hdnDtData.value = pstr_Data
	document.thisForm.pstr_Operacao.value = "AI"
	document.thisForm.hdnBloqueaData.value = "1"
	document.thisForm.submit();
}

function imprimir(){
	document.thisForm.action = "horasmanutencaoprint.asp";
	document.thisForm.submit();
}


function incluir(){
	document.thisForm.pstr_Operacao.value = "I"
	document.thisForm.submit();
}

function ajustarAlmoco(objForm, pstr_DtData) {
	
	document.thisForm.hdnDtData.value = pstr_DtData
	document.thisForm.hdnExecutar.value = "AJUSTAR_HORA"
	objForm.action = "horasmanutencao.asp";
	objForm.submit();
}

function alterarLote(objForm){
	
	objForm.pstr_Operacao.value = "V"
	objForm.action = "horasmanutencaoalteracao.asp";	
	objForm.submit();
}