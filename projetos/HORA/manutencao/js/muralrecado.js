function voltar() {	
	document.location = "../horas/horaslancamento.asp"
}

function enviar() {
	document.thisForm.pstr_Operacao.value = "I"
	document.thisForm.submit();
}

function excluir() {
	
	var iframe = iframeMensagem.document.getElementsByName("chkExcluirMensagem");
	
	var cont = 0;
	for(i=0; i < iframe.length ; i++){
		
		if(iframe[i].checked==true){
			
			
			if(cont == 0){
				
				document.thisForm.hdnExcluir.value = document.thisForm.hdnExcluir.value + iframe[i].value
				
			}else{
				
				document.thisForm.hdnExcluir.value = document.thisForm.hdnExcluir.value + "," + iframe[i].value
				
			}
			cont = cont + 1
		}
	}
	
	document.thisForm.pstr_Operacao.value = "EXCLUIR"
	document.thisForm.submit();
}