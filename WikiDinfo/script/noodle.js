var original_text = "";
var _confirm = true;

function _onload(){
	original_text = document.getElementById("contents").value;
}

function _onunload(){
	if (_confirm && (document.getElementById("contents").value != original_text))
	{
		if (window.confirm("As alterações na foram salvas.\n\nEscolha 'OK' para salvar.\nEscolha 'Cancelar' para descartar as alterações.")){
			document.getElementById("editor").submit();
		}
	}
}
