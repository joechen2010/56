function checkform(){
	if (document.loginbox.userid.value.length ==0){
		alert("��û����д�û�����");
		document.loginbox.userid.focus();
		return false;
	}
	if (document.loginbox.password.value.length==0){
		alert("��û����д���롣");
		document.loginbox.password.focus();
		return false;
	}
	return true;
}	