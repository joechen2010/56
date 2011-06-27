function checkform(){
	if (document.loginbox.userid.value.length ==0){
		alert("您没有填写用户名。");
		document.loginbox.userid.focus();
		return false;
	}
	if (document.loginbox.password.value.length==0){
		alert("您没有填写密码。");
		document.loginbox.password.focus();
		return false;
	}
	return true;
}	