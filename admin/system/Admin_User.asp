<!--#include file="../Inc/Syscode.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
Dim action,UserID,UserName,Password,ConfirmPassword,Purview
Dim Rs,Sql
UserID = request("UserID")
UserName = trim(request("UserName"))
Password = trim(request("Password"))
ConfirmPassword = trim(request("ConfirmPassword"))
action = trim(request("action"))
Purview = request("Purview")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<SCRIPT language=JavaScript src="../../Inc/JChar.js"></SCRIPT>
<SCRIPT language=JavaScript>
  function CheckForm()
  {
      var msg="";
	if (document.myform.UserName.value=="") 
	{
		msg="对不起，用户名不能为空，且不能超过30个字符!";
		alert(msg);
	 	document.myform.UserName.focus();
	return false;
	}
   
	if (!IsValidString(document.myform.Password.value))
	{
		msg="对不起，密码必须为4个以上的数字或字母！\n并且不能使用汉字当作密码！";
		alert(msg);
	 	document.myform.Password.focus();
	return false;
	}
	if (document.myform.Password.value!==document.myform.ConfirmPassword.value)
	{
		msg="对不起，您输入的密码不一致！\n请重新输入密码！";
		alert(msg);
	 	document.myform.Password.focus();
	return false;
	}
 }
</SCRIPT>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/admin_style.css" rel="stylesheet" type="text/css">
</head>

<body>
<br> <br> <br> <br> <%if action="add" then
call useradd()
elseif action="addok" then
call usersave()
elseif action="modify" then
call modify()
elseif action="modifyok" then
call modifyok()
end if
sub useradd()
%> <table width="80%" border="0"> <form name="myform" action="Admin_User.asp" method="post"> 
<tr> <td width="50%" align="right">用户名：</td><td width="50%"><input name="UserName" type="text" id="UserName"></td></tr> 
<tr> <td align="right">密码：</td><td><input name="Password" type="password" id="Password"></td></tr> 
<tr> <td align="right">密码确认： </td><td><input name="ConfirmPassword" type="password" id="ConfirmPassword"></td></tr> 
<tr> <td colspan="2" align="center"><INPUT TYPE="hidden" NAME="purview" value="1"><input type="hidden" name="action" value="addok"> 
<input type="submit" name="Submit" value="&nbsp;提&nbsp;交&nbsp;"  style="font-size: 9pt; height: 19; width: 60"> 
&nbsp; <input name="Cancel" type="button" id="Cancel" value="&nbsp;取&nbsp;消&nbsp;" onClick="window.location.href='User_Manage.asp'" style="font-size: 9pt; height: 19; width: 60"></td></tr> 
</form></table><%end sub
sub Modify()
Dim UserRs,UserSql
set UserRs = Server.CreateObject("Adodb.Recordset")
      UserSql = "select * from admin where UserID ="&Request("UserID")&" "
	  UserRs.Open UserSql,Conn,1,1
%> <table width="80%" border="0"> <form action="Admin_User.asp" method="post" onsubmit="return CheckForm()" name="myform"> 
<tr> <td width="50%" align="right">用户名：</td><td width="50%"><input name="UserName" type="text" id="UserName" value="<%=UserRs("UserName")%>" readonly></td></tr> 
<tr> <td align="right">密码：</td><td><input name="Password" type="password" id="Password"></td></tr> 
<tr> <td align="right">密码确认： </td><td><input name="ConfirmPassword" type="password" id="ConfirmPassword"></td></tr> 
<tr align="center"> <td colspan="2"><input name="userid" type="hidden" id="userid" value="<% = UserRs("UserID")%>"> 
<input type="hidden" name="action" value="modifyok"> <input type="submit" name="Submit" value="&nbsp;提&nbsp;交&nbsp;"  style="font-size: 9pt; height: 19; width: 60"> 
&nbsp; <input name="Cancel" type="button" id="Cancel" value="&nbsp;取&nbsp;消&nbsp;" onClick="window.location.href='User_Manage.asp'" style="font-size: 9pt; height: 19; width: 60"></td></tr> 
</form></table><%
UserRs.close
set UserRs=nothing
end sub
sub usersave()
set rs=server.CreateObject("adodb.recordset")
sql="select * from admin"
rs.open sql,conn,1,3
rs.addnew
rs("username") = username
rs("password") = md5(password)
rs("Purview") = Purview
rs.update
rs.close
set rs=nothing
call ShowSuccess("恭喜您，管理员添加成功！","user_manage.asp")
end sub

sub modifyok()
set rs=server.CreateObject("adodb.recordset")
sql="select * from admin where UserID ="&UserID&" "
rs.open sql,conn,1,3
if password="" and ConfirmPassword="" then
rs("password")=rs("password")
else
rs("password")=md5(password)
end if
if Password <> ConfirmPassword then
     Call ShowError ("对不起，密码不一致！")
end if
rs("Purview")=Purview
rs.update
rs.close
set rs=nothing
call ShowSuccess("恭喜您，资料修改成功！","user_manage.asp")
end sub
%> 
</body>
</html>
