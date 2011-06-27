<!--#include file = "../Inc/Syscode.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>管理员登录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../inc/admin_style.css" rel="stylesheet" type="text/css">
<script language=javascript>
<!--
function SetFocus()
{
if (document.form.ClassName.value=="")
	document.form.ClassName.focus();
else
	document.form.ClassName.select();
}
function CheckForm()
{
	if(document.form.ClassName.value=="")
	{
		alert("请输入类别名称！");
		document.form.ClassName.focus();
		return false;
	}
}
//-->
</script>
</head>

<body>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p><%
if request("action")="add" then
call AddNsClass()
else 
call ModifyNsClass()
end if
sub AddNsClass()
%> <form name="form" action="NewsClass_Save.asp?action=add" method="post" onSubmit="return CheckForm();"> 
<table width="346" border="0" align="center" class="tableborder"> <tr align="center"> 
<td height="25" colspan="2" class="title">添加大类</td></tr> <tr> <td width="200" height="25" align="right" class="tdbg">大类名称：</td><td width="200" class="tdbg"><input name="ClassName" type="text" id="ClassName"></td></tr> 
<tr align="center"> <td colspan="2" class="tdbg"> <input   type="submit" name="Submit" value="添 加 " style="font-size: 9pt; height: 19; width: 60"> 
&nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.history.back()" style="font-size: 9pt; height: 19; width: 60"> 
</td></tr> </table></form><%
end sub
sub ModifyNsClass()
  dim rs
  dim strsql
  set rs=server.CreateObject("adodb.recordset")
  strsql="select * from News_Class where ClassID="&request("ClassID")
  rs.open strsql,conn,1,1

%> <form name="form" action="NewsClass_Save.asp?action=modify" method="post" onSubmit="return CheckForm();"> 
<table width="346" border="0" align="center" class="tableborder"> <tr align="center"> 
<td height="25" colspan="2" class="title">大类修改</td></tr> <tr> <td width="200" height="25" align="right" class="tdbg">大类名称：</td><td width="200" class="tdbg"><input name="ClassName" type="text" id="ClassName" value="<%=rs("ClassName")%>"></td></tr> 
<tr align="center"> <td colspan="2" class="tdbg"><input type="hidden" name="ClassID" value="<%=rs("ClassID")%>"> 
<input   type="submit" name="Submit" value="修 改" style="font-size: 9pt; height: 19; width: 60"> 
&nbsp;<input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.history.back()" style="font-size: 9pt; height: 19; width: 60"> 
</td></tr> </table></form><%end sub %> 
</body>
</html>
