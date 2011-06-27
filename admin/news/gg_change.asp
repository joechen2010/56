<!--#include file = "../Inc/Syscode.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from gonggao where id="&request("id")&""
rs.open sql,conn,1,1
%>
<title></title>
	<Script Language=JavaScript>
		// 表单提交客户端检测
	function doSubmit(){
		if (document.myform.bt.value==""){
			alert("新闻标题不能为空！");
			return false;
		}
		// getHTML()为eWebEditor自带的接口函数，功能为取编辑区的内容
		if (eWebEditor1.getHTML()==""){
			alert("新闻内容不能为空！");
			return false;
		}
		document.myform.submit();
	}
	</Script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/admin_style.css" rel="stylesheet" type="text/css">
</head>

<body>

  <form action="gg_changeok.asp" method="post" name="myform">
  <table width="650" align=center cellspacing=3>
    <tr> 
      <td width="71" align="right">标题：</td>
      <td width="564"> <input type="text" name="bt" value="<%=rs(1)%>" size="35"></td>
    </tr>
    <tr> 
      <td align="right">内容：</td>
      <td> <input name="Content" type="hidden" value="<%=Server.HtmlEncode(rs(2))%>"> 
        <iframe id="eWebEditor1" src="../Editor/ewebeditor.asp?id=Content&style=" frameborder="0" scrolling="no" width="550" height="350"></iframe>      </td>
    </tr>
  </table>
  <input name="ID" type="hidden" value="<%=rs(0)%>">
  <p align=center><input type=button name=btnSubmit value=" 提 交 " onClick="doSubmit()" style="font-size: 9pt; height: 19; width: 60"> <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='News_Manage.asp'" style="font-size: 9pt; height: 19; width: 60"> </p>
</form>


</body>
</html>
