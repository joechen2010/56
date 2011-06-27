<!--#include file = "../Inc/Syscode.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>添加小类</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../inc/admin_style.css" rel="stylesheet" type="text/css">
<script language=javascript>
<!--
function SetFocus()
{
if (document.myform.NClassName.value=="")
	document.myform.NClassName.focus();
else
	document.myform.UserName.select();
}
function CheckForm()
{
	if(document.myform.NClassName.value=="")
	{
		alert("请输入栏目名称！");
		document.myform.NClassName.focus();
		return false;
	}
}
//-->
</script>
</head>

<body>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<%
dim rs,sql,ClassID
ClassID=trim(request("ClassID"))
if request("action")="add"  then
    call addNsNClass()
elseif request("action")="modify" then
    call modifyNsNClass()
end if
%>
<%sub addNsNClass()
    set rs=server.CreateObject("adodb.recordset")
        sql="select * from News_Class order by ClassID desc"
	   rs.open sql,conn,1,1
%>
<form name="myform" action="NewsNClass_Save.asp?action=add" method="post" onSubmit="return CheckForm();">
  <table width="346" border="0" align="center" class="tableborder">
    <tr> 
      <td height="25" colspan="2" align="center" class="title">添加小类</td>
    </tr>
    <tr> 
      <td width="200" height="25" align="right" class="tdbg">一级栏目：</td>
      <td width="200" class="tdbg"><select name="ClassID" size="1" id="ClassID" style="font-size: 14px">
	      	  <%While Not rs.EOF%>
          <option value="<%=rs("ClassID")%>" <%if trim(rs("ClassID"))= classid then response.write "selected"%>><% =rs("ClassName")%> 
          </option>
		  <%
            rs.MoveNext
  			Wend
			rs.close
			set rs=nothing
 	      %>
          </select></td>
    </tr>
    <tr> 
      <td align="right" class="tdbg">二级栏目：</td>
      <td class="tdbg"><input name="NClassName" type="text" id="NClassName"> </td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="tdbg"> <input   type="submit" name="Submit" value="添 加 " style="font-size: 9pt; height: 19; width: 60"> 
        &nbsp; <input name="reset" type="reset"  id="reset" value="取 消" style="font-size: 9pt; height: 19; width: 60" onClick="window.history.back()"> 
      </td>
    </tr>
  </table>
</form>
<%end sub%>
<%sub modifyNsNClass()
dim OldNClassID,OldNClassName,OldClassID,OldClassName
set rs=server.CreateObject("adodb.recordset")
    sql="select NClassID,NClassName,ClassID,(select ClassName from News_Class where ClassID=A.ClassID) as ClassName from News_NClass as A where A.NClassID="&request("NClassID")
    rs.open sql,conn,1,1
	OldNClassID=rs("NClassID")
	OldNClassName=rs("NClassName")
	OldClassID=rs("ClassID")
	OldClassName=rs("ClassName")
	rs.close
	set rs=nothing
%>
<form name="myform" action="NewsNClass_Save.asp?action=modify" method="post" onSubmit="return CheckForm();">
  <table width="346" border="0" align="center" class="tableborder">
    <tr> 
      <td height="25" colspan="2" align="center" class="title">修改类</td>
    </tr>
    <tr> 
      <td width="200" height="25" align="right" class="tdbg">一级栏目：</td>
      <td width="200" class="tdbg"><select name="ClassID" size="1" id="ClassID" style="font-size: 14px">
	      <option selected value="<%response.write OldClassID%>"><%response.write OldClassName%></option>
	      <% dim rs_class
		  set rs_class=server.CreateObject("adodb.recordset")
		  sql="select * from News_Class"
		  rs_class.open sql,conn,1,1
		  While Not rs_class.EOF%>
          <option value="<% =rs_class("ClassID")%>"><% =rs_class("ClassName")%> 
          </option>
		  <%
            rs_class.MoveNext
  			Wend
			rs_class.close
			set rs_class=nothing
 	      %>
          </select></td>
    </tr>
    <tr> 
      <td align="right" class="tdbg">二级栏目：</td>
      <td class="tdbg"><input name="NClassName" type="text" id="NClassName" value="<%response.write OldNClassName%>"> </td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="tdbg"><input type="hidden" value="<%response.write OldNClassID%>" name="NClassID"> <input   type="submit" name="Submit" value="修 改 " style="font-size: 9pt; height: 19; width: 60"> 
        &nbsp; <input name="reset" type="reset"  id="reset" value="取 消" style="font-size: 9pt; height: 19; width: 60" onClick="window.history.back()"> 
      </td>
    </tr>
  </table>
</form>
<%end sub%>
</body>
</html>
