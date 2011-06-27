<!--#include file = "../Inc/Syscode.asp"-->
<%
if request("action")="add" then
  set rs=server.CreateObject("adodb.recordset")
  sql="select * from advertise "
  rs.open sql,conn,1,3
  rs.addnew
  rs("adclass")=trim(request("adclass"))
  rs("webpic")=trim(request("bigimg"))
  rs("webname")=trim(request("webname"))
  rs("weburl")=trim(request("weburl"))
  rs("addtime")=now()
  rs.update
  rs.close
  set rs=nothing
  response.Write "<script>alert('添加成功');location.href='add.asp';</script>"
 elseif request("action")="modify" then 
  set rs=server.CreateObject("adodb.recordset")
  sql="select * from advertise where id="&request("id")&""
  rs.open sql,conn,1,3
  rs("adclass")=trim(request("adclass"))
  rs("webpic")=trim(request("bigimg"))
  rs("webname")=trim(request("webname"))
  rs("weburl")=trim(request("weburl"))
  rs("addtime")=now()
  rs.update
  rs.close
  set rs=nothing
  response.Write "<script>alert('修改成功');location.href='ad_manage.asp?page="&request("page")&"';</script>"  
end if

if request("manage")="" then 
    call add()
  elseif request("manage")="modify" then
    call modify()
end if
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../inc/Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" bgcolor="#e1eefd">
<p>&nbsp;</p>
<%sub add()%>
<table width="90%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#6699ff">
  <form method="POST" action="add.asp">
  <tr align="center"> 
    <td height="27" colspan="2" background="images/Left01.gif"><b>添 加 广 告 信 息</b></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="30" align="right">版块选择：</td>
    <td height="30">
                          <select name="adclass">
                            <option value="1">找货找车</option>
							<option value="2">货运专线</option>
							<option value="3">客运专线</option>
							<option value="4">车辆档案</option>
							<option value="5">修理配件</option>
							<option value="6">库房信息</option>
							<option value="7">招聘求职</option>
							<option value="8">供求信息</option>
                          </select>	</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">公司名称：</td>
      <td height="30" width="81%">
      <input name="webname" type="text" id="webname" style="border: 1px outset #C0C0C0" size="20">    </td>
  </tr>  
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">广告图片：</td>
      <td height="30" width="81%">
 <iframe name="BigImg" frameborder=0 width=450 height=20 scrolling=no src=BigUpload.asp></iframe> 
            <br>
            图片路径： 
            <input name="BigImg" type="TEXT" id="BigImg"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100 readonly>  
	  </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="28" width="19%" align="right">链接网址：</td>
      <td width="81%" height="28">
	  <input name="weburl" type="text" id="weburl" style="border: 1px outset #C0C0C0" value="http://" size="20"></td>
  </tr>
    <tr bgcolor="#FFFFFF" align="center"> 
      <td height="30" colspan="2"> 
        <input type="submit" value=" 添 加 "  style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="reset" value=" 重 置 " name="B2" style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="hidden" name="action" value="add">      
	</td>
  </tr>
  </form>
</table>
<%end sub%>

<%sub modify()%>
<%
set web=conn.execute("select * from advertise where id="&request("id")&"")
if not (web.eof and web.bof) then

	adclass=web("adclass")
	select case adclass
	  case "1"
	  adtype="找货找车"
	  case "2"
	  adtype="货运专线"
	  case "3"
	  adtype="客运专线"
	  case "4"
	  adtype="车辆档案"
	  case "5"
	  adtype="修理配件"	  
	  case "6"
	  adtype="库房信息"
	  case "7"
	  adtype="招聘求职"
	  case "8"
	  adtype="供求信息"	  	  
	end select
%>
<table width="90%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#6699ff">
  <form method="POST" action="add.asp">
  <tr align="center"> 
    <td height="27" colspan="2" background="images/Left01.gif"><b>修 改 广 告 信 息</b></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="30" align="right">版块选择：</td>
    <td height="30">
                          <select name="adclass">
						    <option value="<%=web("adclass")%>"><%=adtype%></option>
                            <option value="1">找货找车</option>
							<option value="2">货运专线</option>
							<option value="3">客运专线</option>
							<option value="4">车辆档案</option>
							<option value="5">修理配件</option>
							<option value="6">库房信息</option>
							<option value="7">招聘求职</option>
							<option value="8">供求信息</option>
                          </select>	</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">公司名称：</td>
      <td height="30" width="81%">
      <input name="webname" type="text" value="<%=web("webname")%>" style="border: 1px outset #C0C0C0" size="20">    </td>
  </tr>  
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">广告图片：</td>
      <td height="30" width="81%">
 <iframe name="BigImg" frameborder=0 width=450 height=20 scrolling=no src=BigUpload.asp></iframe> 
            <br>
            图片路径： 
            <input name="BigImg" type="TEXT" id="BigImg" value="<%=web("webpic")%>"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100 readonly>  
	  </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="28" width="19%" align="right">链接网址：</td>
      <td width="81%" height="28">
	  <input name="weburl" type="text" id="weburl" value="<%=web("weburl")%>" style="border: 1px outset #C0C0C0" size="20"></td>
  </tr>
    <tr bgcolor="#FFFFFF" align="center"> 
      <td height="30" colspan="2"> 
    <input type="submit" value="提交"  style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="hidden" name="action" value="modify">
	<input type="hidden" name="id" value="<%=web("id")%>">
	<input type="hidden" name="page" value="<%=request("page")%>">      
	</td>
  </tr>
  </form>
</table>
<%
web.close
set web=nothing
end if
end sub
%>
</body>
</html>
