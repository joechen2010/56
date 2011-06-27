<!--#include file = "../Inc/Syscode.asp"-->
<%

%>
<% if request("dels")<>"" then
   conn.Execute("delete * from danwei where id="&request("dels"))
'删除其它信息
   conn.Execute("delete * from zhaopin where sqmc='"&request("delcompany")&"'")
'   conn.Execute("delete * from ypsj where uname='"&request("delcompany")&"'")
 '  conn.Execute("delete * from peixun where sqmc='"&request("delcompany")&"'")
   'conn.Execute("delete * from dwsq where dwmc='"&request("delcompany")&"'")
'删除结束
   end if 
   
   if request("action")="更新" then
   conn.Execute("update danwei set area='"&request("area")&"' where id="&request("id"))
    conn.Execute("update danwei set enddate='"&request("enddate")&"' where id="&request("id"))

   end if   
%>
<% if request("xgs")<>"" then
   if request("zgs")="末开通"then
   conn.Execute("update danwei set zhige='开通',tdate='"&date()&"' where id="&request("xgs"))
   end if 
   if request("zgs")="开通"then
   conn.Execute("update danwei set zhige='末开通',tdate='"&date()&"' where id="&request("xgs"))
   end if 
   end if
%>
<%uname=session("admin")%>

<html>
<head>
<title>后台管理</title>
<meta content="text/html; charset=gb2312" http-equiv="Content-Type">
<LINK href="images/Admin.css" rel=stylesheet>
</head>

<SCRIPT LANGUAGE="JavaScript">
 <!--//
function check()
{
   if (isNaN(go2to.page.value))
		alert("请正确填写转到页数！");
   else if (go2to.page.value=="")
	     {
		alert("请输入转到页数！");
		 }
   else
		go2to.submit();
}
//-->
</SCRIPT>

<SCRIPT LANGUAGE="JavaScript">
function qy_check()
{
    if (qyss.gjc_.value=="") 
		alert("请输入企业名称！");
   	else
		qyss.submit();
}
//-->
</SCRIPT>
<body leftmargin="0" topmargin="0" bgcolor="#e1eefd">
<table border="0" cellpadding="2" cellspacing="1" width="96%" bordercolordark="#FFFFFF" bgcolor="#6699ff" align="center">
  <tr>
          
    <td colspan="2" height="27" background="Images/Left01.gif" align="center"><b>单位会员管理</b></td>
  </tr>
        <% set rs=server.createobject("adodb.recordset")
              if not isempty(request("page")) then
	          pagecount=cint(request("page"))
              else
	          pagecount=1
              end if
		 	  sq2="select * from danwei order by idate desc"
			  rs.open sq2,conn,1,1


              rs.pagesize=20
              if pagecount>rs.pagecount or pagecount<=0 then
              pagecount=1
              end if
              rs.AbsolutePage=pagecount %>
        <tr> 
          <% if rs.pagecount=1 then %>
          <td ?height="19" colspan="2" valign="bottom" height="25" bgcolor="#FFFFFF"><font color="#000000" size="2">共有[<font color="#ff0000"><%=rs.recordcount%></font>]条<%=jylb%>记录 
            以下是[<font color="red">1～<%=rs.recordcount%></font>]条</font></td>
        </tr>
        <%else%>
        <tr> 
          <td colspan="2" valign="bottom" height="25" bgcolor="#FFFFFF"><font color="#000000"> 
            <% page_start=(pagecount-1)*rs.pagesize
            if pagecount=1 then page_start=1
		    page_end=rs.pagesize*pagecount
		    if pagecount*rs.pagesize=>rs.recordcount then page_end=rs.recordcount end if%>
            共有[<font color="#ff0000"><%=rs.recordcount%></font>]条<%=jylb%>记录 以下是[<font color="red"><%=page_start%>～<%=page_end%></font>]条</font></td>
        
        <tr> 
          <td colspan="9" valign="bottom"></td>
        <tr> 
          
    <td colspan="9" valign="bottom" bgcolor="#FFFFFF"> 
      <% response.write"<form name=go2to form method=Post action='dwzl.asp?key="+key+"&jylb="+jylb+"'>"
		     response.write "<font color='#000064'></font>"
		     if pagecount=1 then
			 response.write "<font color='#000064'>首页 上一页</font></font>&nbsp;"
			 else
             response.write "<a href=dwzl.asp?page=1&key="+key+"><font color='0000BE'>首页</font></font></a>&nbsp;"
	         response.write "<a href=dwzl.asp?page="+cstr(pagecount-1)+"&key="+key+"><font color='0000BE'>上一页</font></font></a>&nbsp;"
			 end if
             if rs.PageCount-pagecount<1 then
             response.write "<font color='#000064'>下一页 尾页</font></font>"
			 else
             response.write "<a href=dwzl.asp?page="+cstr(pagecount+1)+"&key="+key+"><font color='0000BE'>下一页</font></font></a>&nbsp;"
			 response.write "<a href=dwzl.asp?page="+cstr(rs.PageCount)+"&key="+key+"><font color='0000BE'>尾页</font></font></a>"
             end if
			response.write "<font color='000064'>&nbsp;页次:<font color=blue>"&pagecount&"</font>/"&rs.pagecount&"页</font></font>"
			response.write "<font color='000064'> 转到第<input type='text' name='page' size=2 maxLength=3 style='font-size: 9pt; color:#00006A; position: relative; height: 18' value="&PageCount&">页</font></font>&nbsp;"
			response.write "<input class=button type='button' value='确 定' onclick=check() style='font-family: 宋体; font-size: 9pt; color: #000073; position: relative; height: 19'>"
			response.write "</form>"
			%> 
			
    </td>
          <%end if%>
        </tr>
</table>
<br>
<table border="1" cellspacing="0" cellpadding="0" bordercolor="#808080" bordercolordark="#FFFFFF" width="96%" align="center">
  <tr bgcolor="#E7E7E7" align="center"> 
    <td height="32" width="6%">编号</td>
    <td width="17%">单位名称</td>
    <td width="8%">帐号</td>
    <td width="8%">密码</td>
    <td width="9%">级别</td>
    <td width="7%">加入日期</td>
    <td width="10%">操作</td>
  </tr>
  <% do while not rs.eof %>
  <form name="form1" method="post" action="dwzl.asp">
  <tr> 
      <td width="6%" height="30"><a class="sbtxt" href="gzl.asp?uid=<%=rs("id")%>" target="_blank" ><font color="#0080FF"><%=rs("id")%></font></a></td>
      <td width="17%" height="30"><a class="sbtxt" href="gzl.asp?uid=<%=rs("id")%>" target="_blank" ><font color="#0080FF"><%=rs("company")%></font></a></td>
      <td width="8%" align="center"><font color="#5786944"> <%=rs("loginmc")%></font></td>
      <td width="8%" align="center"> <%=rs("loginmm")%></td>
      <td align="center" width="9%"><a class="sbtxt" href="dwzl.asp?xgs=<%=rs("id")%>&zgs=<%=rs("zhige")%>"> 
        <%if rs("zhige")<>"末开通" then%>
        <font color="#ff0000"><u><%=rs("zhige")%></u></font> 
        <%else%>
        <%=rs("zhige")%> 
        <%end if%>
        </a></td>
      <td align="center" width="7%"> <%=rs("idate")%></td>
      <td width="10%" align="center"><a class="sbtxt" href="dwzl.asp?dels=<%=rs("id")%>&delcompany=<%=rs("company")%>" >删除</a>
        <input type="hidden" name="id" value="<%=rs("id")%>"></td>
  </tr>
</form>
  <% c=c+1
     rs.movenext
     if c>=20 then exit do
     loop
     rs.close
     set rs=nothing %>
</table>
                                  
  
<br>
<table border="0" cellpadding="2" cellspacing="1" width="96%" align="center" bgcolor="#6699ff">
  <tr> 
    <td colspan="3" align="center" height="25" background="Images/Left01.gif"><b>单位会员搜索</b></td>
  </tr>
  <form name="qyss" action="qyss.asp" method="post" target="_blank">
    <tr> 
      <td align="center" height="25" bgcolor="#FFFFFF"> 企业名称 
        <input type="text" name="gjc_" size="11" maxlength="30">
        <input type="button" value="搜 索" style="position: relative; color: #000000; font-family: 宋体; font-size: 10pt; height: 21; width: 50" onClick="qy_check()" name="button">
      </td>
    </tr>
  </form>
</table>
</body>

</html>