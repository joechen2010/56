<!--#include file="../inc/syscode.asp" -->
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="../inc/Admin_style.css" rel=stylesheet>
<META content="MSHTML 6.00.2462.0" name=GENERATOR>
<TITLE>网站基本设置</TITLE>
</HEAD>
<BODY bgColor=#e1eefd leftMargin=13 Dnpmargin="10" TOPMARGIN="10">
<%
dim rs,sql
if request("action")="edit" then
call edit()
end if%> <%
set rs=server.CreateObject ("ADODB.RecordSet")
sql="select * from system"
rs.Open sql,conn,1,1
%> <form method="POST" action="system.asp?action=edit"> <div align="center"> <center> 
<table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" width="90%" BGCOLOR="#6699ff"><tr ALIGN="CENTER"><td COLSPAN="2" height="27" BACKGROUND="images/Left01.gif"><B>联 
系 我 们</B></td></tr> <%do while not rs.eof%> <tr> <td width="20%" height="25" align="right" BGCOLOR="#FFFFFF">公司名称：</td><td width="46%" height="25" BGCOLOR="#FFFFFF"> 
&nbsp;<input type="text" name="company" size="60" value="<%=rs("company")%>"> 
<input type="hidden" name="id" value="<%=rs("id")%>"> </td></tr> <tr> <td width="20%" height="25" align="right" BGCOLOR="#FFFFFF">公司地址：</td><td width="46%" height="25" BGCOLOR="#FFFFFF"> 
&nbsp;<input type="text" name="address" size="60" value="<%=rs("address")%>"></td></tr> 
<tr> <td width="20%" height="25" align="right" BGCOLOR="#FFFFFF">电话：</td><td width="46%" height="25" BGCOLOR="#FFFFFF"> 
&nbsp;<input type="text" name="tel" size="60" value="<%=rs("tel")%>"></td></tr> 
<tr> <td width="20%" height="25" align="right" BGCOLOR="#FFFFFF">传真：</td><td width="46%" height="25" BGCOLOR="#FFFFFF"> 
&nbsp;<input type="text" name="fax" size="60" value="<%=rs("fax")%>"></td></tr> 
<tr> <td width="20%" height="25" align="right" BGCOLOR="#FFFFFF">电子邮箱：</td><td width="46%" height="25" BGCOLOR="#FFFFFF"> 
&nbsp;<input type="text" name="email" size="60" value="<%=rs("email")%>"></td></tr><tr><td width="20%" height="25" align="right" BGCOLOR="#FFFFFF">企业邮局：</td><td width="46%" height="25" BGCOLOR="#FFFFFF">&nbsp;<input type="text" name="MailUrl" size="60" value="<%=rs("MailUrl")%>"></td></tr> 
<tr> <td width="20%" height="25" align="right" BGCOLOR="#FFFFFF">网址：</td><td width="46%" height="25" BGCOLOR="#FFFFFF"> 
&nbsp;<input type="text" name="website" size="60" value="<%=rs("Website")%>"></td></tr> 
<tr><td COLSPAN="2" height="27" BACKGROUND="images/Left01.gif" ALIGN="CENTER"><B>公 
司 简 介</B></td></tr> <tr> <td width="20%" height="60" align="right" BGCOLOR="#FFFFFF">公司简介：</td><td width="46%" height="60" BGCOLOR="#FFFFFF">&nbsp;<textarea rows="15" name="Brief" cols="60"><%=dvHTMLCode(rs("Brief"))%></textarea></td></tr> 
<tr> <td width="66%" height="50" colspan="2" align="center" BGCOLOR="#FFFFFF"><input type="submit" value="提交" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<input type="reset" value="重置" name="B2"></td></tr> <%rs.movenext
loop%> </table></center></div></form>

</body>

</html>

<%
sub edit()
set rs=server.CreateObject ("ADODB.RecordSet")
sql="select * from system where id="&request("id")
rs.Open sql,conn,1,3
rs("company")=request("company")
rs("address")=request("address")
rs("MailUrl")=request("MailUrl")
rs("tel")=request("tel")
rs("fax")=request("fax")
rs("email")=request("email")
rs("Website")=request("website")
rs("Brief")=dvHTMLEncode(request("brief"))
rs.update
rs.close
conn.close
set rs=nothing
set conn=nothing
response.redirect "system.asp"

end sub
%>