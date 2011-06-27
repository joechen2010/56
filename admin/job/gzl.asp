
<!--#include file = "../Inc/Syscode.asp"-->
<%

uid_=request("uid")
set rs=server.createobject("adodb.recordset")
sql1="select * from danwei where id="&uid_
rs.open sql1,conn,1,1
if rs.eof then
response.write "发生错误！"
response.end
end if
%>
<html>
<head>
<title><%=rs("company")%>公司信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="images/Admin.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<tr>
  <td width="100%" valign="top" height="20">&nbsp;</td>
                                  </tr>
                                  <tr> 
                                    
  <td width="100%" valign="top" height="30">
    <div align="center">
      <table border="1" cellpadding="1" cellspacing="0" width="96%" bordercolor="#808080" bordercolordark="#FFFFFF">
        <tr valign="middle" bgcolor="#F5F5F5" align="center"> 
                            <td height="32" width="10%"><font size="2" color="#000000"><font color="#0000FF">&nbsp;<%=rs("company")%></font></font></td>
                          </tr>
                        </table><br>
    </div>
  </td>
</tr>
<tr>
  <td width="100%" valign="bottom" align="center"></td>
                                  </tr>
                                  <tr> 
                                    
  <td width="100%" align="center"> 
    <table width="96%" border="0" cellspacing="1" cellpadding="2" align="center" bordercolordark="#FFFFFF">
      <tr> 
        <td width="163" bgcolor="#E7E7E7">企业性质：</td>
        <td width="623">&nbsp;<%=rs("qyxz")%></td>
      </tr>
      <tr> 
        <td width="163" bgcolor="#E7E7E7">经营模式：</td>
        <td width="623">&nbsp;<%=rs("jyms")%></td>
      </tr>
      <tr> 
        <td width="163" bgcolor="#E7E7E7">成立时间：</td>
        <td width="623">&nbsp;<%=rs("clsj")%></td>
      </tr>
      <tr> 
        <td width="163" bgcolor="#E7E7E7">联系人：</td>
        <td width="623">&nbsp;<%=rs("lxr")%></td>
      </tr>
      <tr> 
        <td width="163" bgcolor="#E7E7E7">员工人数：</td>
        <td width="623">&nbsp;<%=rs("ygrs")%></td>
      </tr>
      &nbsp; 
      <%if rs("leibie")<>"商家" then%>
      &nbsp; 
      <%end if%>
      <tr valign="top"> 
        <td height="19" colspan="2">&nbsp;</td>
      </tr>
      <tr> 
        <td width="163" bgcolor="#E7E7E7">联系电话：</td>
        <td width="623">&nbsp;<%=rs("tel")%></td>
      </tr>
      <tr> 
        <td width="163" bgcolor="#E7E7E7">联系传真：</td>
        <td width="623">&nbsp;<%=rs("fax")%></td>
      </tr>
      <tr> 
        <td width="163" bgcolor="#E7E7E7">公司地址：</td>
        <td width="623">&nbsp;<%=rs("address")%>(&nbsp;<%=rs("diqu")%>)</td>
      </tr>
      <tr> 
        <td width="163" bgcolor="#E7E7E7">邮政编码：</td>
        <td width="623">&nbsp;<%=rs("pc")%></td>
      </tr>
      <tr> 
        <td width="163" bgcolor="#E7E7E7">公司网址：</td>
        <td width="623">&nbsp;<%=rs("http")%></td>
      </tr>
      <tr> 
        <td width="163" bgcolor="#E7E7E7">电子信箱：</td>
        <td width="623">&nbsp;<%=rs("email")%></td>
      </tr>
      <tr valign="top"> 
        <td height="20" colspan="2">&nbsp;</td>
      </tr>
      <tr> 
        <td width="163" bgcolor="#E7E7E7">公司简介：</td>
        <td width="623">&nbsp;</td>
      </tr>
      <tr> 
        <td colspan="2">&nbsp;<%=rs("about")%></td>
      </tr>
    </table>
  
<%
rs.close
set rs=nothing
%>
