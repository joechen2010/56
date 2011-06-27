<%@ codepage ="936" %>
<!--#include file="../../Inc/Conn.asp" -->
<LINK href="../style.css" rel=stylesheet type=text/css> 
<%
dim sql
dim rs
sql="select * from qyml where id="&request("id")
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 

      
<TABLE border=0 cellspacing=1 width="96%" style="border-collapse: collapse" cellpadding="6" align="center" bgcolor="#e8f4ff">
  <TBODY> 
  <tr bgcolor="#e8f4ff"> 
    <TD colspan=2 height=30 width="460"><b>修改注册会员信息</b></TD>
  </tr>
  <TR bgcolor="#FFFFFF"> 
    <TD align="center">　会员登录名：</TD>
    <TD width=312><%=rs("user")%></TD>
  </TR>
  <TR bgcolor="#FFFFFF"> 
    <TD align="center">　密 码：</TD>
    <TD width="312"><%=rs("pass")%> </TD>
  </TR>
  <TR bgcolor="#FFFFFF"> 
    <TD align="center">　密码提示问题</TD>
    <TD width="312"><%=rs("question")%></TD>
  </TR>
  <TR vAlign="middle" bgcolor="#FFFFFF"> 
    <TD align="center">　密码提示答案：</TD>
    <TD width="312"><%=rs("answer")%></TD>
  </TR>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">　您的姓名：</TD>
    <TD width="482"><%=rs("name")%> <%=rs("ch")%> </TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center" height="19">　您的职务：</TD>
    <TD width="312"><%=rs("zw")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">　部&nbsp;&nbsp;&nbsp; 门：</TD>
    <TD width="312"><%=rs("bm")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">　街道地址：</TD>
    <TD width="312"><%=rs("address")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">　邮&nbsp;&nbsp;&nbsp; 编：</TD>
    <TD width="312"><%=rs("post")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">　省级地区：</TD>
    <TD width="312"><%=rs("province")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">　城&nbsp;&nbsp;&nbsp; 市：</TD>
    <TD width="312"><%=rs("city")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">　电话：</TD>
    <TD width="312"> 
      <%response.write rs("phone1")&"-"&rs("phone2")&"-"&rs("phone3")%>
    </TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">　传&nbsp;&nbsp;&nbsp; 真：</TD>
    <TD width="312"> 
      <%response.write rs("fax1")&"-"&rs("fax2")&"-"&rs("fax3")%>
    </TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">　您的电子邮件：</TD>
    <TD width="312"><%=rs("email")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">　网&nbsp;&nbsp;&nbsp; 址：</TD>
    <TD width="312"><%=rs("web")%> </td>
  </tr>
  <tr bgcolor="#e8f4ff"> 
    <TD colspan=2 height=30 width="460"><b>修改公司信息</b></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">公司名称：</TD>
    <TD width="312"><%=rs("qymc")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">法人代表：</TD>
    <TD width="312"><%=rs("frdb")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">公司性质：</TD>
    <TD width="312" ><%=rs("qyxz")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">注册资金：</TD>
    <TD width="312"><%=rs("zczj")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">主要产品：</TD>
    <TD width="312"><%=rs("zycp")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">雇员人数：</TD>
    <TD width="312"><%=rs("ygrs")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">年营业额：</TD>
    <TD width="312"><%=rs("nyye")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">公司介绍：</TD>
    <TD width="312"><%=rs("qyjj")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD height="10" width="145" ></TD>
    <TD height="10" width="145" ></TD>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <TD height="20" colspan="2"><a href="javascript:window.close()">关闭</a></TD>
  </tr>
  </TBODY> 
</TABLE>
