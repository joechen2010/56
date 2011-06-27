
<!--#include file = "../Inc/Syscode.asp"-->
<%

uid_=request("uid")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from geren where id="&uid_
rs.open sql,conn,1,1
if rs.eof then
response.write "发生错误！"
response.end
end if
%>
<html>
<head>
<title><%=rs("iname")%>个人资料</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="images/Admin.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table border="0" cellpadding="2" cellspacing="1" width="95%" align="center" bgcolor="#6699ff">
  <tr align="center"> 
    <td bgcolor="#FFFFFF" valign="bottom" height="18" colspan="6"><b><%=rs("iname")%>详细资料</b><b>>>></b> 
    </td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" valign="bottom" height="18" width="171"><font  color="#000099">姓名:</font></td>
    <td bgcolor="#FFFFFF" valign="bottom" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" valign="bottom" height="18" width="215"><%=rs("iname")%></td>
    <td bgcolor="#FFFFFF" valign="bottom" height="18" width="121" align="right"><font  color="#000099">性别:</font></td>
    <td bgcolor="#FFFFFF" valign="bottom" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" valign="bottom" height="18" width="238"><%=rs("sex")%></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" height="18" width="171"> 
      <p align="right"><font  color="#000099">出生年月:</font></p>
    </td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("bday")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">身份证号码:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"></td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("code")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">民族:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("mzhu")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">婚姻状况:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("marry")%></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" height="18" width="171"> 
      <p align="right"><font  color="#000099">户籍所在地:</font></p>
    </td>
    <td bgcolor="#FFFFFF" height="18" width="22">　</td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("huji")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">现住地:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("szd")%></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" height="18" width="171"> 
      <p align="right"><font  color="#000099">最高的教育程度:</font></p>
    </td>
    <td bgcolor="#FFFFFF" height="18" width="22">　</td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("edu")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">专业:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("zye")%></td>
  </tr>
  <tr> 
    <td height="18" bgcolor="#FFFFFF" width="171"> 
      <p align="right"><font  color="#000099">毕业院校:</font></p>
    </td>
    <td height="18" bgcolor="#FFFFFF" width="22">　</td>
    <td height="18" bgcolor="#FFFFFF" width="215"><%=rs("school")%></td>
    <td height="18" bgcolor="#FFFFFF" width="121" align="right"><font  color="#000099">职业:</font></td>
    <td height="18" bgcolor="#FFFFFF" width="23"> </td>
    <td height="18" bgcolor="#FFFFFF" width="238"><%=rs("work")%></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" height="18" width="171"> 
      <p align="right"><font  color="#000099">政治面貌:</font></p>
    </td>
    <td bgcolor="#FFFFFF" height="18" width="22">　</td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("zzmm")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font s color="#000099">现有职称:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("zchen")%></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" height="18" width="171"><font color="#000099"></font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font color="#000099"></font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"></td>
    <td bgcolor="#FFFFFF" height="18" width="238"> </td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">联系电话:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("phone")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font " color="#000099">手机/拷机:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("callnum")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">电子邮件:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("email")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font color="#000099">联系地址:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("address")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" colspan="6"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr> 
          <td width="100%" height="5">..............................................................................................................</td>
        </tr>
        <tr> 
          <td width="100%" height="10"></td>
        </tr>
        <tr> 
          <td width="100%" height="20"><b>求职信息>>></b></td>
        </tr>
      </table>
    </td>
  </tr>
  <%
uname_=rs("uname")
iname_=rs("iname")
 rs.close
set rs=nothing
 Set rs = Server.CreateObject("ADODB.Recordset")
sql2="select * from qiuzhi where uname='"&uname_&"' and iname='"&iname_&"'"
rs.open sql2,conn,1,1
if rs.eof then
response.write "<tr><td>暂无求职信息！</td></tr></table>"
response.end
end if
%>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font color="#000099">寻求职位:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("biaoti")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"> </td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"> </td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">外语特长:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("waiyu")%> 等级：<%=rs("dengji")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">计算机能力</font>:</td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("computer")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font color="#000099">主要技能:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" colspan="4"><%=rs("jineng")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">技能经验:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215">至今相关经验共有<%=rs("gznum")%> 年</td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"> </td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"> </td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">工作简历:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" colspan="4"><%=rs("gzjl")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">曾任职的最大企业规模:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("qyrs")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">所任职务:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("zhiwu")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">可胜任工作:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" colspan="4"><%=rs("job")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">希望工作地点:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("gzdd")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">薪金要求:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("yuex")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">其他要求:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" colspan="4"><%=rs("qtyq")%></td>
  </tr>
  <tr> 
    <td width="171" bgcolor="#FFFFFF"></td>
    <td width="22" bgcolor="#FFFFFF"></td>
    <td width="215" bgcolor="#FFFFFF"></td>
    <td width="121" bgcolor="#FFFFFF"></td>
    <td width="23" bgcolor="#FFFFFF"></td>
    <td width="238" bgcolor="#FFFFFF"></td>
  </tr>
</table>
                                          </body>
</html>
<%
rs.close
set rs=nothing
%>