<!--#include file = "../Inc/Syscode.asp"-->
<%

%>
<%
if request("action")="join" then
      conn.execute("update geren set rcjy=true Where id="&request("id")&" ") 
end if
if request("action")="cancle" then
      conn.execute("update geren set rcjy=false Where id="&request("id")&" ") 
end if
%>
<% if request("dels")<>"" then
   conn.Execute("delete * from geren where id="&request("dels"))
   conn.execute("delete * from qiuzhi where uname='"&request("uname")&"'")
   end if %>
<%uname=session("admin")%>
<html>
<head>
<title>��̨����</title>
<meta content="text/html; charset=gb2312" http-equiv="Content-Type">
<meta content="zh-cn" http-equiv="Content-Language">
<LINK href="images/Admin.css" rel=stylesheet>
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta content="FrontPage.Editor.Document" name="ProgId">
</head>
<SCRIPT LANGUAGE="JavaScript">
 <!--//
function check()
{
   if (isNaN(go2to.page.value))
		alert("����ȷ��дת��ҳ����");
   else if (go2to.page.value=="")
	     {
		alert("������ת��ҳ����");
		 }
   else
		go2to.submit();
}
//-->
</SCRIPT>
<body>
  
<table border="1" cellpadding="1" cellspacing="0" width="96%" bordercolor="#808080" bordercolordark="#FFFFFF" align="center">
  <tr valign="middle" bgcolor="#F5F5F5" align="center"> 
      
    <td height="30" width="10%">���˼�������</td>
  </tr>
</table>
  
<table border="0" cellpadding="2" cellspacing="1" width="96%" height="0" bordercolor="#316594" bordercolordark="#FFFFFF" align="center">
  <% set rs=server.createobject("adodb.recordset")
              if not isempty(request("page")) then
	          pagecount=cint(request("page"))
              else
	          pagecount=1
              end if
			  sq2="select * from geren"
			  if request("action")="serch" then
			 		if request("type")="1" then
		 	  sq2=sq2& " where uname like '%"&request("key")&"%'"
			  else 
			  sq2=sq2& " where iname like '%"&request("key")&"%'"
			  end if 
			  end if
			  
			 sq2=sq2&" order by id desc" 
			  
			  rs.open sq2,conn,1,3

                	if rs.eof and rs.bof then
       		response.write "<p>&nbsp;</p><p>&nbsp;</p><p align='center'> �� û �� �� �� �� Ϣ </p>"
   	else
              rs.pagesize=20
              if pagecount>rs.pagecount or pagecount<=0 then
              pagecount=1
              end if
              rs.AbsolutePage=pagecount %>
  <tr>                         
                                <% if rs.pagecount=1 then %>
                                <td colspan="2" valign="bottom" height="10">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" valign="bottom" height="11"><font color="#000000" size="2">����[<font color="#ff0000"><%=rs.recordcount%></font>]��<%=jylb%>��¼ 
    ������[<font color="red">1��<%=rs.recordcount%></font>]��</font></td>
  </tr>
                              <%else%>
                              <tr> 
                                
    <td height="12" colspan="2" valign="bottom"><form name="form1" method="post" action=grzl.asp?action=serch>
        <div align="center">����ؼ���
          <input name="key" type="text" id="key">
          ����
          <select name="type" id="type">
              <option value="1">�û���</option>
              <option value="2">����</option>
                      </select>
          ����
          <input type="submit" name="Submit" value="��ѯ">
        </div>
      </form></td>
                              <tr>
                                <td colspan="2" valign="bottom" height="12"><font color="#000000">
                                  <% page_start=(pagecount-1)*rs.pagesize
            if pagecount=1 then page_start=1
		    page_end=rs.pagesize*pagecount
		    if pagecount*rs.pagesize=>rs.recordcount then page_end=rs.recordcount end if%>
����[<font color="#ff0000"><%=rs.recordcount%></font>]��<%=jylb%>��¼ ������[<font color="red"><%=page_start%>��<%=page_end%></font>]��</font></td>
                              <tr> 
                                <td colspan="9" valign="bottom"></td>
                              <tr> 
                                <td colspan="9" valign="bottom">
                                  <% response.write"<form name=go2to form method=Post action='grzl.asp?key="+key+"&jylb="+jylb+"'>"
		     response.write "<font color='#000064'></font>"
		     if pagecount=1 then
			 response.write "<font color='#000064'>��ҳ ��һҳ</font></font>&nbsp;"
			 else
             response.write "<a href=grzl.asp?page=1&key="+key+"><font color='0000BE'>��ҳ</font></font></a>&nbsp;"
	         response.write "<a href=grzl.asp?page="+cstr(pagecount-1)+"&key="+key+"><font color='0000BE'>��һҳ</font></font></a>&nbsp;"
			 end if
             if rs.PageCount-pagecount<1 then
             response.write "<font color='#000064'>��һҳ βҳ</font></font>"
			 else
             response.write "<a href=grzl.asp?page="+cstr(pagecount+1)+"&key="+key+"><font color='0000BE'>��һҳ</font></font></a>&nbsp;"
			 response.write "<a href=grzl.asp?page="+cstr(rs.PageCount)+"&key="+key+"><font color='0000BE'>βҳ</font></font></a>"
             end if
			response.write "<font color='000064'>&nbsp;ҳ��:<font color=blue>"&pagecount&"</font>/"&rs.pagecount&"ҳ</font></font>"
			response.write "<font color='000064'> ת����<input type='text' name='page' size=2 maxLength=3 style='font-size: 9pt; color:#00006A; position: relative; height: 18' value="&PageCount&">ҳ</font></font>&nbsp;"
			response.write "<input class=button type='button' value='ȷ ��' onclick=check() style='font-family: ����; font-size: 9pt; color: #000073; position: relative; height: 19'>" %>                                </td>
                                <%end if%>
                              </tr>
                              <tr> 
                                <td height="25" valign="center"  align="left"> 
                                  
                    <table border="1" cellspacing="0" cellpadding="0" bordercolor="#808080" bordercolordark="#FFFFFF" width="100%">
        <tr bgcolor="#E7E7E7" align="center"> 
          <td width="8%" height="30">���</td>
          <td width="8%">�û�����</td>
          <td width="17%">����</td>
          <td width="14%">����</td>
          <td width="7%">�Ա�</td>
          <td width="15%">�绰</td>
          <td width="11%">ע��ʱ��</td>
          <td width="9%">����</td>
        </tr>
        <% do while not rs.eof %>
        <tr> 
          <td height="30" align="center"><a class="sbtxt" href="zl.asp?uid=<%=rs("id")%>" target="_blank" ><font color="#0080FF"><%=rs("id")%></font></a></td>
          <td height="30" align="center"><a class="sbtxt" href="zl.asp?uid=<%=rs("id")%>" target="_blank" ><font color="#0080FF"><%=rs("uname")%></font></a></td>
          <td align="center"><font color="#5786944"> <%=rs("pwd")%></font></td>
          <td align="center">&nbsp;<%=rs("iname")%></td>
          <td align="center">&nbsp;<%=rs("sex")%></td>
          <td align="center">&nbsp;<%=rs("phone")%></td>
          <td align="center"><%=rs("idate")%></td>
          <td align="center"> <a class="sbtxt" href="grzl.asp?dels=<%=rs("id")%>&uname=<%=rs("uname")%>" >ɾ��</a></td>
        </tr>
        <% c=c+1
     rs.movenext
     if c>=50 then exit do
     loop
     rs.close
     set rs=nothing %>
	 <%end if%>
      </table>                                </td>
                              </tr>
                            </table>
</body>

</html>