<!--#include file = "../Inc/Syscode.asp"-->
<%

%>
<% if request("dels")<>"" then
   conn.Execute("delete * from danwei where id="&request("dels"))
'ɾ��������Ϣ
   conn.Execute("delete * from zhaopin where sqmc='"&request("delcompany")&"'")
'   conn.Execute("delete * from ypsj where uname='"&request("delcompany")&"'")
 '  conn.Execute("delete * from peixun where sqmc='"&request("delcompany")&"'")
   'conn.Execute("delete * from dwsq where dwmc='"&request("delcompany")&"'")
'ɾ������
   end if 
   
   if request("action")="����" then
   conn.Execute("update danwei set area='"&request("area")&"' where id="&request("id"))
    conn.Execute("update danwei set enddate='"&request("enddate")&"' where id="&request("id"))

   end if   
%>
<% if request("xgs")<>"" then
   if request("zgs")="ĩ��ͨ"then
   conn.Execute("update danwei set zhige='��ͨ',tdate='"&date()&"' where id="&request("xgs"))
   end if 
   if request("zgs")="��ͨ"then
   conn.Execute("update danwei set zhige='ĩ��ͨ',tdate='"&date()&"' where id="&request("xgs"))
   end if 
   end if
%>
<%uname=session("admin")%>

<html>
<head>
<title>��̨����</title>
<meta content="text/html; charset=gb2312" http-equiv="Content-Type">
<LINK href="images/Admin.css" rel=stylesheet>
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

<SCRIPT LANGUAGE="JavaScript">
function qy_check()
{
    if (qyss.gjc_.value=="") 
		alert("��������ҵ���ƣ�");
   	else
		qyss.submit();
}
//-->
</SCRIPT>
<body leftmargin="0" topmargin="0" bgcolor="#e1eefd">
<table border="0" cellpadding="2" cellspacing="1" width="96%" bordercolordark="#FFFFFF" bgcolor="#6699ff" align="center">
  <tr>
          
    <td colspan="2" height="27" background="Images/Left01.gif" align="center"><b>��λ��Ա����</b></td>
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
          <td ?height="19" colspan="2" valign="bottom" height="25" bgcolor="#FFFFFF"><font color="#000000" size="2">����[<font color="#ff0000"><%=rs.recordcount%></font>]��<%=jylb%>��¼ 
            ������[<font color="red">1��<%=rs.recordcount%></font>]��</font></td>
        </tr>
        <%else%>
        <tr> 
          <td colspan="2" valign="bottom" height="25" bgcolor="#FFFFFF"><font color="#000000"> 
            <% page_start=(pagecount-1)*rs.pagesize
            if pagecount=1 then page_start=1
		    page_end=rs.pagesize*pagecount
		    if pagecount*rs.pagesize=>rs.recordcount then page_end=rs.recordcount end if%>
            ����[<font color="#ff0000"><%=rs.recordcount%></font>]��<%=jylb%>��¼ ������[<font color="red"><%=page_start%>��<%=page_end%></font>]��</font></td>
        
        <tr> 
          <td colspan="9" valign="bottom"></td>
        <tr> 
          
    <td colspan="9" valign="bottom" bgcolor="#FFFFFF"> 
      <% response.write"<form name=go2to form method=Post action='dwzl.asp?key="+key+"&jylb="+jylb+"'>"
		     response.write "<font color='#000064'></font>"
		     if pagecount=1 then
			 response.write "<font color='#000064'>��ҳ ��һҳ</font></font>&nbsp;"
			 else
             response.write "<a href=dwzl.asp?page=1&key="+key+"><font color='0000BE'>��ҳ</font></font></a>&nbsp;"
	         response.write "<a href=dwzl.asp?page="+cstr(pagecount-1)+"&key="+key+"><font color='0000BE'>��һҳ</font></font></a>&nbsp;"
			 end if
             if rs.PageCount-pagecount<1 then
             response.write "<font color='#000064'>��һҳ βҳ</font></font>"
			 else
             response.write "<a href=dwzl.asp?page="+cstr(pagecount+1)+"&key="+key+"><font color='0000BE'>��һҳ</font></font></a>&nbsp;"
			 response.write "<a href=dwzl.asp?page="+cstr(rs.PageCount)+"&key="+key+"><font color='0000BE'>βҳ</font></font></a>"
             end if
			response.write "<font color='000064'>&nbsp;ҳ��:<font color=blue>"&pagecount&"</font>/"&rs.pagecount&"ҳ</font></font>"
			response.write "<font color='000064'> ת����<input type='text' name='page' size=2 maxLength=3 style='font-size: 9pt; color:#00006A; position: relative; height: 18' value="&PageCount&">ҳ</font></font>&nbsp;"
			response.write "<input class=button type='button' value='ȷ ��' onclick=check() style='font-family: ����; font-size: 9pt; color: #000073; position: relative; height: 19'>"
			response.write "</form>"
			%> 
			
    </td>
          <%end if%>
        </tr>
</table>
<br>
<table border="1" cellspacing="0" cellpadding="0" bordercolor="#808080" bordercolordark="#FFFFFF" width="96%" align="center">
  <tr bgcolor="#E7E7E7" align="center"> 
    <td height="32" width="6%">���</td>
    <td width="17%">��λ����</td>
    <td width="8%">�ʺ�</td>
    <td width="8%">����</td>
    <td width="9%">����</td>
    <td width="7%">��������</td>
    <td width="10%">����</td>
  </tr>
  <% do while not rs.eof %>
  <form name="form1" method="post" action="dwzl.asp">
  <tr> 
      <td width="6%" height="30"><a class="sbtxt" href="gzl.asp?uid=<%=rs("id")%>" target="_blank" ><font color="#0080FF"><%=rs("id")%></font></a></td>
      <td width="17%" height="30"><a class="sbtxt" href="gzl.asp?uid=<%=rs("id")%>" target="_blank" ><font color="#0080FF"><%=rs("company")%></font></a></td>
      <td width="8%" align="center"><font color="#5786944"> <%=rs("loginmc")%></font></td>
      <td width="8%" align="center"> <%=rs("loginmm")%></td>
      <td align="center" width="9%"><a class="sbtxt" href="dwzl.asp?xgs=<%=rs("id")%>&zgs=<%=rs("zhige")%>"> 
        <%if rs("zhige")<>"ĩ��ͨ" then%>
        <font color="#ff0000"><u><%=rs("zhige")%></u></font> 
        <%else%>
        <%=rs("zhige")%> 
        <%end if%>
        </a></td>
      <td align="center" width="7%"> <%=rs("idate")%></td>
      <td width="10%" align="center"><a class="sbtxt" href="dwzl.asp?dels=<%=rs("id")%>&delcompany=<%=rs("company")%>" >ɾ��</a>
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
    <td colspan="3" align="center" height="25" background="Images/Left01.gif"><b>��λ��Ա����</b></td>
  </tr>
  <form name="qyss" action="qyss.asp" method="post" target="_blank">
    <tr> 
      <td align="center" height="25" bgcolor="#FFFFFF"> ��ҵ���� 
        <input type="text" name="gjc_" size="11" maxlength="30">
        <input type="button" value="�� ��" style="position: relative; color: #000000; font-family: ����; font-size: 10pt; height: 21; width: 50" onClick="qy_check()" name="button">
      </td>
    </tr>
  </form>
</table>
</body>

</html>