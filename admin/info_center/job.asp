<!--#include file = "../Inc/Syscode.asp"-->
<% if request("action")="del" then
   page=request("page")
  	if not isempty(request("selAnnounce")) then
     		idlist=request("selAnnounce")
     if instr(idlist,",")>0 then
			dim idarr
			idArr=split(idlist)
			dim id
		for i = 0 to ubound(idarr)
	       		id=clng(idarr(i))
	       		call deleteannounce(id)
		next
		response.Redirect "job.asp?page="&page&""
     		else
			call deleteannounce(clng(idlist))
			response.Redirect "job.asp?page="&page&""
     		end if
  	end if 
	end if
sub deleteannounce(id)
    dim rs_del,sql_del
    set rs_del=server.createobject("adodb.recordset")
    sql_del="delete from zp_info where infoid="&cstr(id)
    conn.execute sql_del
End sub	
if request("pass")="0" then
  conn.execute("update zp_info set pass=false where infoid="&request("id"))
  response.Redirect "?page="&request("page")&""
  elseif request("pass")="1" then
   conn.execute("update zp_info set pass=true where infoid="&request("id"))
  response.Redirect "?page="&request("page")&"" 
end if
'delete -----------
'if  Request.QueryString("del")<>"" then
   'delsql="delete from job where id="&Request.QueryString("del")
  ' conn.execute(delsql)
'end if
'delete end---------------
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../inc/Admin_Style.css">
<script language="javascript">
<!--

function mOvr(src,clrOver){
	if (!src.contains(event.fromElement)) {
		//src.style.cursor = 'hand';
		src.bgColor = clrOver;
	}
}
function mOut(src,clrIn)  {
	if (!src.contains(event.toElement)) {
		src.style.cursor = 'default';
		src.bgColor = clrIn;
	}
}
function mClk1(src,id) {
		window.location .href =src;
 }
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ�������¼��        \n���ȷ��ɾ����\n���ȡ�����أ�"))
     return true;
   else
     return false;
	 
}


function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled==false)
       e.checked = form.chkAll.checked;
    }
  }
function cform(){
 if(!confirm("Are you sure?"))
 return false;
}  
//-->
</script>
<script language="javascript">
function openwindow(url,width,height) { 	left1 = (screen.width-800)/2; 	top1 = (screen.height-350)/2; 	window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); }

</script>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#e1eefd">
<div align="center"> 
  <table width="96%" border="0" align="center" cellspacing="0" class="px14">
    <tr> 
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#6699ff" height="30"><font color="#FFFFFF"> ��Ƹ��ְ����</font></td>
    </tr>
  </table>
   <br>
   <form method="post" name="myform" action="job.asp">
   <% 
	  set rs=Server.Createobject("ADODB.Recordset")
    sql1="select * from zp_info a,qyml b where a.gsid=b.user order by infoid desc"
       rs.open sql1,conn,1,3
	  if rs.recordcount=0 then
	  Response.Write "������"
	  Response.end
	  else
          rs.pagesize=20
          if request("page")="" then
           page=0
           else
           page=clng(request("page"))
          end if
           if page<1 then page=1
           if page>rs.pagecount then page=rs.pagecount
          rs.AbsolutePage=page
		%>
  <table width="96%" border="0" cellspacing="1">
    <tr bgcolor="#CCCCCC"> 
    <td width="197" height="25" bgcolor="#e1eefd" align="center">��Ϣ����</td>
    <td width="123" height="25" bgcolor="#e1eefd" align="center">��Ƹְλ/��ְ��</td>
	<td width="110" height="25" bgcolor="#e1eefd" align="center">��Ƹ��λ</td>
    <td width="109" height="25" bgcolor="#e1eefd" align="center">����ʡ��</td>
	<td width="245" height="25" bgcolor="#e1eefd" align="center">����ʱ��</td>
	<td width="157" height="25" bgcolor="#e1eefd" align="center">���</td>
    <td width="167"  bgcolor="#e1eefd" align="center">ɾ��</td>
    </tr>
	 <% for ipage=1 to rs.pagesize%>
   <tr bgcolor="#F6F6F6"  onMouseOut="mOut(this,'#F6F6F6');" onMouseOver="mOvr(this,'#EBFFBB');"> 
      <td>&nbsp;<%=rs("infotype")%></td>
      <td width="123">
	  <a href="javascript:openwindow('../../job/detail.asp?infoid=<%=rs("InfoID")%>',500,400)"><%=rs("job")%></a></td>
      <td  width="110" align="center">
				  <%if len(rs("qymc"))>8 then%>
				  <%=left(rs("qymc"),6)%>...
				  <%else%>
				 <%=rs("qymc")%>
				  <%end if%>	  
	  </td>
      <td width="109" align="center"><%=rs("a.province")%>&nbsp;<%=rs("a.city")%></td>
	  <td align="center"><%=rs("addtime")%></td>
	<td width="157" align="center">
	  <%
	  if rs("a.pass")=true then
	    response.Write "<a href='?pass=0&page="&page&"&id="&rs("infoid")&"'>�����</a>"
	   else
	    response.write "<a href='?pass=1&page="&page&"&id="&rs("infoid")&"'>δ���</a>"	
	  end if
	  %>	
	</td>	  
      <td align="center"><input type='checkbox' name='selAnnounce' value='<%=cstr(rs("InfoID"))%>'>
	  <!--<a href="?del=<%'=rs("id")%>"  onClick="return cform()">ɾ��</a>--></td>
    </tr>
	<%rs.movenext
            if rs.eof then exit for
            next%>
  </table> 
  <table width="96%" height="21" border="0" cellpadding="0" cellspacing="1" background="images/bg-xg.gif" bgcolor="#EFEFEF">
    <tr> 
    <td width="37%" align="center" bgcolor="#e7e7e7"> 
	 <input type="hidden" name="action" value="del">
	 <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;"><label for="chkAll">ȫ��ѡ��</label>&nbsp;&nbsp;<input type="submit" name="sc" value="ɾ&nbsp;��&nbsp;��&nbsp;ѡ" onClick="return cform()"></td>
	 </tr>
	 <tr><td width="63%" align="right" bgcolor="#e1eefd">	ÿҳ��ʾ<font color="#FF0000"><%=rs.pagesize%></font>��&nbsp;&nbsp;	
	     ����<font color="#FF0000"><%=rs.recordcount%></font>����Ϣ	&nbsp;&nbsp;
			��ǰ<font color="#FF0000"><%=page%></font>/<%=rs.pagecount%>
              <% if page<>1 then %>
              <a href="?page=1">��ҳ</a>&nbsp;<a href=?page=<%=(page-1)%>>��һҳ</a> 
              <%end if%>
              <% if page<>rs.pagecount then %>
              <a href=?page=<%=(page+1)%>>��һҳ</a>&nbsp;<a href="?page=<%=rs.pagecount%>">���һҳ</a> 
              <%end if%>
			 <input type="text" name="page" value="<%=page%>" size="4">
			 <input type="submit" value="��&nbsp;ת">
        </td>
    </tr>
  </table> 
	  <% 	rs.close
			set rs=nothing
			End If %>	
  </form>
</div>
<br>
</body>
</html>
<% 
conn.close
set conn=nothing %>