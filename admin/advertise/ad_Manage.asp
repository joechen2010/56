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
		response.Redirect "ad_manage.asp?page="&page&""
     		else
			call deleteannounce(clng(idlist))
			response.Redirect "ad_manage.asp?page="&page&""
     		end if
  	end if 
	end if
sub deleteannounce(id)
    dim rs_del,sql_del
    set rs_del=server.createobject("adodb.recordset")
    sql_del="delete from advertise where id="&cstr(id)
    conn.execute sql_del
End sub	
'delete -----------
'if  Request.QueryString("del")<>"" then
   'delsql="delete from goods where id="&Request.QueryString("del")
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
      <td align="center" bgcolor="#6699ff" height="30"><font color="#FFFFFF"> ������</font></td>
    </tr>
  </table>
   <br>
   <form method="post" name="myform" action="ad_manage.asp">
   <% 
	  set rs=Server.Createobject("ADODB.Recordset")
    sql="select * from advertise order by id desc,addtime desc"
      rs.open sql,conn,1,3
	  if rs.recordcount=0 then
	  Response.Write "δ������"
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
    <td width="79" height="25" bgcolor="#e1eefd" align="center">ID</td>
    <td width="176" height="25" bgcolor="#e1eefd" align="center">���ڰ��</td>
	<td width="212" height="25" bgcolor="#e1eefd" align="center">��˾����</td>
    <td width="152" height="25" bgcolor="#e1eefd" align="center">������ַ</td>
    <td width="140" height="25" bgcolor="#e1eefd" align="center">��������</td>
	<td width="81"  bgcolor="#e1eefd" align="center">�޸�</td>
	<td width="108"  bgcolor="#e1eefd" align="center">ɾ��</td>
    </tr>
	 <% for ipage=1 to rs.pagesize%>
    <tr bgcolor="#F6F6F6"  onMouseOut="mOut(this,'#F6F6F6');" onMouseOver="mOvr(this,'#EBFFBB');"> 
    <td align="center"><%=rs("id")%></td>
	<td>
	<%
	adclass=rs("adclass")
	select case adclass
	  case "1"
	  response.Write "�һ��ҳ�"
	  case "2"
	  response.Write "����ר��"
	  case "3"
	  response.Write "����ר��"
	  case "4"
	  response.Write "��������"
	  case "5"
	  response.Write "�������"	  
	  case "6"
	  response.Write "�ⷿ��Ϣ"
	  case "7"
	  response.Write "��Ƹ��ְ"
	  case "8"
	  response.Write "������Ϣ"	  	  
	end select
	
	%>
	</td>
    <td align="Left"><%=rs("webname")%></td>
    <td><a href="<%=rs("weburl")%>" target="_blank"><%=rs("weburl")%></a></td>
    <td><%=rs("addtime")%></td>
	<td align="center"><a href="add.asp?manage=modify&id=<%=rs("id")%>&page=<%=page%>">�޸�</a></td>
	<td align="center"><input type='checkbox' name='selAnnounce' value='<%=cstr(rs("ID"))%>'>
	  <!--<a href="?del=<%'=rs("id")%>"  onClick="return cform()">ɾ��</a>--></td>
    </tr>
	<%      rs.movenext
            if rs.eof then exit for
            next%>
  </table> 
  <table width="96%" height="21" border="0" cellpadding="0" cellspacing="1" background="images/bg-xg.gif" bgcolor="#EFEFEF">
    <tr> 
    <td width="37%" align="center" bgcolor="#e7e7e7"> 
	 <input type="hidden" name="action" value="del" >
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