<!--#include file = "../Inc/Syscode.asp"-->
<%
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<link rel="stylesheet" type="text/css" href="../style.css"> 
<link rel="stylesheet" type="text/css" href="../inc/Admin_Style.css">

<body bgcolor="#FFFFFF" marginheight=0 marginwidth=0 leftmargin=0>
<script language="javascript">
function show_set(mylink)
{
window.open(mylink,"new",'top=110,left=280,width=240,height=60,scrollbars=no')
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
</script> 

<%
dim keyword,strField,sfilename
strField=trim(request("Field"))
keyword=trim(request("keyword"))

if keyword<>"" then 
	keyword=ReplaceBadChar(keyword)
end if

if keyword="" then
sfilename="default.asp"
else
sfilename="default.asp?Field=" & strField & "&Keyword=" & keyword & ""
end if

dim totalPut 
dim CurrentPage
dim TotalPages
dim i,j
const MaxPerPage=12
if not isempty(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if
dim search
search=request("search")
if not isempty(request("selAnnounce")) then
idlist=request("selAnnounce")
if instr(idlist,",")>0 then
dim idarr
idArr=split(idlist)
dim id
for i = 0 to ubound(idarr)
id=clng(idarr(i))
if request("submit")="ɾ��" then
    call deleteannounce(id)
	response.write "<script>alert('�����ɹ���');history.back();</script>"
elseif request("submit")="���" then
    call pass(id)
	response.write "<script>alert('�����ɹ���');history.back();</script>"
else
    response.write "����"
	response.end
end if
next
else
    if request("submit")="ɾ��" then   
        call deleteannounce(clng(idlist))
		response.write "<script>alert('�����ɹ���');history.back();</script>"
	elseif request("submit")="���" then
	    call pass(clng(idlist))
		response.write "<script>alert('�����ɹ���');history.back();</script>"
	else
	    response.write "����"
	    response.end
	end if
end if
end if 
dim rs
dim sql

set rs=server.createobject("adodb.recordset") 
sql="select * from qyml where uflag=0"
if keyword<>"" then
   sql=sql & " and "&strField&" like '%"&keyword&"%'"
end if 
sql=sql & " order by ID desc"
rs.open sql,conn,1,1 
if rs.eof and rs.bof then 
response.write "<p align='center'>�Բ���û����Ҫ��ѯ����Ϣ</p>" 
else 
totalPut=rs.recordcount 
if currentpage<1 then 
currentpage=1 
end if 
if (currentpage-1)*MaxPerPage>totalput then 
if (totalPut mod MaxPerPage)=0 then 
currentpage= totalPut \ MaxPerPage 
else 
currentpage= totalPut \ MaxPerPage + 1 
end if 
end if 
if currentPage=1 then 
showContent
call showpage2(sfilename,totalPut,maxperpage,true,true,strUnit)
else 
if (currentPage-1)*MaxPerPage<totalPut then 
rs.move (currentPage-1)*MaxPerPage 
dim bookmark 
bookmark=rs.bookmark 
showContent
call showpage2(sfilename,totalPut,maxperpage,true,true,strUnit)
else 
currentPage=1 
showContent
call showpage2(sfilename,totalPut,maxperpage,true,true,strUnit)
end if 
end if 
rs.close 
end if 
set rs=nothing
sub showContent
%><BR> 
<TABLE border=0 cellPadding=6 cellSpacing=1 width="98%" align="center" bgcolor="#0066FF">
  <TBODY> 
  <Form name="search" method="POST" action="Check.ASP"> 
	    
    <TR bgcolor="#FFFFFF"> 
      <TD height=30> <font color="#ff0000">������ƽ�����Ӧ����</font> </td>
      <TD> <IMG border=0 src="../images/search.gif"></td>
      <TD>&nbsp;�����ؼ���: 
        <input type="text" name="keyword" size="25" style="font-size: 9pt">&nbsp;&nbsp; <select style="font-size: 9pt" name="Field">
<option value="user" selected>�û���</option> <option value="pass">����</option><option value="name">����</option>
</select><INPUT align=absMiddle border=0 src="../images/search1.gif" type=image></td></TR> 
     </FORM>
  </TBODY>
</TABLE><BR>
<CENTER>
  <Form name="myform" method="POST" action=""> 
    <TABLE border="0" cellspacing="1" width="98%" cellpadding="6" bordercolorlight="#D7EBFF" bordercolordark="#D7EBFF" style="border-collapse: collapse" bgcolor="#0066FF">
      <TBODY> 
      <TR height=25 bgcolor="#e8f4ff"> 
        <TD width="60" align="left" height="32" bgcolor="#e8f4ff"><b>ID��</b></TD>
        <TD width="221" align="center" height="23"><b>��˾����</b></TD>
        <TD width="83" align="center" height="23" bgcolor="#e8f4ff"><b>�û���</b></TD>
        <TD width="76" align="center" height="23"><b>����</b></TD>
        <TD width="63" align="center" height="23"><font color="#FF6600"><b>�û�����</b></font></TD>
        <TD width="98" align="center" height="23"><font color="#FF6600"><b>ע��ʱ��</b></font></TD>
        <TD width="75" align="center" height="23"><b>���</b></TD>
        <TD width="42" align="center" height="23">����</TD>
      </TR>
      <%do while not rs.eof%>
      <TR height="20" bgcolor="#ffffff"> 
        <TD width="60" align="center"><font face="Arial"><b><%=rs("id")%></b></font>��</td>
        <TD width="221" align="center">
		<a title="<%=rs("Qymc")%>" href="../../showroom/index.asp?gsid=<%=rs("id")%>" target="_blank"><%=rs("qymc")%></a>
		</td>
        <TD width="83" align="center"><a href='EDIT.asp?id=<%=rs("id")%>' target="_blank"><%=rs("user")%></a></td>
        <TD width="76" align="center"><%=rs("pass")%>��</td>
        <TD width="63" align="center"> <a href="EDIT.asp?id=<%=rs("id")%>" target="_blank"><%=rs("name")%></a> 
        </td>
        <TD width="98" align="center"> <%=rs("idate")%> </td>
        <TD width="75" align="center"> 
          <% if rs("uFlag")=1 then %>
          <a href="SHENGHE.asp?action=cuflag&id=<%=rs("id")%>&page=<%=CurrentPage%>"> 
          <font color="#FF0000">�����</font></a> 
          <%else%>
          <a href="SHENGHE.asp?action=juflag&id=<%=rs("id")%>&page=<%=CurrentPage%>"> 
          <font color="#008000">δ���</font></a> 
          <%end if%>
        </td>
        <TD width="42" align="center"> 
          <input type='checkbox' name='selAnnounce' value='<%=cstr(rs("ID"))%>'>
        </td>
      </TR>
      <% i=i+1
if i>=MaxPerPage then exit do
rs.movenext
loop
%>
      <TR height="20" bgcolor="#ffffff"> 
        <TD colspan="8">���ѡ��ȫ�� 
          <input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;">
          <input type="submit" name="Submit" value="���">
          <input type='submit' name="submit" value='ɾ��'>
        </td>
      </TR>
    </TABLE>
    </form></center>
<%
end sub
sub deleteannounce(id)
dim rs,sql
set rs=server.createobject("adodb.recordset")
sql="select * from [qyml] where id="&cstr(id)
rs.open sql,conn,1,1
if rs.eof and rs.bof then

else
    userid=rs("user")
    conn.execute("delete from [qyml] where user='"&userid&"'")
	conn.execute("delete from [info] where gsid='"&userid&"'")
	conn.execute("delete from [spzs] where gsid='"&userid&"'")
	conn.execute("delete from [qyml] where id="&cstr(id))
	
    if err.Number<>0 then
    err.clear
    response.write "ɾ �� ʧ �� !<br>"
    end if
end if
rs.close
set rs=nothing
End sub

sub pass(id)
conn.execute("update qyml set uflag=1 where id="&cstr(id))
End sub
%>
