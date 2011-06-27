<!--#include file = "../Inc/Syscode.asp"-->
<%
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
dim action
action=request("action")
if action="确定" then
     conn.execute("update qyml set flag="&request("hydj")&",BeginDate='"&Date()&"',EndDate='"&date()+365&"' where user='"&request("user")&"'")
	 response.write "<script>alert('会员操作成功！');history.back();</script>"
elseif action="删除" then
     conn.execute("delete from [qyml] where user='"&request("user")&"'")
	 conn.execute("delete from [info] where gsid='"&request("user")&"'")
	 conn.execute("delete from [spzs] where gsid='"&request("user")&"'")
	 response.write "<script>alert('数据删除成功！');history.back();</script>"
else
%>
<link rel="stylesheet" type="text/css" href="../style.css"> 
<body bgcolor="#e1eefd" marginheight=0 marginwidth=0 leftmargin=0>

<%
dim keyword,strField,sfilename
strField=trim(request("Field"))
keyword=trim(request("keyword"))

if keyword<>"" then 
	keyword=ReplaceBadChar(keyword)
end if
dim jb
jb=request("jb")

sfilename="check.asp?Field=" & strField & "&Keyword=" & keyword & "&jb=" & jb & ""

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



dim rs
dim sql
set rs=server.createobject("adodb.recordset") 
sql="select * from qyml where uflag=1"
if keyword<>"" then
   sql=sql & " and "&strField&" like '%"&keyword&"%'"
end if 
if jb<>"" then
    sql=sql & "and flag="&jb&""
end if
sql=sql & " order by ID desc"
rs.open sql,conn,1,1 
if rs.eof and rs.bof then 
response.write "<p align='center'>对不起，没有您要查询的信息</p>" 
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
showContent totalput,MaxPerPage
call showpage2(sfilename,totalPut,maxperpage,true,true,strUnit)
else 
if (currentPage-1)*MaxPerPage<totalPut then 
rs.move (currentPage-1)*MaxPerPage 
dim bookmark 
bookmark=rs.bookmark 
showContent totalput,MaxPerPage
call showpage2(sfilename,totalPut,maxperpage,true,true,strUnit)
else 
currentPage=1 
showContent totalput,MaxPerPage
call showpage2(sfilename,totalPut,maxperpage,true,true,strUnit)
end if 
end if 
rs.close 
end if 
set rs=nothing
sub showContent (totalput,MaxPerPage)
%><BR> 
<TABLE border=0 cellPadding=6 cellSpacing=1 width="96%" align="center" bgcolor="#0099FF">
  <TBODY> 
  <Form name="search" method="POST" action="Check.ASP"> 
	    
    <TR bgcolor="#FFFFFF"> 
      <TD height=30> <font color="#ff0000">点击名称进行相应操作</font> </td>
      <TD> <IMG border=0 src="../images/search.gif"></td>
      <TD>&nbsp;搜索关键字: 
        <input type="text" name="keyword" size="25" style="font-size: 9pt">&nbsp;&nbsp; 
		<select style="font-size: 9pt" name="Field">
          <option value="user" selected>用户名</option> <option value="qymc">企业名称</option>
        </select>&nbsp;<INPUT align=absMiddle border=0 src="../images/search1.gif" type=image></td></TR> 
     </FORM>
  </TBODY>
</TABLE><BR> 
<TABLE border="0" cellspacing="1" width="96%" cellpadding="6" bordercolorlight="#D7EBFF" bordercolordark="#D7EBFF" style="border-collapse: collapse" align="center" bgcolor="#0099FF">
  <TBODY> 
  <TR height=25 bgcolor="#e8f4ff"> 
    <TD width="34" align="left" height="32"><b>ID号</b></TD>
    <TD width="166" align="center"><b>用户名</b></TD>
    <TD width="76" align="center"><b>密码</b></TD>
    <TD width="195" align="center"><b><font color="#FF6600">企业名称</font></b></TD>
    <TD width="74" align="center"><font color="#FF6600"><b>注册时间</b></font></TD>
    <TD width="46" align="center"><b>审核</b></TD>
    <TD width="137" align="center"><font color="#FF6600"><b>操作</b></font></TD>
  </TR>
  <%do while not rs.eof%>
  <TR height="20" bgcolor="#ffffff"> 
    <TD width="34" align="center" height="20"><font face="Arial"><b><%=rs("id")%></b></font>　</td>
    <TD align="center" width="166" height="20"><a title="<%=rs("Qymc")%>" href='edit.asp?id=<%=rs("id")%>' target="_blank"><%=rs("user")%></a>（ 
      <%if rs("uflag")=1 then 
	response.write "已审核" 
	else 
	response.write "未审核"
	end if%>
      ）</td>
    <TD width="76" align="center" height="20"><%=rs("pass")%>　</td>
    <TD width="195" height="20"> <b><a title="<%=rs("Qymc")%>" href='../../showroom/index.asp?gsid=<%=rs("id")%>' target="_blank"><%=rs("qymc")%></a></b></td>
    <TD width="74" align="center" height="20"> <%=rs("idate")%></td>
    <TD colspan="2" align="center" height="20"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0"><form name="form1" method="post" action="">
          <tr> 
            <td width="42%"> 
              <select name="hydj" size="1">
                <option value="0" <%if rs("flag")=0 then response.write "selected"%>>免费会员</option>
                <option value="1" <%if rs("flag")=1 then response.write "selected"%>>铜牌会员</option>
                <option value="2" <%if rs("flag")=2 then response.write "selected"%>>银牌会员</option>
                <option value="3" <%if rs("flag")=3 then response.write "selected"%>>金牌会员</option>
              </select></td>
            <td width="58%"> 
              <input type="submit" name="action" value="确定">
              <input type="submit" name="action" value="删除">
              <input type="hidden" name="user" value="<%=rs("user")%>">
            </td>
          </tr></form>
        </table>  
    </td>
  </TR>
  <% i=i+1
if i>=MaxPerPage then exit do
rs.movenext
loop
%>
</TABLE>
<%
end sub
end if
%>