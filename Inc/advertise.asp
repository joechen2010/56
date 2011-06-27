<!--#include file="conn.asp"-->
<%sub advertise(adclass)%>
<TABLE cellSpacing=0 cellPadding=2 width=310 align=center border=0>
<%
set rs=server.CreateObject("adodb.recordset")
sql="select top 10 webpic,weburl from advertise where adclass="&adclass&" order by id desc"
rs.open sql,conn,1,1
if not (rs.eof and rs.bof) then
  do while not rs.eof 
%>
  <TR> 
    <TD align="center">
<%
   response.Write "<a href='"&rs("weburl")&"' target='_blank'><img src='../admin/advertise/uppic/"&rs("webpic")&"' border='0'></a><br>"
%>
	</TD>
    </TR>
<%
  rs.movenext
  loop
  end if
  rs.close
  set rs=nothing
%>	
</TABLE>
<%end sub%>

