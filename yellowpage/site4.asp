
<META http-equiv="Content-Type" content="text/html; chars10_ly3et=gb2312">
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<%'tj=request("tj")
   'action=request("ly")
%>


<table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93" align="center">
              <tr>
                <td valign="top" bgcolor="#FFFFFF"> 
<%if request("user")=session("user") then%>

<% '////////////////////////////////我发布的全部留言////////////////
sql="select * from Message where F_User='"&session("user")&"' order by id desc"
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
if rs.eof and rs.bof then 
   response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA"">暂无信息!</td></tr> "
else 
%>				
                    <table border=0 cellpadding=5 cellspacing=1 width="100%" style="border-collapse: collapse" bgcolor="#B1D4F2">
                           <tr align="center"> 
                            <td height="20" colspan="4" bgcolor="#EDF6FF">
							我发布的全部留言
							</td>
                          </tr>                         
						  <tr align="center"> 
                            <td width="25%" height="20" background="images/title_bg_2.gif">标题</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">内容</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">发送时间</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">操作</td>
                          </tr>
                          <%do while not rs.eof%>
                          <tr> 
                            <td height="19" bgcolor="#EDF6FF">
							<a href="javascript:openwindow('read.asp?id=<%=rs("id")%>',800,600)"><%=rs("subject")%></a> 
                              <%if rs("new")=0 then response.write "(<font color=#ff000>New</font>)"%>                            </td>
                            <td height="19" bgcolor="#EDF6FF"><%=left(rs("content"),10)%>...</td>
                            <td height="19" bgcolor="#EDF6FF"><%=rs("TF_Time")%></td>
                            <td height="19" bgcolor="#EDF6FF">
							<a href="javascript:openwindow('reply.asp?id=<%=rs("id")%>',800,600)">修改</a> 
            <a href="msg_del.asp?action_del=del&id=<%=rs("id")%>&page=<%=currentpage%>&infoid=<%=request("infoid")%>&user=<%=request("user")%>&action=ly">删除</a></td>
                          </tr>
<%
rs.movenext 
loop 
rs.close 
set rs=nothing 
%>
                      </table>
<% '////////////////////////////////我收到的全部留言////////////////
sql="select * from Message where T_User='"&session("user")&"' order by id desc"
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
if rs.eof and rs.bof then 
   response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA"">暂无信息!</td></tr> "
else 
%>				
                    <table border=0 cellpadding=5 cellspacing=1 width="100%" style="border-collapse: collapse" bgcolor="#B1D4F2">
                           <tr align="center"> 
                            <td height="20" colspan="4" bgcolor="#EDF6FF">
							我收到的全部留言
							</td>
                          </tr>                         
						  <tr align="center"> 
                            <td width="25%" height="20" background="images/title_bg_2.gif">标题</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">内容</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">发送时间</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">操作</td>
                          </tr>
                          <%do while not rs.eof%>
                          <tr> 
                            <td height="19" bgcolor="#EDF6FF">
							<a href="javascript:openwindow('read2.asp?id=<%=rs("id")%>',800,600)"><%=rs("subject")%></a> 
                              <%if rs("new")=0 then response.write "(<font color=#ff000>New</font>)"%>                            </td>
                            <td height="19" bgcolor="#EDF6FF"><%=left(rs("content"),10)%>...</td>
                            <td height="19" bgcolor="#EDF6FF"><%=rs("TF_Time")%></td>
                            <td height="19" bgcolor="#EDF6FF">
							<a href="javascript:openwindow('reply2.asp?T_User=<%=rs("F_User")%>',800,600)">回复</a> 
            <a href="msg_del.asp?action_del=del&id=<%=rs("id")%>&page=<%=currentpage%>&infoid=<%=request("infoid")%>&user=<%=request("user")%>&action=ly">删除</a></td>
                          </tr>
<%
rs.movenext 
loop 
rs.close 
set rs=nothing 
%>
<%end if%>
<%end if%>


                        </table>
						<%end if%>
                  </td>
                </tr>
              </table>
<%if request("user")<>session("user") then%>

<% 
sql="select * from Message where F_User='"&session("user")&"' and T_User='"&request("user")&"' order by id desc"
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
if rs.eof and rs.bof then 
   response.write "<tr align=""center"" bgcolor=""#FFDBBF"" width=534><td colspan=""7"" bgcolor=""#FFCEAA"">暂无信息!</td></tr> "
else 
%>				
<table border=0 cellpadding=5 cellspacing=1 width=534 style="border-collapse: collapse" bgcolor="#B1D4F2" align="center">
                           <tr align="center"> 
<%set rs2=server.createobject("adodb.recordset")
sql2="select * from qyml where user='"&request("user")&"'"		
rs2.open sql2,conn,1,1	
if rs2.eof and rs2.bof then
    response.write "没有数据"
else		  
%>                       <td colspan="4" align="center" bgcolor="#EDF6FF">我给<font color="#FF0000"><%=rs2("qymc")%></font>的留言</td>
<%end if%>
                          </tr>                         
						  <tr align="center"> 
                            <td width="25%" height="20" background="images/title_bg_2.gif">标题</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">内容</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">发送时间</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">操作</td>
                          </tr>
                          <%do while not rs.eof%>
                          <tr> 
                            <td height="19" bgcolor="#EDF6FF">
							<a href="javascript:openwindow('read.asp?id=<%=rs("id")%>',800,600)"><%=rs("subject")%></a> 
                              <%if rs("new")=0 then response.write "(<font color=#ff000>New</font>)"%>                            </td>
                            <td height="19" bgcolor="#EDF6FF"><%=left(rs("content"),10)%>...</td>
                            <td height="19" bgcolor="#EDF6FF"><%=rs("TF_Time")%></td>
                            <td height="19" bgcolor="#EDF6FF">
							<a href="javascript:openwindow('reply.asp?id=<%=rs("id")%>',800,600)">修改</a> 
            <a href="msg_del.asp?action_del=del&id=<%=rs("id")%>&page=<%=currentpage%>&infoid=<%=request("infoid")%>&user=<%=request("user")%>&action=ly">删除</a></td>
                          </tr>
<%
rs.movenext 
loop 
rs.close 
set rs=nothing 
end if
%>
</table>

<table border=0 cellpadding=5 cellspacing=1 style="border-collapse: collapse" bgcolor="#B1D4F2" width="534" align="center">
     <tr align="center"> 
       <td height="20" colspan="4" bgcolor="#EDF6FF">
			<a href="javascript:openwindow('tj.asp?T_User=<%=request("user")%>',800,600)">添加新留言</a>
		</td>
     </tr>
</table>

<%end if%>