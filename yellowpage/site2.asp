
<META http-equiv="Content-Type" content="text/html; chars3et=gb2312">
<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%">
              <TR align="center"> 
                <TD bgColor=#ff6600 height="26"><STRONG><FONT 
                        color=#ffffff>ר������</FONT></STRONG></TD>
                <TD bgColor=#ff6600><STRONG><FONT 
                        color=#ffffff>;��(��ת)����</FONT></STRONG></TD>
                <TD bgColor=#ff6600><STRONG><FONT 
                        color=#ffffff>��˾����</FONT></STRONG></TD>
                <TD bgColor=#ff6600><STRONG><font 
                        color=#ffffff>ר��</font></STRONG></TD>
                <TD bgColor=#ff6600><STRONG><FONT 
                        color=#ffffff>����ʱ��</FONT></STRONG></TD>
              </TR>
                    <%
					dim rs3
    sql="select  * from zx_info a,qyml b where infotype='����ר��' and b.user='"&user&"' and  a.gsid='"&user&"' order by infoid desc"
    Set rs3= Server.CreateObject("ADODB.Recordset")
	rs3.open sql,conn,1,1
  	if rs3.eof and rs3.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> �� û �� �� �� �� Ϣ</td></tr> </p>"
   	else
      	
	
'dim i
'i=0
%>  

				<%do while not rs3.eof%>		  	  
              <TR onMouseOver="this.style.backgroundColor='#FAEFE0'" 
                    style="CUrs3OR: hand; BACKGROUND-COLOR: #f9f9f9" 
                    onmouseout="style.backgroundColor='#F9F9F9'" align="center"> 
                <TD height=25><%=rs3("a.city")%>&nbsp;<%=rs3("a.area")%>����<%=rs3("city2")%>&nbsp;<%=rs3("area2")%>
                </TD>
                <TD height=25>
					  <%
						   set rs31=server.CreateObject("adodb.recordset")
						    sql1 = "select * from zx_info2 where infoid="&rs3("infoid")&" order by id desc"
						   rs31.open sql1,conn,1,1
						   if not rs31.eof then
						    for j=1 to rs31.recordcount
						   %>
					  <%=rs31("city")%>&nbsp;
                          <%
							  rs31.movenext
							  next
							  end if
							  %>				
				</TD>
                <TD height=25><%=rs3("qymc")%></TD>
                <TD height=25><A href="javascript:openwindow('../fline/detail.asp?InfoID=<%=rs3("InfoID")%>',500,400)">����</A></TD>
                <TD><%=rs3("addtime")%></TD>
              </TR>
                    <% 
'i=i+1
rs3.movenext 
loop 
%>
<%end if%>
</TABLE>