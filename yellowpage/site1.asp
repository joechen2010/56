
<META http-equiv="Content-Type" content="text/html; chars3et=gb2312">
<table cellspacing="0" cellpadding="2" rules="cols" bordercolor="#DEDFDE" border="1" width="778" align="center">
  <tr align="Center" style="color:White;background-color:#6B696B;font-weight:bold;"> 
    <td height="25"> 
      <div align="center">��</div>
    </td>
    <td height="25"> 
      <div align="center">�����ص�</div>
    </td>
    <td height="25"> 
      <div align="center">Ŀ�ĵ�</div>
    </td>
    <td height="25"> 
      <div align="center">����</div>
    </td>
    <td height="25"> 
      <div align="center">����(��)</div>
    </td>
    <td height="25"> 
      <div align="center">��������</div>
    </td>
    <td height="25"> 
      <div align="center">����</div>
    </td>
  </tr>
                    <%

				
    sql="select  * from info2 where (infotype='��Դ��Ϣ' or infotype='��Դ��Ϣ') and gsid='"&user&"' order by addtime desc"
	sfilename="index.asp"
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
  	if rs.eof and rs.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> �� û �� �� �� �� Ϣ</td></tr> </p>"
   	else


 

%>  

				<%do while not rs.eof%>
  <tr align="Center" style="background-color:#F7F7DE;"> 
    <td> 
	<%if rs("infotype")="��Դ��Ϣ" then %>
      <div align="center"><img src="../images/car.gif" width="25" height="20">
      </div>
	  <%else%>
	        <div align="center"><img src="../images/goods.gif" width="25" height="20">
      </div>
	  <%end if%>
    </td>
    <td align="Right"> 
      <div align="center"><%=rs("city")%>&nbsp;<%=rs("area")%>��</div>
    </td>
    <td align="Left"> 
      <div align="center"><%=rs("city2")%>&nbsp;<%=rs("area2")%></div>
    </td>
    <td> 
      <div align="center"><%=rs("cartype")%></div>
    </td>
    <td style="font-family:����;"> 
      <div align="center"><%=rs("carload")%>��</div>
    </td>
    <td> 
      <div align="center"><%=rs("addtime")%></div>
    </td>
    <td> 
      <div align="center"><A href="javascript:openwindow('../truck/detail.asp?InfoID=<%=rs("InfoID")%>',500,400)">����</A></div>
    </td>
  </tr>
<% 

rs.movenext 
loop
end if 
%>

</table>