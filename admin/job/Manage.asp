<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<link rel="stylesheet" type="text/css" href="inc/style.css">
</head>
<!--#include file = "../Inc/Syscode.asp"-->
<%
   	const MaxPerPage=18
   	dim totalPut   
   	dim CurrentPage
   	dim TotalPages
   	dim i,j
   	dim idlist
  
    ClassID=trim(request("ClassID"))
	title=request("txtitle")
   	if not isempty(request("page")) then
      		currentPage=cint(request("page"))
   	else
      		currentPage=1
   	end if
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
     		else
			call deleteannounce(clng(idlist))
     		end if
  	end if 
   	dim sql
   	dim rs
%>

<body>
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="100%">
<form method=Post action="Manage.asp">
<%
    sql="select * from zhaopin order by fbsj desc"
	Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1

  	if rs.eof and rs.bof then
       		response.write "<p align='center'> �� û �� �� �� �� ˾ </p>"
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
           		'showpage totalput,MaxPerPage,"Manage.asp"
            		showContent
            		showpage totalput,MaxPerPage,"Manage.asp"
       		else
          		if (currentPage-1)*MaxPerPage<totalPut then
            			rs.move  (currentPage-1)*MaxPerPage
            			dim bookmark
            			bookmark=rs.bookmark
           			'showpage totalput,MaxPerPage,"Manage.asp"
            			showContent
             			showpage totalput,MaxPerPage,"Manage.asp"
        		else
	        		currentPage=1
           			'showpage totalput,MaxPerPage,"Manage.asp"
           			showContent
           			showpage totalput,MaxPerPage,"Manage.asp"
	      		end if
	   	end if
   	rs.close
   	end if
	        
   	set rs=nothing  
   	conn.close
   	set conn=nothing
  
   	sub showContent
       	dim i
	   i=0
%>
      <br>
        <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
          <tr class="title" height="22"> 
            <td width="25" align="center" bgcolor="#D0D0D0" height="20">&nbsp;</td>
            <td width="33" align="center" bgcolor="#D0D0D0"><strong>ID</strong></td>
            <td width="279" align="center" bgcolor="#D0D0D0">��Ƹ��λ����</td>
            <td width="234" align="center" bgcolor="#D0D0D0">��Ƹְλ</td>
            <td width="215" align="center" bgcolor="#D0D0D0">����ʱ��</td>
            <td width="138" align="center" bgcolor="#D0D0D0">��Ч��</td>
          </tr>
          <%do while not rs.eof%>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td height="21" width="25"> 
              <p align="center"> 
                <input type='checkbox' name='selAnnounce' value='<%=cstr(rs("id"))%>'>
            </td>
            <td width="33" align="center"><%=rs("id")%></td>
            <td width="279"><%=rs("sqmc")%></td>
            <td><%=rs("zpzw")%></td>
            <td><%=rs("fbsj")%></td>
            <td><%=rs("hfsj")%></td>
          </tr>
          <%
	i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
	loop
%>
          <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
            <td height="21" colspan="9" align="center"> 
              <input type='submit' class=buttonface value='&nbsp;ɾ&nbsp;&nbsp;��&nbsp;'>            </td>
          </tr>
        </table>
<%
   end sub 

function showpage(totalnumber,maxperpage,filename)
  dim n
  if totalnumber mod maxperpage=0 then
     n= totalnumber \ maxperpage
  else
     n= totalnumber \ maxperpage+1
  end if
  response.write "<p align='center'>&nbsp;"
  if CurrentPage<2 then
    response.write "<font color='#000080'>��ҳ ��һҳ</font>&nbsp;"
  else
    response.write "<a href="&filename&"?page=1>��ҳ</a>&nbsp;"
    response.write "<a href="&filename&"?page="&CurrentPage-1&">��һҳ</a>&nbsp;"
  end if
  if n-currentpage<1 then
    response.write "<font color='#000080'>��һҳ βҳ</font>"
  else
    response.write "<a href="&filename&"?page="&(CurrentPage+1)&">"
    response.write "��һҳ</a> <a href="&filename&"?page="&n&">βҳ</a>"
  end if
   response.write "<font color='#000080'>&nbsp;ҳ�Σ�</font><strong><font color=red>"&CurrentPage&"</font><font color='#000080'>/"&n&"</strong>ҳ</font> "
    response.write "<font color='#000080'>&nbsp;��&nbsp;<b>"&totalnumber&"</b>&nbsp;��&nbsp;&nbsp; <b>"&maxperpage&"</b>&nbsp;��/ҳ</font> "
       
end function

  sub deleteannounce(id)
    dim rs,sql
    set rs=server.createobject("adodb.recordset")
    sql="delete from zhaopin where id="&cstr(id)
    conn.execute sql
	
  '  if err.Number<>0 then
	'err.clear
	'response.write "ɾ �� ʧ �� !<br>"
  '  else
    'response.write "�����ɹ���<br>"
   ' end if
  End sub
%>
</form></td>
  </tr>
</table>
</body>
</html>
