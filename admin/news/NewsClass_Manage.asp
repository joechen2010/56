<!--#include file = "../Inc/Syscode.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<link rel="stylesheet" type="text/css" href="../inc/admin_style.css">
</head>


<body>
<div align="center"><center> <table border="0" width="92%" cellspacing="0" cellpadding="0"> 
<tr> <td width="100%"><p align="center"><br> <form method=Post action="NewsClass_Manage.asp"> 
<div align="center"> <%
   	const MaxPerPage=18
   	dim totalPut   
   	dim CurrentPage
   	dim TotalPages
   	dim i,j
   	dim idlist

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

	sql="select * from News_Class order by ClassID desc"
	Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1

  	if rs.eof and rs.bof then
       		response.write "<p align='center'> �� û �� �� �� �� �� </p>"
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
           		'showpage totalput,MaxPerPage,"NewsClass_Manage.asp"
            		showContent
            		showpage totalput,MaxPerPage,"NewsClass_Manage.asp"
       		else
          		if (currentPage-1)*MaxPerPage<totalPut then
            			rs.move  (currentPage-1)*MaxPerPage
            			dim bookmark
            			bookmark=rs.bookmark
           		'	showpage totalput,MaxPerPage,"NewsClass_Manage.asp"
            			showContent
             			showpage totalput,MaxPerPage,"NewsClass_Manage.asp"
        		else
	        		currentPage=1
           		'	showpage totalput,MaxPerPage,"NewsClass_Manage.asp"
           			showContent
           			showpage totalput,MaxPerPage,"NewsClass_Manage.asp"
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
%> <div align="center"><center><br><table class="border" border="0" cellspacing="1" width="80%" cellpadding="0"> 
<tr class="title" height="22"> <td width="82" align="center" bgcolor="#D0D0D0" height="20"><strong>ID</strong></td><td width="300" align="center" bgcolor="#D0D0D0"><strong>һ����Ŀ����</strong></td><td width="300" align="center" BGCOLOR="#D0D0D0">����</td><td width="208" align="center" bgcolor="#D0D0D0"><input type='submit' class=buttonface value='&nbsp;ɾ&nbsp;&nbsp;��&nbsp;'></td></tr> 
<%do while not rs.eof%> <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
<td height="23" width="82"><p align="center"><%=rs("ClassID")%></td><td align="center"><a href="NewsNClass_Manage.asp?ClassID=<%=rs("ClassID")%>"><%=rs("ClassName")%></a></td><td width="300" align="center"> 
<a href="NewsClass.asp?action=modify&ClassID=<%=rs("ClassID")%>">�޸�</a>����<a href="NewsNClass.asp?action=add&ClassID=<%=rs("ClassID")%>">�������Ŀ</a></td><td width="208" align="center"><input type='checkbox' name='selAnnounce' value='<%=cstr(rs("ClassID"))%>'></td></tr> 
<%
	i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
	loop
%> </table></center></div><%
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
    sql="delete from News_Class where ClassID="&cstr(id)
    conn.execute sql
  '  if err.Number<>0 then
	'err.clear
	'response.write "ɾ �� ʧ �� !<br>"
  '  else
    'response.write "�����ɹ���<br>"
   ' end if
  End sub
%> </div></form></td></tr> </table></center></div>
</body>
</html>
