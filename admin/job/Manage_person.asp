<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Աɾ��</title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<link rel="stylesheet" type="text/css" href="inc/style.css">
</head>
<!--#include file = "../Inc/Syscode.asp"-->
<%
   	const MaxPerPage=30
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

<body bgcolor="#e1eefd">
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="100%">
<form method=Post action="Manage_person.asp">
<%
    sql="select * from qiuzhi where biaoti<>'' order by idate desc"
	Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1

  	if rs.eof and rs.bof then
       		response.write "<p align='center'> ��û����ְ��Ϣ </p>"
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
           		'showpage totalput,MaxPerPage,"Manage_person.asp"
            		showContent
            		showpage totalput,MaxPerPage,"Manage_person.asp"
       		else
          		if (currentPage-1)*MaxPerPage<totalPut then
            			rs.move  (currentPage-1)*MaxPerPage
            			dim bookmark
            			bookmark=rs.bookmark
           			'showpage totalput,MaxPerPage,"Manage_person.asp"
            			showContent
             			showpage totalput,MaxPerPage,"Manage_person.asp"
        		else
	        		currentPage=1
           			'showpage totalput,MaxPerPage,"Manage_person.asp"
           			showContent
           			showpage totalput,MaxPerPage,"Manage_person.asp"
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
        <table width="98%" border="0" align="center" cellpadding="6" cellspacing="1" class="border" bgcolor="#6699ff">
          <tr class="title" height="22"> 
            <td colspan="7" align="center" height="20" background="Images/Left01.gif"><b>��Աɾ��</b></td>
          </tr>
          <tr class="title" height="22"> 
            <td width="27" align="center" height="20" bgcolor="#e1eefd">&nbsp;</td>
            <td width="37" align="center" bgcolor="#e1eefd"><strong>ID</strong></td>
            <td width="90" align="center" bgcolor="#e1eefd">����</td>
            <td width="59" align="center" bgcolor="#e1eefd">�Ա�</td>
            <td width="179" align="center" bgcolor="#e1eefd">רҵ</td>
            <td width="171" align="center" bgcolor="#e1eefd">����</td>
            <td width="81" align="center" bgcolor="#e1eefd">����ʱ��</td>
          </tr>
          <%do while not rs.eof%>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td height="21" width="27" bgcolor="#FFFFFF"> 
              <p align="center"> 
                <input type='checkbox' name='selAnnounce' value='<%=cstr(rs("id"))%>'>
            </td>
            <td width="37" align="center" bgcolor="#FFFFFF"><%=rs("id")%></td>
            <td width="90" bgcolor="#FFFFFF"><%=rs("iname")%></td>
            <td bgcolor="#FFFFFF"><%=rs("sex")%></td>
            <td bgcolor="#FFFFFF"><%=rs("zye")%></td>
            <td bgcolor="#FFFFFF"><%=rs("biaoti")%></td>
            <td bgcolor="#FFFFFF"><%=rs("idate")%></td>
          </tr>
          <%
	i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
	loop
%>
          <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
            <td height="21" colspan="10" align="center" bgcolor="#FFFFFF"> 
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
    sql="delete from qiuzhi where id="&cstr(id)
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
