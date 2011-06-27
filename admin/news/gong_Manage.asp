<!--#include file = "../Inc/Syscode.asp"-->
<% 
'============================================
'
'
'
'=============================================
   	const MaxPerPage=18
   	dim totalPut   
   	dim CurrentPage
   	dim TotalPages
   	dim i,j
   	dim idlist,NClassID
 
   	if not isempty(request("page")) then
      		currentPage=cint(request("page"))
   	else
      		currentPage=1
   	end if
	
   Dim NewsID
   NewsID=Request("NewsID")
   
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<link rel="stylesheet" type="text/css" href="inc/style.css">
<link href="../inc/admin_style.css" rel="stylesheet" type="text/css">
</head>
<%
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
<form method=Post action="gong_Manage.asp">
<%
	     Sql="select * from gonggao order by id desc"
	'NewsID=Rs(0) NClassID=Rs(1) NClassName=Rs(2) Title=Rs(3) Editor=Rs(4) Hits=Rs(5)  IncludePic=Rs(6)
	'Recommend=Rs(7)
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1

  	if rs.eof and rs.bof then
       		response.write "<p align='center'> 还 没 有 任 何 文 章 </p>"
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
           		'showpage totalput,MaxPerPage,"News_Manage.asp"
            		showContent
            		showpage totalput,MaxPerPage,"gong_Manage.asp"
       		else
          		if (currentPage-1)*MaxPerPage<totalPut then
            			rs.move  (currentPage-1)*MaxPerPage
            			dim bookmark
            			bookmark=rs.bookmark
           			'showpage totalput,MaxPerPage,"News_Manage.asp"
            			showContent
             			showpage totalput,MaxPerPage,"gong_Manage.asp"
        		else
	        		currentPage=1
           			'showpage totalput,MaxPerPage,"News_Manage.asp"
           			showContent
           			showpage totalput,MaxPerPage,"gong_Manage.asp"
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
            <td width="40" align="center" height="20">&nbsp;</td>
            <td width="51" align="center"><strong>ID</strong></td>
            <td align="center">新闻标题</td>
          </tr>
          <%do while not rs.eof%>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td height="21" width="40"> 
              <p align="center"> 
                <input type='checkbox' name='selAnnounce' value='<%=cstr(rs(0))%>'>
            </td>
            <td width="51" align="center"><%=rs(0)%></td>
            <td><a href="gg_change.asp?id=<%=rs(0)%>"><%=rs(1)%></a></td>
          </tr>
          <%
	i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
	loop
%>
          <tr align="center" class="tdbg"> 
            <td height="21" colspan="3">
              <input type='submit' class=buttonface value='&nbsp;删&nbsp;&nbsp;除&nbsp;'>            </td>
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
    response.write "<font color='#000080'>首页 上一页</font>&nbsp;"
  else
    response.write "<a href="&filename&"?page=1>首页</a>&nbsp;"
    response.write "<a href="&filename&"?page="&CurrentPage-1&">上一页</a>&nbsp;"
  end if
  if n-currentpage<1 then
    response.write "<font color='#000080'>下一页 尾页</font>"
  else
    response.write "<a href="&filename&"?page="&(CurrentPage+1)&">"
    response.write "下一页</a> <a href="&filename&"?page="&n&">尾页</a>"
  end if
   response.write "<font color='#000080'>&nbsp;页次：</font><strong><font color=red>"&CurrentPage&"</font><font color='#000080'>/"&n&"</strong>页</font> "
    response.write "<font color='#000080'>&nbsp;共&nbsp;<b>"&totalnumber&"</b>&nbsp;项&nbsp;&nbsp; <b>"&maxperpage&"</b>&nbsp;项/页</font> "
       
end function

sub deleteannounce(id)
    dim rs,sql
    set rs=server.createobject("adodb.recordset")
    sql="delete from gonggao where ID="&cstr(id)
    conn.execute sql
End sub
%>
      </form></td>
  </tr>
</table>
</body>
</html>
