<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
province=request("province")
province2=request("province2")
city=request("city")
city2=request("city2")
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<LINK href="images/page.css" type=text/css rel=stylesheet>
<script language="javascript">
function openwindow(url,width,height) { 	left1 = (screen.width-800)/2; 	top1 = (screen.height-350)/2; 	window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); }

</script>
</head>

<BODY leftMargin=0 topMargin=0>
<TABLE cellSpacing=0 cellPadding=0 width=500 border=0 align="center">
        <TR> 
          <TD class=bg width=7 background=images/bg02.gif bgColor=#ffe8d9>&nbsp;</TD>
          <TD> 
            <TABLE cellSpacing=0 cellPadding=0 border=0 width="100%">
              <TR align="center"> 
                <TD bgColor=#ff6600 height="26"><STRONG><FONT 
                        color=#ffffff>专线名称</FONT></STRONG></TD>
                <TD bgColor=#ff6600><STRONG><FONT 
                        color=#ffffff>途经(中转)城市</FONT></STRONG></TD>
                <TD bgColor=#ff6600><STRONG><FONT 
                        color=#ffffff>公司名称</FONT></STRONG></TD>
                <TD bgColor=#ff6600><STRONG><font 
                        color=#ffffff>专线</font></STRONG></TD>
                <TD bgColor=#ff6600><STRONG><FONT 
                        color=#ffffff>发布时间</FONT></STRONG></TD>
              </TR>
<%
    const MaxPerPage=10
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
dim sql
dim rs
dim qymc
dim strUnit,sfilename
strUnit="个信息"
        if request("city")<>"地级市" then
	      p=" and a.province='"&request("province")&"' and a.city='"&request("city")&"'"
		  else 
		  if request("province")<>"省份" then
	       p=" and a.province='"&request("province")&"'"
		  else
		   p=""
		  end if
	    end if
        if request("city2")<>"地级市" then
	      p2=" and (province2='"&request("province2")&"' and city2='"&request("city2")&"')"
		  else 
		  if request("province2")<>"省份" then
	       p2=" and province2='"&request("province2")&"'"
		  else
		   p2=""
		  end if	    
		end if				   
    sql="select * from zx_info a,qyml b where 1=1"&p&p2&" and infotype='货运专线' and a.gsid=b.user order by infoid desc"
	sfilename="search.asp?province="&request("province")&"&province2="&request("province2")&"&city="&request("city")&"&city2="&request("city2")&""
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
  	if rs.eof and rs.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> 还 没 有 任 何 信 息</td></tr> </p>"
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
            		call showpage2(sfilename,totalPut,maxperpage,false,false,strUnit)
       		else
          		if (currentPage-1)*MaxPerPage<totalPut then
            			rs.move  (currentPage-1)*MaxPerPage
            			dim bookmark
            			bookmark=rs.bookmark
            			showContent
             			call showpage2(sfilename,totalPut,maxperpage,false,false,strUnit)
        		else
	        		currentPage=1
           			showContent
           			call showpage2(sfilename,totalPut,maxperpage,false,false,strUnit)
	      		end if
	   	end if
   	rs.close
   	end if
	set rs=nothing  

	
sub showContent 
dim i
i=0
%>  

				<%do while not rs.eof%>		  	  
              <TR onMouseOver="this.style.backgroundColor='#FAEFE0'" 
                    style="CURSOR: hand; BACKGROUND-COLOR: #f9f9f9" 
                    onmouseout="style.backgroundColor='#F9F9F9'" align="center"> 
                <TD height=25><%=rs("a.city")%>&nbsp;<%=rs("a.area")%>←→<%=rs("city2")%>&nbsp;<%=rs("area2")%>
                </TD>
                <TD height=25>
					  <%
						   set rs1=server.CreateObject("adodb.recordset")
						    sql1 = "select * from zx_info2 where infoid="&rs("infoid")&" order by id desc"
						   rs1.open sql1,conn,1,1
						   if not rs1.eof then
						    for j=1 to rs1.recordcount
						   %>
					  <%=rs1("city")%>&nbsp;
                          <%
							  rs1.movenext
							  next
							  end if
							  %>				
				</TD>
                <TD height=25>
				  <%if len(rs("qymc"))>8 then%>
				  <%=left(rs("qymc"),6)%>...
				  <%else%>
				 <%=rs("qymc")%>
				  <%end if%>				
				</TD>
                <TD height=25><A href="javascript:openwindow('detail.asp?InfoID=<%=rs("InfoID")%>',500,400)">详情</A></TD>
                <TD><%=rs("addtime")%></TD>
              </TR>
<% 
i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
%>
<%end sub %> 
            </TABLE>
          </TD>
          <TD class=bg width=7 background=images/bg02.gif bgColor=#ffe8d9>&nbsp;</TD>
        </TR>
      </TABLE>
</body>
</html>
