<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
province=request("province")
infotype=request("infotype")
city=request("city")
kftype=request("kftype")
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<script language="javascript">
function openwindow(url,width,height) { 	left1 = (screen.width-800)/2; 	top1 = (screen.height-350)/2; 	window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); }

</script>
</head>

<body>
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#3366CC" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="500">
    <tr align="Center" style="color:#CCCCFF;background-color:#003399;font-weight:bold;">		
      <td style="width:130px;" width="128">��Ϣ����</td>
      <td width="80">�ⷿ����</td>
      <td style="width:160px;" width="160">�ֿ�ص�</td>
      <td style="width:160px;" width="168">���</td>
      <td style="width:160px;" width="150">�۸�</td>
	  <td style="width:160px;" width="150">��������</td>
      <td style="width:30px;" width="30" align="center">����</td>
	</tr>
<%
    const MaxPerPage=15
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
strUnit="����Ϣ"
dim sql2,city,infotype,kftype
dim rs2

        if request("city")<>"�ؼ���" then
	    c=" and province='"&request("province")&"' and city='"&request("city")&"'"
		 else
		 if request("province")<>"ʡ��" then
		 c=" and province='"&request("province")&"'"
		 end if
	    end if
	    if request("infotype")<>"" then
	    info_type=" and infotype='"&request("infotype")&"'"
	    end if
	    if request("area")<>"" then
	    area=" and area='"&request("area")&"'"
	    end if
	    if request("kftype")<>"" then
	    kf_type=" and kftype='"&request("kftype")&"'"
	    end if					
    sql="select  * from kf_info where 1=1"&c&info_type&kf_type&" order by infoid desc"
	sfilename="search.asp?province="&request("province")&"&city="&request("city")&"&infotype="&request("infotype")&"&kftype="&request("kftype")&""
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
  	if rs.eof and rs.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> �� û �� �� �� �� Ϣ</td></tr> </p>"
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
	
	<tr align="Center" style="color:#003399;background-color:White;">		
      <td><%=rs("infotype")%></td>
      <td  align="center"><%=rs("kftype")%></td>
      <td align="center"><%=rs("city")%>&nbsp;<%=rs("area")%></td>
      <td align="center"><%=rs("mji")%>ƽ����</td>
      <td  align="center"><%=rs("prices")%>Ԫ/ƽ����</td>
	  <td align="center"><%=rs("addtime")%></td>
      <td><A href="javascript:openwindow('detail.asp?InfoID=<%=rs("InfoID")%>',500,400)">����</A></td>
	</tr>
                    <% 
i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
%>
<%end sub%> 
</table>
</body>
</html>
