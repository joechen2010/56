<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<%
city=request("city")
infoid=request("infoid")
%>

<html>
<head>
<title><%=WebName%>-管理产品信息</title>
<style type="text/css">
<!--
body {
	background-color: #2C68B1;
	margin: 0px;
}
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
<link href="images/main.css" rel="stylesheet" type="text/css">
<script language=JavaScript src="../inc/p_c_a.js"></script>
<script language="JavaScript" type="text/JavaScript">

function check()
{
if (document.Form1.province.value =="省份") 
{ 
alert("请选择区域！"); 
document.Form1.province.focus(); 
return (false); 
} 
if(fIsNumber(document.Form1.prices.value,"1234567890")!=1) {
		alert("请输入数字!");
		document.Form1.prices.focus();
		return false;
	}

if(fIsNumber(document.Form1.tiji.value,"1234567890")!=1) {
		alert("请输入数字!");
		document.Form1.tiji.focus();
		return false;
	}
}

}
function ConfirmDel()
{
   if(confirm("确定要删除该项记录吗？        \n点击确定删除！\n点击取消返回！"))
     return true;
   else
     return false;
	 
}

//****判断是否是Number.
function fIsNumber (sV,sR) {
	var sTmp;
	if(sV.length==0){ return (false);}
	for (var i=0; i < sV.length; i++){
		sTmp= sV.substring (i, i+1);
		if (sR.indexOf (sTmp, 0)==-1) {return (false);}
	}
	return (true);
}
</script>
</head> 
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"> 
<!--#include file="head.htm"-->
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="160" height="566" valign="top" bgcolor="#2C68B1">
<!--#include file="left.htm"-->
	</td>
    <td id="main" align="center" valign="top" bgcolor="#FFFFFF">
<table width="94%"  border="0" cellspacing="0" cellpadding="3">
      <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 已发布的客运专线</td>
          <td width="17%"><img src="images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="534"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> <%=session("user")%> </font> </strong>，欢迎您回来！</td>
        </tr>
      </table>
      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top"><table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="534"  border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="12" height="21" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_1.jpg" width="12" height="25"></td>
                      <td valign="middle" background="images/title_2.jpg" bgcolor="#FFFFFF" >
                        <div align="center"><font color="#FFFFFF">已发布的客运专线</font></div>
                      </td>
                      <td width="13" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_3.jpg" width="15" height="25"></td>
                      <td width="392" valign="middle">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
                <td valign="top" bgcolor="#FFFFFF"> 
                  <table width="520" border="0" cellpadding="4" cellspacing="1" bgcolor="#FF6500">
                    <tr bgcolor="#FFDBBF"> 
                      <td align="center" bgcolor="#FFDBBF">区间</td>
                      <td align="center" bgcolor="#FFCEAA">重量单价</td>
                      <td align="center" bgcolor="#FFCEAA">体积单价</td>
                      <td width="97" bgcolor="#FFCEAA" align="center">操作</td>
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
strUnit="个信息"
    sql="select * from zx_info2 where infoid="&request("infoid")&" order by id desc"
	sfilename="zzz_edit.asp"
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
   	conn.close
   	set conn=nothing
	
sub showContent 
dim i
i=0
%>
                    <%do while not rs.eof%>
                    <tr bgcolor="#FFDBBF"> 
                      <td align="center" bgcolor="#FFDBBF"><%=city%>←→<%=rs("city")%></td>
                      <td align="center" bgcolor="#FFCEAA"><%=rs("prices")%></td>
                      <td align="center" bgcolor="#FFCEAA"><%=rs("tiji")%></td>
                      <td bgcolor="#FFCEAA" align="center">
					  <a href="zzz_Edit2.asp?ID=<%=rs("ID")%>&city=<%=request("city")%>&infoid=<%=request("infoid")%>">修改</a>&nbsp;<a href="zzz_Del2.asp?id=<%=rs("ID")%>&city=<%=request("city")%>&infoid=<%=request("infoid")%>&action=Del" Onclick="return ConfirmDel();" >删除</a></td>
                    </tr>
                    <% 
i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
%>
<%end sub%>
                  </table>
                </td>
            </tr>
          </table>
		  
     <table width="534"  border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td align="center" valign="top">
				 <%if request("action")="" then%>
				  <form method="post" action="zzz_add2.asp?action=add" name="Form2">
				  <input type="hidden" name="city" value="<%=city%>">
				  <input type="hidden" name="infoid" value="<%=infoid%>">
				<input type="submit" name="button2" value="添加途经(中转)城市及报价">
				</form>
				<%end if%>				
				</td>
              </tr>
     </table>		  
		  <%if request("action")="add" then%>
		  
     <table width="534"  border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td align="right" valign="top">
  <form method="post" action="zzz_save2.asp?city=<%=request("city")%>&infoid=<%=request("infoid")%>" name="Form1" onSubmit="javascript:return check(this);">
		                  
                    <table border="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="90" cellspacing="1" bgcolor="#B1D4F2">
                      <tr> 
                        <td colspan="2" height="22" background="images/title_bg.gif"><font color="#FFFFFF">发布货运专线</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td height="22" colspan="2"><%=request("city")%><%'request("infoid")%>←→到达地: 
                          <select id="s1" name="province">
                            <option value="省份">省份</option>
                          </select>
                          <select id="s2" name="city_tj">
                            <option>地级市</option>
                          </select>
	                      </td>
                      </tr>
  <SCRIPT language=javascript>
    setup();
  </SCRIPT>
            <tr bgcolor="#EDF6FF"> 
                        <td width="35%" height="22">重量单价(元/吨): </td>
                        <td width="65%" valign="top" height="22" bgcolor="#EDF6FF">
						<input name="prices" type="text" maxlength="4">						</td>
                      </tr>
             <tr bgcolor="#EDF6FF"> 
                        <td width="35%" height="22">体积单价(元/M3): </td>
                        <td width="65%" valign="top" height="22" bgcolor="#EDF6FF">
						<input name="tiji" type="text" maxlength="4">						</td>
                      </tr>         
                      <tr bgcolor="#EDF6FF"> 
                        <td colspan="2" height="19"> 
                          <input type="submit" value="发 布"  name="button">
                          &nbsp;&nbsp;&nbsp;&nbsp; 
                          <input type="reset" value="重 填">						  </td>
                      </tr>
                    </table>
                  </form> 
				
				</td>
              </tr>
            </table>		  
		  
		  <%end if%>
     <table width="534"  border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td align="right"><a href="#top"><img src="images/icon_top.gif" alt="回到页面顶部" border="0" align="absmiddle" width="15" height="18"></a> 
                  <a href="#top">再检查一遍</a></td>
              </tr>
            </table>
			</td>
        </tr>
      </table>
</td>
  </tr>
</table>

<!--#include file="bottom.htm"--></body></html>