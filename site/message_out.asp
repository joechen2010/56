
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%infoid=request("infoid")
  user=request("user")
%>
		<%
        dim infoid
         infoid=Request.QueryString("infoid")
        if isnumeric(infoid)=0 or infoid="" then
         response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from qyml where id="&request("infoid")&""
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
<HTML>
<HEAD>
<TITLE><%=rs("qymc")%>-诚信物流网</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>
BODY {
	MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px
}
.style1 {
	FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #ffffff
}
.style2 {
	FONT-SIZE: 12px; LINE-HEIGHT: 23px
}
.style3 {
	FONT-SIZE: 12px; LINE-HEIGHT: 17px
}
.style4 {
	FONT-WEIGHT: bold; FONT-SIZE: 14px
}
.style5 {
	FONT-SIZE: 28px
}
A.a03:link {
	FONT-SIZE: 12px; COLOR: #333333; TEXT-DECORATION: none
}
A.a03:visited {
	FONT-SIZE: 12px; COLOR: #333333; TEXT-DECORATION: none
}
A.a03:hover {
	FONT-SIZE: 12px; COLOR: #f46624; TEXT-DECORATION: underline
}
A.a04:link {
	FONT-SIZE: 12px; COLOR: #ffffff; TEXT-DECORATION: none
}
A.a04:visited {
	FONT-SIZE: 12px; COLOR: #ffffff; TEXT-DECORATION: none
}
A.a04:hover {
	FONT-SIZE: 12px; COLOR: #ffffff; TEXT-DECORATION: underline
}
A:link {
	FONT-SIZE: 12px; COLOR: #333333; TEXT-DECORATION: none
}
A:visited {
	FONT-SIZE: 12px; COLOR: #333333; TEXT-DECORATION: none
}
A:hover {
	FONT-SIZE: 12px; COLOR: #f4560d; TEXT-DECORATION: underline
}
A:active {
	FONT-SIZE: 12px; COLOR: #333333; TEXT-DECORATION: none
}
.style6 {
	COLOR: #000000
}
.style81 {
	FONT-SIZE: 12px; COLOR: #000000; LINE-HEIGHT: 20px; FONT-FAMILY: "宋体"
}
.unnamed11 {
	FONT-SIZE: 13px; FONT-FAMILY: "宋体"
}
.style7 {
	FONT-SIZE: 12px; LINE-HEIGHT: 25px
}
BODY {
	FONT-SIZE: 12px
}
TD {
	FONT-SIZE: 12px
}
TH {
	FONT-SIZE: 12px
}
.unnamed111 {
	FONT-SIZE: 13px; LINE-HEIGHT: 23px; FONT-FAMILY: "宋体"
}
.STYLE82 {
	FONT-SIZE: 9px
}
</STYLE>

<SCRIPT language=javascript> 
<!-- 
function Addme(){ 
url = "../b2b/index.asp?providerid=018100000408"; 
title = "苏州市富士达仓储成套设备有限公司-物流搜索 Soo56.com"; 
window.external.AddFavorite(url,title); }
--> 
</SCRIPT>

<SCRIPT>
function safe(frmcp)
{
  var n,thestr,myname,tempstr,alertstr;  
      n=frmcp.length;
//输入不能为空
     tempstr=frmcp.guest.value
     alertstr="请输入发送人！"
     if (tempstr=="")
      {
       alert(alertstr);
       frmcp.guest.focus();
       return false;		 
       }

     tempstr=frmcp.subject.value
     alertstr="请输入主题！"
     if (tempstr=="")
      {
       alert(alertstr);
       frmcp.subject.focus();
       return false;		 
       }


 
     tempstr=frmcp.memo.value
     alertstr="请输入详细说明！"
     if (tempstr=="")
      {
       alert(alertstr);
       frmcp.memo.focus();
       return false;		 
       }



//	所有输入均不允许出现“<”或“>	 
		var n,tempstr,alertstr;
        n=frmcp.length;
        for (j=0;j<n;j++)
	      {  tempstr=frmcp.elements (j).value;
	         if (tempstr.indexOf("'") !=-1)
		     {
		     alertstr="";
		     myname=document.frmcp.elements(j).name;
		     frmcp.elements(j).focus();
		     alertstr="所有输入均不允许出现半角引号“'”！"
		     alert(alertstr);
		     return false;
		     }
	      }
		 


return true;
}
</SCRIPT>

<SCRIPT language=JavaScript type=text/JavaScript>
<!--

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</SCRIPT>

<META content="MSHTML 6.00.2800.1561" name=GENERATOR></HEAD>
<BODY>
<!--#include file="top.htm"-->
<table cellspacing=0 cellpadding=0 width="100%" border=0>
  <tr> 
    <td colspan=2 height=120 background="images/topbg.jpg"> 
      <div class="style1 style5" align=left>&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("qymc")%></div>
    </td>
  </tr>
  <tr> 
    <td width="19%"> 
      <table cellspacing=0 cellpadding=0 width="100%" border=0>
        <tr> 
          <td valign=top width="17%" bgcolor=#e8e8e8 height=185> 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
              <tr> 
                <td> 
                  <!--#include file="left.asp"-->
                </td>
              </tr>
            </table>
          </td>
          <td valign=top width="83%" bgcolor="#fde7db"> 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
              <tr> 
                <td class=style2 bgcolor=#e8e8e8 height=30> 
                  <div align=left>&nbsp; 您现在的位置：<%=rs("qymc")%> &nbsp;&gt;<strong> 
                    客户留言</strong></div>
                </td>
              </tr>
              <tr> 
                <td height=69 valign="top"> 
                  <table cellspacing=0 cellpadding=0 width="97%" align=center border=0>
                    <tbody> 
                    <tr> 
                      <td class=righttitlel width=14 
          background=images/img_subject1_14x30.gif>&nbsp;</td>
                      <td class=righttitlem 
          background=images/img_title_bg_1x30_.gif 
          height=30><strong>客户留言</strong></td>
                      <td class=righttitler width=14 
          background=images/img_subject2_14x30.gif>&nbsp;</td>
                    </tr>
                    </tbody> 
                  </table>
                  <table height=64 cellspacing=0 cellpadding=0 width="97%" align=center border=0>
                    <tr> 
                      <td style="PADDING-RIGHT: 5px; PADDING-LEFT: 15px; PADDING-BOTTOM: 8px; PADDING-TOP: 8px" valign=top bgcolor=#ffffff height=100> 
                        <%if session("user")<>"" then%>
                        <%if request("user")=session("user") then%>
                        <%'*******************************%>
                        <table width="80%"  border="0" cellpadding="0" cellspacing="0" align="center">
                          <tr> 
                            <td valign="top" bgcolor="#FFFFFF"> 
                              <% '////////////////////////////////我发布的全部留言////////////////
    const MaxPerPage=20
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
sfilename="message_out.asp?infoid="&request("infoid")&"&user="&request("user")&""								  
sql="select * from Message where F_User='"&session("user")&"' order by id desc"
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
if rs.eof and rs.bof then 
   response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA"">暂无信息!</td></tr> "
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
   	end if
	sub showContent 
	dim i
	i=0	
%>
          <table border=0 cellpadding=5 cellspacing=1 width="100%" style="border-collapse: collapse" bordercolor="#DEDFDE">
                                <tr align="center" bgcolor="#fde7db"> 
                                  <td height="20" colspan="4"> <b><font color="#FF0000">我发布的全部留言</font></b>                                  </td>
                                </tr>
                                <tr align="center" style="color:White;background-color:#6B696B;font-weight:bold;"> 
                                  <td width="25%" height="20">标题</td>
                                  <td width="25%" height="20">内容</td>
                                  <td width="25%" height="20">发送时间</td>
                                  <td width="25%" height="20">操作</td>
                                </tr>
                                <%do while not rs.eof%>
                                <tr bgcolor="#fde7db"> 
                                  <td height="19"> <a href="javascript:openwindow('read.asp?id=<%=rs("id")%>',800,600)"><%=rs("subject")%></a> 
                                    <%if rs("new")=0 then response.write "(<font color=#ff000>New</font>)"%>                                  </td>
                                  <td height="19"><%=left(rs("content"),10)%>...</td>
                                  <td height="19"><%=rs("TF_Time")%></td>
                                  <td height="19"> <a href="javascript:openwindow('reply.asp?id=<%=rs("id")%>',800,600)">修改</a> 
                                    <a href="msg_del.asp?action_del=del&id=<%=rs("id")%>&page=<%=currentpage%>&infoid=<%=request("infoid")%>&user=<%=request("user")%>&action=ly">删除</a></td>
                                </tr>
  <% 
i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
rs.close
set rs=nothing
%>
  <%end sub %>
                              </table>                           
							 </td>
                          </tr>
                        </table>
						<%end if%>
                        <%end if%>
                        <%if session("user")="" then%>
                              <% Response.Write("<font color='#999999'>为保证信息发布者的隐私权，您需要<a href='../Login/login.asp'>登录</a>才可以查看联系方式</font>")
    Response.Write("<br><a href='../Reg/User_Reg.asp' target='_blank'>未注册用户请点击>></a>")
	%>						
                        <%end if%>
                      </td>
                    </tr>
                    <tbody> </tbody> 
                  </table>
                  <br>
                  <table bordercolor=#FFBD6E width="97%" align="center"  border=1>
                    <tbody> 
                    <tr > 
                      <td bgcolor="#FFFFFF"> 
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td rowspan="2" height="13" width="4%"><img height=22 src="images/209.gif" 
                        width=23></td>
                            <td width="96%" height="22"><font color="#FF0000">免责声明：以上所展示的信息由企业自行提供，内容的真实性、准确性和合法性由发布企业负责。诚信物流网对此不承担任何保证责任。</font></td>
                          </tr>
                          <tr> 
                            <td width="96%" height="22"><font color="#FF0000">友情提醒：为保障您的利益， 
                              建议优先选择诚信物流网会员！ </font></td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td valign=bottom height=30> 
      <hr width="100%" noShade size=1>
    </td>
  </tr>
  <tr> 
    <td valign=top> 
      <!--#include file="bottom.asp"-->
    </td>
  </tr>
</table>
</BODY></HTML>
