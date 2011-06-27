<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="News_Code.asp"-->
<% 
'============================================
'
'
'
'=============================================
Dim NClassID,NewsNClassName
    NClassID=Saferequest("NClassID",1)
if NClassID="" or NClassID=0 then
   call WriteErrMsg(errmsg)
end if
set rs_nc=conn.execute("select NClassName From News_NClass where NClassID="&NClassID&"")
if rs_nc.eof and rs_nc.bof then
    call WriteErrMsg(errmsg)
else
    NewsNClassName=rs_nc(0)
end if
rs_nc.close
set rs_nc=nothing
dim sfilename,strUnit
strUnit="条新闻"
sfilename="Newslist.asp?NClassID="&NClassID&""
%>
<html>
<head>
<title>诚信物流网-行业资讯-<%=NewsNClassName%></title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<!--#include file="../inc/top.htm"-->
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td bgcolor=#FFFFFF height=5></td>
  </tr>
  </tbody> 
</table>
<table width=778 height=% 
      border=0 align="center" cellpadding=0 cellspacing=1 bgcolor=#B4B4B4>
  <tbody> 
  <tr> 
    <td height="" align="center" valign="top" bgcolor=#FFFFFF> 
      <table width="100%" height="25" 
border="0" align="center" cellpadding="0" cellspacing="0" background="">
        <tbody> 
        <tr> 
          <td width="121" height="25" align="center" valign="middle"><img src="images/zixun.gif" width="100" height="25" /></td>
          <td width="655" align="left" valign="middle"> 
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
              <tbody> 
              <tr> 
                <td class="a01" align="left" height="22"> 
                  <p><font color="#FFFFFF"><font color="#FFFFFF"><font color="#FFFFFF"><a href="index.asp" class="14link"> 
                    资讯首页</a> <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=1" class="14link">物流动态</a></font></font></font> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=2" class="14link">政策法规</a> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=3" class="14link">物流论文</a> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=4" class="14link">物流案例</a> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=5" class="14link">物流知识</a> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=6" class="14link">分析评论</a> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=7" class="14link">网站动态</a> 
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td bgcolor=#FFFFFF height=5></td>
  </tr>
  </tbody> 
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" background="images/shutiao.gif">
  <tr> 
    <td height="28"> 
      <table width="95%" border="0" align="right" cellpadding="0" cellspacing="0">
        <form name="form1" method="get" action="search.asp">
          <tr> 
            <td width="350" align="right"> 
              <div align="center"><font color="#FFFFFF"> 　　　　　　　　　　　　　　　　　　　　　　　全站搜索:</font></div>
            </td>
            <td width="213"> 
              <input name="key" type="text" value="请输入新闻关键词" size="30">
            </td>
            <td width="96"> 
              <div align="center"> 
                <select name="select">
                  <option value="物流资讯" selected>物流资讯</option>
                  <option value="公司库">公司库</option>
                  <option value="物流专线">物流专线</option>
                  <option value="物流社区">物流社区</option>
                  <option value="车源信息">车源信息</option>
                  <option value="货源信息">货源信息</option>
                </select>
              </div>
            </td>
            <td width="80" align="center"> 
              <input type="submit" name="Submit" value="搜索">
            </td>
          </tr>
        </form>
      </table>
    </td>
  </tr>
</table>
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td bgcolor=#FFFFFF height=5></td>
  </tr>
  </tbody> 
</table>
<TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>
  <TBODY> 
  <TR>
      <TD vAlign=top width=217 background=images/left_bg.jpg rowSpan=2> 
        <TABLE cellSpacing=0 cellPadding=0 width=217 bgColor=#ffffff border=0>
        <TBODY>
        <TR>
          <TD><IMG height=27 src="images/index_left2_r1_c1.gif" 
          width=30></TD>
          <TD vAlign=bottom align=middle width=208 
          background=images/index_left2_r1_c2.gif>
            <TABLE height=23 cellSpacing=0 cellPadding=0 width="100%" 
              border=0><TBODY>
              <TR>
                <TD>
                  <DIV class=yin2 align=left>每周关注</DIV></TD></TR></TBODY></TABLE></TD>
          <TD>
            <DIV align=right><IMG height=27 
            src="images/index_left2_r1_c5.gif" 
      width=25></DIV></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>
            <TABLE cellSpacing=0 cellPadding=0 width="94%" align=center 
border=0>
              <TBODY>
              <TR>
                <TD vAlign=top align=left height=5></TD></TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="95%" align=center 
border=0>
              <TBODY>
              <TR>
                <TD background=images/newsbill38.gif height=23> <%=News_left2("Recommend",6,28,0)%></TD>
              </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="94%" border=0>
        <TBODY>
        <TR>
          <TD vAlign=top align=left height=5></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=217 bgColor=#ffffff border=0>
        <TBODY>
        <TR>
          <TD><IMG height=27 src="images/index_left2_r1_c1.gif" 
          width=30></TD>
          <TD vAlign=bottom align=middle width=208 
          background=images/index_left2_r1_c2.gif>
            <TABLE height=23 cellSpacing=0 cellPadding=0 width="100%" 
              border=0><TBODY>
              <TR>
                <TD>
                  <DIV class=yin2 align=left>阅读排行</DIV></TD></TR></TBODY></TABLE></TD>
          <TD>
            <DIV align=right><IMG height=27 
            src="images/index_left2_r1_c5.gif" 
      width=25></DIV></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>
            <TABLE cellSpacing=0 cellPadding=0 width="94%" border=0>
              <TBODY>
              <TR>
                <TD vAlign=top align=left height=5></TD></TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="95%" align=center 
border=0>
              <TBODY>
              <TR>
                <TD background=images/newsbill38.gif height=23> <%=News_left2("hot",6,28,0)%></TD>
              </TR></TBODY></TABLE></TD>
        </TR></TBODY></TABLE></TD>
    <TD width=10 rowSpan=2>&nbsp;</TD>
    <TD vAlign=top>
      <TABLE cellSpacing=0 cellPadding=3 width="100%" border=0>
        <TBODY>
        <TR>
            <TD width="8%" height=5><IMG height=17 src="images/along.gif" 
            width=24></TD>
          <TD class=p14 vAlign=bottom width="66%">
            <DIV align=left><%=NewsNClassName%> 列表页</DIV></TD>
            <TD align=right width="26%">&nbsp;</TD>
          </TR>
        <TR>
          <TD bgColor=#ff7300 colSpan=3 height=3></TD></TR></TBODY></TABLE><BR>
		  <%
dim rssum,maxpage,thepages,viewpage,i
maxpage=30	
	
	Sql="select * from NewsData where NClassID = "&NClassID&" order by NewsID desc"
	Set Rs=server.createobject("Adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	    response.write "<p align=""center"">没有内容</p>"
	else
	    rssum=rs.recordcount
        if rssum mod maxpage > 0 then
            thepages=rssum\maxpage+1
        else
            thepages=rssum\maxpage
        end if

        viewpage=trim(request("page"))
        if viewpage="" or isnull(viewpage) or not(isnumeric(viewpage)) then
            viewpage=1
        elseif cint(viewpage)<1 then
            viewpage=1
        else
            if cint(viewpage)>thepages then
                viewpage=1
            end if
       end if
       i=1
       if viewpage<>"1" then
           rs.move (viewpage-1)*maxpage
       end if
	   %>
	   
      <table width="100%" border="0" align="center" cellspacing="0" class="border" cellpadding="0">
        <%
      for i=1 to maxpage
           if rs.eof then exit for
     if i mod 2=1 then
    bgcolor="#f7f7f7"
 else
    bgcolor="#ffffff"
 end if
 %>
        <tr bgcolor="<%=bgcolor%>">
          <td height="24">&nbsp;<IMG height=10 hspace=5 src="images/arrow.gif" width=10>&nbsp;<a href="NewsDetail.asp?NewsID=<%=rs("NewsID")%>" target="_blank"><%=rs("Title")%></a>&nbsp;<SPAN 
                  class=grey><font color="#999999">(<%=rs("CopyFrom")%>,<%=Month(rs("Regtime"))%>月<%=Day(rs("Regtime"))%>日)</font></SPAN></td>
      </tr>
  <%
           rs.movenext
       next
%>
</table>
<%
call showpage (strUnit,sfilename,10,"#ff0000")
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

        
		</TD>
   </TR>
  <TR>
    <TD vAlign=top>&nbsp; </TD>
  </TR></TBODY></TABLE>
<!--#include file="../inc/down.htm"-->
</body>
</html>
