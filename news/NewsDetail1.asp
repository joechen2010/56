<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="News_Code.asp"-->
<%
dim NewsID
NewsID = SafeRequest("NewsID",1)
if NewsID="" then
       call WriteErrMsg(errmsg)
end if
conn.execute("update NewsData set hits=hits+1 where NewsID="&NewsID&"")
Dim Rs,Sql,sNewsID,sTitle,sContent
set Rs=server.createobject("Adodb.Recordset")
    Sql="select NewsID,Title,NClassID,(Select NClassName from News_NClass where NClassID=A.NClassID) as NClassName,Content,Picture,CopyFrom,IncludePic,Hits,RegTime from NewsData A where NewsID="&NewsID&""
	 Rs.open Sql,conn,1,1
	 if not Rs.eof then
	    sNewsID = Rs(0)
		sTitle = Rs(1)
		sNClassID = Rs(2)
		sNClassName = Rs(3)
		sContent = Rs(4)
		sPicture = Rs(5)
	    sCopyFrom = Rs(6)
		sIncludePic = Rs(7)
		sHits = Rs(8)
		sRegTime = Rs(9)
	 else
	    call WriteErrMsg(errmsg)
	 end if
	 Rs.close
	 set Rs = Nothing
%>
<html>
<head>
<title>诚信物流网--行业资讯--<%=sTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/css.css" type=text/css rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0">
<table width="480" border="0" cellpadding="2" cellspacing="2" align="center">
  <tr> 
    <td align="center" height="40"><embed src="../flash/banner.swf" width="480" height="60" type="application/x-shockwave-flash">
      </embed></td>
  </tr>
  <tr> 
    <td align="center" height="40"> <span class="bt24"><b><font size="4"><%=sTitle%></font></b></span> 
    </td>
  </tr>
  <tr> 
    <td class="title"> 
      <hr size="1">
    </td>
  </tr>
  <tr> 
    <td align="center" height="40">来源：<font color="#990000"><%=sCopyFrom%></font>&nbsp;发布时间：<font color="#990000"><%=sRegTime%></font></td>
  </tr>
  <tr> 
    <td class="Content"> <span class=wz14><%=sContent%></span> </td>
  </tr>
</table>
</body>
</html>