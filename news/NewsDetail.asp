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
<title>����������--��ҵ��Ѷ--<%=sTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css 
rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0">
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
                    ��Ѷ��ҳ</a> <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=1" class="14link">������̬</a></font></font></font> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=2" class="14link">���߷���</a> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=3" class="14link">��������</a> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=4" class="14link">��������</a> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=5" class="14link">����֪ʶ</a> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=6" class="14link">��������</a> 
                    <font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000">|</font></font></font> 
                    <a href="Newslist.asp?NClassID=7" class="14link">��վ��̬</a> 
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
              <div align="center"><font color="#FFFFFF"> ����������������������������������������������ȫվ����:</font></div>
            </td>
            <td width="213"> 
              <input name="key" type="text" value="���������Źؼ���" size="30">
            </td>
            <td width="96"> 
              <div align="center"> 
                <select name="select">
                  <option value="������Ѷ" selected>������Ѷ</option>
                  <option value="��˾��">��˾��</option>
                  <option value="����ר��">����ר��</option>
                  <option value="��������">��������</option>
                  <option value="��Դ��Ϣ">��Դ��Ϣ</option>
                  <option value="��Դ��Ϣ">��Դ��Ϣ</option>
                </select>
              </div>
            </td>
            <td width="80" align="center"> 
              <input type="submit" name="Submit" value="����">
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
<TABLE height=10 cellSpacing=0 cellPadding=0 width=778 align=center border=0>
  <TBODY> 
  <TR>
    <TD class=S height=30>&nbsp;����ǰ��λ��: <A 
      href="index.asp">��ҳ</A>&nbsp;&gt;&gt;&nbsp;<A 
      href="NewsList.asp?NClassID=<%=sNClassID%>"><%=sNClassName%></A>&gt;&gt; 
  ����</TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=778 align=center bgColor=#003366 
border=0>
  <TBODY> 
  <TR>
    <TD height=2></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=778 align=center 
border=0>
  <tr> 
    <td valign="top" align="center" bgcolor="EFF7FF"> 
      <table width="98%" border="0" cellpadding="2" cellspacing="2" align="center">
        <tr> 
          <td align="center" HEIGHT="40">
            <span class="bt24"><b><%=sTitle%></b></span>
          </td>
        </tr>
        <tr> 
          <td class="title"> 
            <hr size="1">
          </td>
        </tr>
        <tr> 
          <td align="center" height="40">��Դ��<font color="#990000"><%=sCopyFrom%></font>&nbsp;����ʱ�䣺<font color="#990000"><%=sRegTime%></font></td>
        </tr>
        <tr> 
          <td class="Content">
            <SPAN class=wz14><%=sContent%></SPAN>
          </td>
        </tr>
      </table>
      
      <p>&nbsp;</p>
      <p><font color="#999999">����Ϣ��ʵ��δ������������֤ʵ���������ο���</font></p>
    </td>
    <td width="188" bgcolor="#f0f0f0" valign="top"> 
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD class=login align=middle 
          background="images/index_news_title_bg.gif" 
          height=28>&nbsp;���ר���Ƽ�</TD>
        </TR>
        <TR> 
          <TD><%=News_left2("Recommend",10,24,sNClassID)%> </TD>
        </TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>&nbsp;</TD></TR></TBODY></TABLE>
      <table class=table6 bordercolor=#999999 cellspacing=0 
                  cellpadding=0 width="100%">
        <tbody> 
        <tr bgcolor=#f6f6f6> 
          <td width="4%" 
                      background="images/index_meun2.jpg" 
                      height=25> 
            <div align=center></div>
          </td>
          <td width="49%" 
                      background="images/index_meun2.jpg"><span 
                        class=login>ͼƬ��棺</span></td>
          <td class=12p width="47%" 
                      background="images/index_meun2.jpg">&nbsp;</td>
        </tr>
        <tr bgcolor=#e6e6e6> 
          <td colspan=3 height=1></td>
        </tr>
        <tr bgcolor=#ffffff> 
          <td colspan=3 height=110><span class=lh15><a href="../About/keyword.htm" target="_blank"><img src="../Ads/keyword.gif" width="187" height="246" border="0"></a></span></td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
</table>
<!--#include file="../inc/down.htm"-->
</body>
</html>