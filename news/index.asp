<!--#include file="../Inc/Conn.asp"-->
<!--#include file="../Inc/Function.asp"-->
<!--#include file="News_Code.asp"-->
<HTML>
<HEAD>
<TITLE>诚信物流网·行业资讯</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<STYLE type=text/css>.bg {
	BACKGROUND-POSITION: 50% top; BACKGROUND-REPEAT: no-repeat
}
</STYLE></HEAD>
<BODY leftmargin="0" topmargin="0">
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
  <TR> 
    <TD vAlign=top width=250> <img src="images/1750.jpg" width="256" height="190"> 
      <TABLE width=258 border=0 cellPadding=0 
      cellSpacing=0 
      style="BORDER-RIGHT: #959595 1px solid; BORDER-TOP: #959595 1px solid; MARGIN-TOP: 1px; BORDER-LEFT: #959595 1px solid; BORDER-BOTTOM: #959595 1px solid">
        <TBODY> 
        <TR bgColor=#959595> 
          <TD align=middle height=22><font color="#FFFFFF">推荐新闻排行</font></TD>
        </TR>
        <TR> 
          <TD 
          vAlign=top bgColor=#FFFFFF class=l15 
          style="PADDING-LEFT: 1px; PADDING-BOTTOM: 1px; PADDING-TOP: 1px"><%=News_left2("Recommend",5,30,0)%> 
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
      
    </TD>
    <TD vAlign=top> 
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY> 
        <TR> 
          <TD> 
            <TABLE cellSpacing=0 cellPadding=0 width=99% align=center border=0>
              <TBODY> 
              <TR vAlign=top> 
                <TD> 
                  <table cellspacing=0 cellpadding=0 width="99%" border=0>
                    <tbody> 
                    <tr bgcolor=#FFCC00> 
                      <td height=3 colspan=3></td>
                    </tr>
                    <tr> 
                      <td width="10%" bgcolor=#FF9900 height=21>&nbsp;</td>
                      <td class=s_ffffff_14px width="72%" bgcolor=#FF9900><font color="#FFFFFF">物流动态 
                        </font></td>
                      <td class=s_ffffff_14px width="18%" bgcolor=#FF9900> 
                        <div class=white_12px align=right> 
                          <div align=center><a href="Newslist.asp?NClassID=1"><font color=#ffffff>更多</font></a></div>
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan=3> 
                        <table cellspacing=1 cellpadding=0 width="100%" bgcolor=#FF9900 
            border=0>
                          <tbody> 
                          <tr> 
                            <td bgcolor=#efffff> 
                              <table cellspacing=0 cellpadding=0 width="100%" 
                  border=0>
                                <tbody> 
                                <tr> 
                                  <td valign=top> <%= News_list(1,6,30)%> </td>
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
                </TD>
                <TD align=right> 
                  <TABLE cellSpacing=0 cellPadding=0 width="99%" border=0>
                    <TBODY> 
                    <TR bgColor=#FFCC00> 
                      <TD height=3 colSpan=3></TD>
                    </TR>
                    <TR> 
                      <TD width="10%" bgColor=#FF9900 height=21>&nbsp;</TD>
                      <TD class=s_ffffff_14px width="72%" bgColor=#FF9900><SPAN 
                        class=white><font color="#ffffff">政策法规 </font></SPAN></TD>
                      <TD class=s_ffffff_14px width="18%" bgColor=#FF9900> 
                        <DIV class=white_12px align=right> 
                          <DIV align=center><a href="Newslist.asp?NClassID=2"><FONT  color=#ffffff>更多</FONT></a></DIV>
                        </DIV>
                      </TD>
                    </TR>
                    <TR> 
                      <TD colSpan=3> 
                        <TABLE cellSpacing=1 cellPadding=0 width="100%" bgColor=#FF9900 
            border=0>
                          <TBODY> 
                          <TR> 
                            <TD align="left" valign="top" bgColor=#efffff> <%= News_list(2,6,30)%> 
                            </TD>
                          </TR>
                          </TBODY> 
                        </TABLE>
                      </TD>
                    </TR>
                    </TBODY> 
                  </TABLE>
                </TD>
              </TR>
              </TBODY> 
            </TABLE>
            <table cellspacing=0 cellpadding=0 width="100%" align=center 
            border=0>
              <tbody> 
              <tr> 
                <td height=8></td>
              </tr>
              </tbody> 
            </table>
            <table cellspacing=0 cellpadding=0 width=99% align=center border=0>
              <tbody> 
              <tr valign=top> 
                <td> 
                  <table cellspacing=0 cellpadding=0 width="99%" border=0>
                    <tbody> 
                    <tr bgcolor=#FFCC00> 
                      <td height=3 colspan=3></td>
                    </tr>
                    <tr> 
                      <td width="10%" bgcolor=#FF9900 height=21>&nbsp;</td>
                      <td class=s_ffffff_14px width="72%" bgcolor=#FF9900><font color="#FFFFFF">物流论文</font></td>
                      <td class=s_ffffff_14px width="18%" bgcolor=#FF9900> 
                        <div class=white_12px align=right> 
                          <div align=center><a href="Newslist.asp?NClassID=3"><font color=#ffffff>更多</font></a></div>
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan=3> 
                        <table cellspacing=1 cellpadding=0 width="100%" bgcolor=#FF9900 
            border=0>
                          <tbody> 
                          <tr> 
                            <td bgcolor=#efffff> 
                              <table cellspacing=0 cellpadding=0 width="100%" 
                  border=0>
                                <tbody> 
                                <tr> 
                                  <td valign=top><%= News_list(3,6,30)%> </td>
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
                </td>
                <td align=right> 
                  <table cellspacing=0 cellpadding=0 width="99%" border=0>
                    <tbody> 
                    <tr bgcolor=#FFCC00> 
                      <td height=3 colspan=3></td>
                    </tr>
                    <tr> 
                      <td width="10%" bgcolor=#FF9900 height=21>&nbsp;</td>
                      <td class=s_ffffff_14px width="72%" bgcolor=#FF9900><span 
                        class=white><font color="#ffffff">物流案例</font></span></td>
                      <td class=s_ffffff_14px width="18%" bgcolor=#FF9900> 
                        <div class=white_12px align=right> 
                          <div align=center><a href="Newslist.asp?NClassID=4"><font  color=#ffffff>更多</font></a></div>
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan=3> 
                        <table cellspacing=1 cellpadding=0 width="100%" bgcolor=#FF9900 
            border=0>
                          <tbody> 
                          <tr> 
                            <td align="left" valign="top" bgcolor=#efffff><%= News_list(4,6,30)%> 
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
            
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
    </TD>
  </TR>
</TABLE>
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td bgcolor=#FFFFFF height=5></td>
  </tr>
  </tbody> 
</table>
<table width="778" border="0" cellspacing="0" align="center">
  <tr> 
    <td> 
      <div align="left"> <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="778" height="70">
          <param name="movie" value="../truck/images/top6.swf">
          <param name="quality" value="high">
          <embed src="../truck/images/top6.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="778" height="70">
          </embed> 
        </object> </div>
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
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tr> 
    <td valign=top width=250> 
      <table 
      style="BORDER-RIGHT: #c61c06 1px solid; BORDER-TOP: #c61c06 1px solid; BORDER-LEFT: #c61c06 1px solid; BORDER-BOTTOM: #c61c06 1px solid" 
      cellspacing=0 cellpadding=0 width=258 border=0>
        <tbody> 
        <tr bgcolor=#c61c06> 
          <td width=257 height=22 align=middle bgcolor="#c61c06"><font color="#FFFFFF">热门新闻排行</font></td>
          <td style="PADDING-TOP: 3px" align=middle width=1></td>
        </tr>
        <tr> 
          <td align=middle bgcolor=#f0f0f0 colspan=2> 
            <table cellspacing=0 cellpadding=0 width="258" align=center 
                  border=0>
              <tbody> 
              <tr bgcolor=#f8f8f8> 
                <td><%=News_left2("hot",6,30,0)%> </td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td valign=top> 
      <table cellspacing=0 cellpadding=0 width="100%" border=0>
        <tbody> 
        <tr> 
          <td> 
            <table cellspacing=0 cellpadding=0 width=99% align=center border=0>
              <tbody> 
              <tr valign=top> 
                <td> 
                  <table cellspacing=0 cellpadding=0 width="99%" border=0>
                    <tbody> 
                    <tr bgcolor=#FFCC00> 
                      <td height=3 colspan=3></td>
                    </tr>
                    <tr> 
                      <td width="10%" bgcolor=#FF9900 height=21>&nbsp;</td>
                      <td class=s_ffffff_14px width="72%" bgcolor=#FF9900><font color="#FFFFFF">物流知识 
                        </font></td>
                      <td class=s_ffffff_14px width="18%" bgcolor=#FF9900> 
                        <div class=white_12px align=right> 
                          <div align=center><a href="Newslist.asp?NClassID=5"><font color=#ffffff>更多</font></a></div>
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan=3> 
                        <table cellspacing=1 cellpadding=0 width="100%" bgcolor=#FF9900 
            border=0>
                          <tbody> 
                          <tr> 
                            <td bgcolor=#efffff> 
                              <table cellspacing=0 cellpadding=0 width="100%" 
                  border=0>
                                <tbody> 
                                <tr> 
                                  <td valign=top><%= News_list(5,6,30)%> </td>
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
                </td>
                <td align=right> 
                  <table cellspacing=0 cellpadding=0 width="99%" border=0>
                    <tbody> 
                    <tr bgcolor=#FFCC00> 
                      <td height=3 colspan=3></td>
                    </tr>
                    <tr> 
                      <td width="10%" bgcolor=#FF9900 height=21>&nbsp;</td>
                      <td class=s_ffffff_14px width="72%" bgcolor=#FF9900><span 
                        class=white><font color="#ffffff">分析评论</font></span></td>
                      <td class=s_ffffff_14px width="18%" bgcolor=#FF9900> 
                        <div class=white_12px align=right> 
                          <div align=center><a href="Newslist.asp?NClassID=6"><font  color=#ffffff>更多</font></a></div>
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan=3> 
                        <table cellspacing=1 cellpadding=0 width="100%" bgcolor=#FF9900 
            border=0>
                          <tbody> 
                          <tr> 
                            <td align="left" valign="top" bgcolor=#efffff><%= News_list(6,6,30)%> 
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
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
</table>
<!--#include file="../inc/down.htm"-->
</BODY>
</HTML>
