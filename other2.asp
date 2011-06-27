
<table width=777 border=0 align=center cellpadding=0 cellspacing=1 >
  <tr> 
    <td valign="top" width="33%">
      <table width=99% border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#C2C2C2" height="180">
        <tr> 
          <td colspan="3" height="20" valign="middle" align="center"> 
            <table height=22 cellspacing=0 cellpadding=0 width="100%" 
              border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#296CBD>&nbsp;&nbsp;<img height=9 
                  src="images/hy2.gif" width=9>&nbsp;<span class="style14"><font color="#FFFFFF"><strong><b>搬运吊装</b></strong></font></span></td>
                <td width=35 height="23" bgcolor=#296CBD><font color="#FFFFFF"><a href="banyun/index.asp" target="_blank"><font color="#FFFFFF">更多</font></a></font></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        <tr> 
          <td colspan="3" valign="top" bgcolor="#FFFFFF"> 
            <%set file_info=server.CreateObject("adodb.recordset")
		    file_info.open "select top 6 * from banyun_info where province='"&fz&"' or city='"&fz&"' order by infoid desc",conn,1,1
		  %>
            <table width="100%" border="0" cellspacing="2" cellpadding="1">
              <%if file_info.bof and file_info.eof then%>
              <tr> 
                <td colspan="3" valign="top">暂时无信息 </td>
              </tr>
              <%else%>
              <%do while not file_info.eof%>
              <tr align="center"> 
                <td width="68" valign="top" align="center"> 
                  <%title=gotTopic(file_info("title"),8)%>
                  <A href="javascript:openwindow('banyun/detail.asp?infoid=<%=file_info("infoid")%>',500,400)"> 
                  <%=title%></a> </td>
                <td width="84" valign="top">&nbsp;<%=file_info("city")%>&nbsp;<%=file_info("area")%></td>
                <td width="77" valign="top"> <%=month(file_info("addtime"))%>-<%=day(file_info("addtime"))%>&nbsp; 
                  <%=hour(file_info("addtime"))%>:<%=minute(file_info("addtime"))%>	
                </td>
              </tr>
              <%file_info.movenext
			    loop
				file_info.close
				set file_info=nothing
			  %>
              <%end if%>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <%'\\\\\\\\\\\\\\\\\\\\\\\\\\\\最新修理配件\\\\\\\\\\\\%>
    <td valign="top"> 
      <table width=99% border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#C2C2C2" height="180">
        <tr bgcolor="#296CBD"> 
          <td colspan="3" height="20" valign="middle" align="center"> 
            <table height=22 cellspacing=0 cellpadding=0 width="100%" 
              border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#296CBD>&nbsp;&nbsp;<img height=9 
                  src="images/hy2.gif" width=9>&nbsp;<span class="style14"><font color="#FFFFFF"><strong><b>最新修理配件</b></strong></font></span></td>
                <td width=35 height="23" bgcolor=#296CBD><font color="#FFFFFF"><a href="Repairs/index.asp" target="_blank"><font color="#FFFFFF">更多</font></a><a href="banyun/index.asp" target="_blank"></a></font></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        <tr> 
          <td colspan="3" valign="top" bgcolor="#FFFFFF"> 
            <%set file_info=server.CreateObject("adodb.recordset")
		    file_info.open "select top 6 * from repair_info where province='"&fz&"' or city='"&fz&"' order by infoid desc",conn,1,1
		  %>
            <table width="100%" border="0" cellspacing="2" cellpadding="1">
              <%if file_info.bof and file_info.eof then%>
              <tr> 
                <td colspan="3" valign="top">暂时无信息 </td>
              </tr>
              <%else%>
              <%do while not file_info.eof%>
              <tr align="center"> 
                <td width="65" valign="top" align="center"> 
                  <%title=gotTopic(file_info("title"),8)%>
                  <A href="javascript:openwindow('repairs/detail.asp?infoid=<%=file_info("infoid")%>',500,400)"> 
                  <%=title%></a> </td>
                <td width="89" valign="top">&nbsp;<%=file_info("city")%>&nbsp;<%=file_info("area")%></td>
                <td width="85" valign="top"> <%=month(file_info("addtime"))%>-<%=day(file_info("addtime"))%>&nbsp; 
                  <%=hour(file_info("addtime"))%>:<%=minute(file_info("addtime"))%>	
                </td>
              </tr>
              <%file_info.movenext
			    loop
				file_info.close
				set file_info=nothing
			  %>
              <%end if%>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td valign="top" width="33%"> 
      <table width=99% border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#C2C2C2" height="180">
        <tr> 
          <td colspan="3" height="20" valign="middle" align="center"> 
            <table height=22 cellspacing=0 cellpadding=0 width="100%" 
              border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#296CBD>&nbsp;&nbsp;<img height=9 
                  src="images/hy2.gif" width=9>&nbsp;<span class="style14"><font color="#FFFFFF"><strong><b>物流服务</b></strong></font></span></td>
                <td width=35 height="23" bgcolor=#296CBD><font color="#FFFFFF"><a href="Repairs/index.asp" target="_blank"></a><a href="banyun/index.asp" target="_blank"></a></font></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        <tr> 
          <td colspan="3" valign="top" bgcolor="#FFFFFF"> 
            <table width="100%" height="80" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hntrans.com:81/gs/xinlichsele.php" target="_blank" class="blacklink">里程查询</a>]</td>
                <td height="22" align="center">　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.mapbar.com:8000/commonmapright/?title=浙江物流网&cityCode=0571&width=760&zm=20&logo=true&cityOption=ture&mapstyle=0&state=map666&advscrolling=false&color=imgblue&headType=lydt&headPic=false" target="_blank" class="blacklink">物流地图</a>]</td>
              </tr>
              <tr> 
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://cb.kingsoft.com/" target="_blank" class="blacklink">英汉工具</a>]</td>
                <td height="22" align="center">　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://zj.183.com.cn/quire/code.htm" target="_blank" class="blacklink">邮编查询</a>]</td>
              </tr>
              <tr> 
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hlbr-l-tax.gov.cn/xxfw/slcx/001%5B1%5D.htm" target="_blank" class="blacklink">税率查询]</a></td>
                <td height="22" align="center">　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://info./jinchu/whhx/default.shtml" target="_blank">结汇核销</a>]</td>
              </tr>
              <tr> 
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hzti.com/" target="_blank" class="blacklink">交通状况</a>]</td>
                <td height="22" align="center"> 　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.xichang.tv/calendar.htm" target="_blank" class="blacklink">电子日历</a>]</td>
              </tr>
              <tr> 
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.5l.com.cn/second/shouce/kgdm.htm" target="_blank">国际空港</a>]</td>
                <td height="22" align="center">　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.5l.com.cn/second/shouce/gonglushoufei.htm" target="_blank" class="blacklink">公路收费</a>]</td>
              </tr>
              <tr> 
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="tool/sjhb.htm" target="_blank">世界货币</a>]</td>
                <td height="22" align="center"> 　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://info./entr/list_cgs.shtml" target="_blank">货物追踪</a>]</td>
              </tr>
              <tr> 
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="tg/bgdm.asp" target="_blank">报关代码</a>]</td>
                <td height="22" align="center">　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.chemaid.com/sample/user.asp" target="_blank" class="blacklink">危险品名</a>]</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <%'**********************政策法规*****************%>
</table>

