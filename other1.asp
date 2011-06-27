<table width=778 border=0 align=center cellpadding=0 cellspacing=1 >
<tr>
<td valign="top">
      <table width=100% border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600" height="180">
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600">
		  <font color="#FFFFFF"><strong>搬运吊装</strong></font></td>
          <td width="204" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF">
			<a href="banyun/index.asp" target="_blank"><font color="#FFFFFF">更多</font></a></font></div>          
			</td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF"> 
            <%set file_info=server.CreateObject("adodb.recordset")
		    file_info.open "select top 8 * from banyun_info order by infoid desc",conn,1,1
		  %>
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
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
			   <%=title%></a>			   </td>
                <td width="84" valign="top">&nbsp;<%=file_info("city")%>&nbsp;<%=file_info("area")%></td>
                <td width="77" valign="top">
						<%=month(file_info("addtime"))%>-<%=day(file_info("addtime"))%>&nbsp;
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
      <table width=100% border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600" height="180">
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>最新修理配件</strong></font></td>
          <td width="204" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF"><a href="Repairs/index.asp" target="_blank"><font color="#FFFFFF">更多</font></a></font></div>          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF"> 
            <%set file_info=server.CreateObject("adodb.recordset")
		    file_info.open "select top 6 * from repair_info order by infoid desc",conn,1,1
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
			   <%=title%></a>			   </td>
                <td width="89" valign="top">&nbsp;<%=file_info("city")%>&nbsp;<%=file_info("area")%></td>
                <td width="85" valign="top">
				<%=month(file_info("addtime"))%>-<%=day(file_info("addtime"))%>&nbsp;
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
	  <%'*************************物流服务*********************%>
	  <td valign="top">
	  <table width=100% border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600" height="180">
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600">
		  <font color="#FFFFFF"><strong>物流服务</strong></font></td>
          <td width="204" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF"><a href="Repairs/index.asp" target="_blank"></a></font></div>          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF"> 
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
      </table>	 </td>	  
  </tr>
  <tr>
  <%'**********************友情链接***********************%>
    <td rowspan="2" valign="top">
      <table cellspacing=0 cellpadding=0 width=100% border=0>
        <tbody> 
        <tr> 
          <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
            <table width="100%" border="0" cellspacing="0">
              <tr> 
                <td width="11%"> 
                  <div align="center"><img src="images/jiantou.gif" width="20" height="20"></div>
                </td>
                <td width="89%"><font color="#FFFFFF">物流动态</font><a href="user/login.asp" target="_blank"></a></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr align=middle> 
          <td valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
            <table 
      style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
      cellspacing=0 cellpadding=1 width=100% border=0>
              <tbody> 
              <%set rs_n=conn.execute("select top 5 * from NewsData where NClassID=1 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
              <tr> 
                <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No 
                  Data!</td>
              </tr>
              <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
              <tr> 
                <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a></td>
              </tr>
              <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
              </tbody> 
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
      <table cellspacing=0 cellpadding=0 width=100% border=0>
        <tr> 
          <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
            <table width="100%" border="0" cellspacing="0">
              <tr> 
                <td width="11%"> 
                  <div align="center"><img src="images/jiantou.gif" width="20" height="20"></div>
                </td>
                <td width="89%"><font color="#FFFFFF">物流案例</font></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr align=middle> 
          <td valign="top" style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" cellspacing=0 cellpadding=1 width=100% border=0>
              <%set rs_n=conn.execute("select top 5 * from NewsData where NClassID=4  Order By updown Desc")
		  if rs_n.eof and rs_n.bof then
		%>
              <tr> 
                <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No 
                  Data!</td>
              </tr>
              <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
              <tr> 
                <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a></td>
              </tr>
              <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
            </table>
          </td>
        </tr>
      </table>
    </td>
  <%'******************物流动态******************%>
  <td height="22" valign="top">
<table cellspacing=0 cellpadding=0 width=100% border=0>
              <tbody> 
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        
                  <div align="center"><img src="images/jiantou.gif" width="20" height="20"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">政策法规</font><a href="user/login.asp" target="_blank"></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                
          <td valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
            <table 
      style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
      cellspacing=0 cellpadding=1 width=100% border=0>
          <tbody> 
          <%set rs_n=conn.execute("select top 5 * from NewsData where NClassID=2 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
            <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a></td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
          </tbody> 
        </table>
        
          </td>
              </tr>
              </tbody> 
            </table>  
  </td>
  <%'***********************物流知识******************%>
  <td valign="top">
<table cellspacing=0 cellpadding=0 width=100% border=0>
              <tbody> 
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        
                  <div align="center"><img src="images/jiantou.gif" width="20" height="20"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">物流论文</font><a href="user/login.asp" target="_blank"></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                
          <td valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" cellspacing=0 cellpadding=1 width=100% border=0>
          <%set rs_n=conn.execute("select top 5 * from NewsData where NClassID=3 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
                <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a></td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
        </table>
          </td>
              </tr>
            </table>  
  </td>
  </tr>
<%'**********************政策法规*****************%>  
  <tr>
    <td height="42" valign="top">
<table cellspacing=0 cellpadding=0 width=100% border=0>
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        
                  <div align="center"><img src="images/jiantou.gif" width="20" height="20"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">物流知识</font></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                
          <td valign="top" style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" cellspacing=0 cellpadding=1 width=100% border=0>
          <%set rs_n=conn.execute("select top 5 * from NewsData where NClassID=5  Order By updown Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
                <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a></td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
        </table>
          </td>
              </tr>
            </table>	
	</td>
	<%'**************分析评论***************%>
    <td>
<table cellspacing=0 cellpadding=0 width=100% border=0>
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        
                  <div align="center"><img src="images/jiantou.gif" width="20" height="20"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">分析评论</font></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                
          <td valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" cellspacing=0 cellpadding=1 width=100% border=0>
          <%set rs_n=conn.execute("select top 5 * from NewsData where NClassID=6 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
                <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a></td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
        </table>
          </td>
              </tr>
            </table>	
	</td>
  </tr>
</table>

