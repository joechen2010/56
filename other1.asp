<table width=778 border=0 align=center cellpadding=0 cellspacing=1 >
<tr>
<td valign="top">
      <table width=100% border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600" height="180">
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600">
		  <font color="#FFFFFF"><strong>���˵�װ</strong></font></td>
          <td width="204" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF">
			<a href="banyun/index.asp" target="_blank"><font color="#FFFFFF">����</font></a></font></div>          
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
                <td colspan="3" valign="top">��ʱ����Ϣ </td>
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
	  	  <%'\\\\\\\\\\\\\\\\\\\\\\\\\\\\�����������\\\\\\\\\\\\%>
	  
    <td valign="top"> 
      <table width=100% border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600" height="180">
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>�����������</strong></font></td>
          <td width="204" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF"><a href="Repairs/index.asp" target="_blank"><font color="#FFFFFF">����</font></a></font></div>          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF"> 
            <%set file_info=server.CreateObject("adodb.recordset")
		    file_info.open "select top 6 * from repair_info order by infoid desc",conn,1,1
		  %>
            <table width="100%" border="0" cellspacing="2" cellpadding="1">
              <%if file_info.bof and file_info.eof then%>
              <tr> 
                <td colspan="3" valign="top">��ʱ����Ϣ </td>
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
	  <%'*************************��������*********************%>
	  <td valign="top">
	  <table width=100% border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600" height="180">
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600">
		  <font color="#FFFFFF"><strong>��������</strong></font></td>
          <td width="204" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF"><a href="Repairs/index.asp" target="_blank"></a></font></div>          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF"> 
            <table width="100%" height="80" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hntrans.com:81/gs/xinlichsele.php" target="_blank" class="blacklink">��̲�ѯ</a>]</td>
                <td height="22" align="center">��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.mapbar.com:8000/commonmapright/?title=�㽭������&cityCode=0571&width=760&zm=20&logo=true&cityOption=ture&mapstyle=0&state=map666&advscrolling=false&color=imgblue&headType=lydt&headPic=false" target="_blank" class="blacklink">������ͼ</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://cb.kingsoft.com/" target="_blank" class="blacklink">Ӣ������</a>]</td>
                <td height="22" align="center">��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://zj.183.com.cn/quire/code.htm" target="_blank" class="blacklink">�ʱ��ѯ</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hlbr-l-tax.gov.cn/xxfw/slcx/001%5B1%5D.htm" target="_blank" class="blacklink">˰�ʲ�ѯ]</a></td>
                <td height="22" align="center">��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://info./jinchu/whhx/default.shtml" target="_blank">������</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hzti.com/" target="_blank" class="blacklink">��ͨ״��</a>]</td>
                <td height="22" align="center"> ��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.xichang.tv/calendar.htm" target="_blank" class="blacklink">��������</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.5l.com.cn/second/shouce/kgdm.htm" target="_blank">���ʿո�</a>]</td>
                <td height="22" align="center">��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.5l.com.cn/second/shouce/gonglushoufei.htm" target="_blank" class="blacklink">��·�շ�</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="tool/sjhb.htm" target="_blank">�������</a>]</td>
                <td height="22" align="center"> ��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://info./entr/list_cgs.shtml" target="_blank">����׷��</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="tg/bgdm.asp" target="_blank">���ش���</a>]</td>
                <td height="22" align="center">��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.chemaid.com/sample/user.asp" target="_blank" class="blacklink">Σ��Ʒ��</a>]</td>
              </tr>
            </table>		  
		  </td>
        </tr>
      </table>	 </td>	  
  </tr>
  <tr>
  <%'**********************��������***********************%>
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
                <td width="89%"><font color="#FFFFFF">������̬</font><a href="user/login.asp" target="_blank"></a></td>
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
                <td width="89%"><font color="#FFFFFF">��������</font></td>
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
  <%'******************������̬******************%>
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
                      <td width="89%"><font color="#FFFFFF">���߷���</font><a href="user/login.asp" target="_blank"></a></td>
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
  <%'***********************����֪ʶ******************%>
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
                      <td width="89%"><font color="#FFFFFF">��������</font><a href="user/login.asp" target="_blank"></a></td>
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
<%'**********************���߷���*****************%>  
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
                      <td width="89%"><font color="#FFFFFF">����֪ʶ</font></td>
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
	<%'**************��������***************%>
    <td>
<table cellspacing=0 cellpadding=0 width=100% border=0>
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        
                  <div align="center"><img src="images/jiantou.gif" width="20" height="20"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">��������</font></td>
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

