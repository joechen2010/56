<script language="javascript" src="Inc/rolling.js"></script>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>

<table cellspacing=0 cellpadding=0 width="100%" border=0>
    <tr> 
    <td class=white12 valign=center height=25 background="images/menu01.jpg"> 
      <table cellspacing=0 width="100%" border=0>
        <tr> 
          <td width="20"> 
            <div align=center></div>
          </td>
          <td style="FONT-SIZE: 14px"><b>最新货源信息</b></td>
          <td width="50%"> <a href="javascript:openwindow('sort.asp',600,400)">发布信息排行榜</a> 
            <font color="#FFFFFF"> | </font><a href="Member/car_manage.asp"><font color="#FFFF00"></font></a> 
            <a href="Member/goods_add.asp">免费发布货源信息</a></td>
          <td align="right"><a href="truck/index.asp">更多&gt;&gt;</a> </td>
        </tr>
      </table>
      </td>
  </tr>
          <tr align=middle> 
    <td style="BORDER-RIGHT: #FF9501 1px solid; BORDER-LEFT: #FF9501 1px solid; BORDER-BOTTOM: #FF9501 1px solid" align="center"> 
      <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 32 * from info2 where infotype='货源信息' order by infoid desc",conn,1,1
	%>
      <%if rs_info.bof and rs_info.eof then%>
      <table cellspacing=0 cellpadding=0 width="99%" border=0>
        <tr> 
                  
          <td  valign=top>无信息</td>
        </tr> 
      </table>
                <%else %>
				<%dim i%>
				<%if request("action")="" then%>
<SCRIPT language=JavaScript>
wireOpen();
</SCRIPT>		
                  <%end if%>
      <table width=99% height="22" border=0 align="center" cellpadding=0 cellspacing=0 bordercolor="#CCFF99" bgcolor="#999999">
        <%do while not rs_info.eof%>
        <tr <%if i mod 2=0 then response.Write "bgcolor=#FFEAD3" end if%> bgcolor="#FFFFFF">
          <td width="7%" align="left"  valign=top> 
            <div align="center"><img src="images/phxxh.gif" width="22" height="20" /></div>
          </td> 
          <td width="22%" align="left"  valign=top><%=rs_info("city")%>&nbsp;<%=rs_info("area")%></td>
          <td width="4%" align="center"  valign=top>→</td>
          <td width="22%" align="center"  valign=top><%=rs_info("city2")%>&nbsp;<%=rs_info("area2")%></td>
          <td width="11%" align="center"  valign=top><%=rs_info("cartype")%></td>
          <td width="12%" align="center"  valign=top><%=rs_info("carload")%>吨</td> 
          <td width="15%" align="center"  valign=top> <%=month(rs_info("addtime"))%>-<%=day(rs_info("addtime"))%>&nbsp; 
            <%=hour(rs_info("addtime"))%>:<%=minute(rs_info("addtime"))%></td>
          <td width="7%"  align="center"><A href="javascript:openwindow('truck/detail.asp?InfoID=<%=rs_info("InfoID")%>',500,400)">详情</A></td>
        </tr> 
				<%rs_info.movenext
				i=i+1
			    loop
				rs_info.close
				set rs_info=nothing
				%>				
      </table>
			  <%if request("action")="" then%>
<SCRIPT language=JavaScript>                                                                                   
	wireClose(); 
</SCRIPT>
      <%end if%>
      <%end if%>
    </td>
          </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="5"></td>
  </tr>
</table>
<TABLE cellSpacing=0 width="100%" border=0>
          <TR> 
    <TD vAlign=top width="48%"> 
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                <TR> 
          <TD class=white12 vAlign=center background="images/menu02.jpg" height=25> 
            <TABLE cellSpacing=0 width="100%" border=0>
              <TR> 
                <TD width="20"> 
                  <DIV align=center></DIV>
                </TD>
                <TD style="FONT-SIZE: 14px"><b>最新空车信息</b></TD>
                <TD width="34%"> <font color="#FFFFFF">| </font><a href="Member/car_manage.asp"><font color="#FFFF00"></font></a> 
                  <a href="Member/car_add.asp">免费发布空车信息</a></TD>
                <TD width="35%" align="right"><font color="#FFFF00"><a 
                        href="good/index.asp">更多&gt;&gt;</a></font><font color="#FFFFFF"> 
                  </font></TD>
              </TR>
            </TABLE>
                  </TD>
                </TR>
                <TR align=middle> 
          <TD style="BORDER-RIGHT: #6BC42C 1px solid; BORDER-LEFT: #6BC42C 1px solid; BORDER-BOTTOM: #6BC42C 1px solid" align="center"> 
            <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 32 * from info2 where infotype='车源信息' order by infoid desc",conn,1,1
				%>
                  <%if rs_info.bof and rs_info.eof then%>
                  <table cellspacing=0 cellpadding=0 width="100%" border=0>
                    <tr> 
                      <td  valign=top>无信息</td>
                </tr> 
              </table>
                <%else %>
				<%dim h%>
				<%if request("action")="" then%>
<SCRIPT language=JavaScript>
wireOpen();
</SCRIPT>		
                  <%end if%>
                  <table width="99%" height="22" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
                    <%do while not rs_info.eof%>
                    <tr <%if h mod 2=0 then response.Write "bgcolor=#EFF7CE" end if%>>
                      
                <td width="7%" align="left"  valign="top">
                  <div align="center"><img src="images/phxxc.gif" width="25" height="20" /></div>
                </td>
                      
                <td width="18%" align="left"  valign="top"><%=rs_info("city")%>&nbsp;<%=rs_info("area")%></td>
                      
                <td width="4%" align="center">→</td>
                      
                <td width="26%" align="center"><%=rs_info("city2")%>&nbsp;<%=rs_info("area2")%></td>
                      
                <td width="11%" align="center"><%=rs_info("cartype")%></td>
                      
                <td width="12%" align="center"><%=rs_info("carload")%>吨</td>
                      
                <td width="15%" align="center"><%=month(rs_info("addtime"))%>-<%=day(rs_info("addtime"))%>&nbsp; 
                  <%=hour(rs_info("addtime"))%>:<%=minute(rs_info("addtime"))%> 
                </td>
                      
                <td width="7%" align="center"><a href="javascript:openwindow('truck/detail.asp?InfoID=<%=rs_info("InfoID")%>',500,400)">详情</a></td>
                    </tr>
                    <%rs_info.movenext
				h=h+1
			    loop
				rs_info.close
				set rs_info=nothing
				%>
                  </table>
                  <%if request("action")="" then%>
                  <SCRIPT language=JavaScript>                                                                                   
	wireClose(); 
</SCRIPT>
                  <%end if%>
                  <%end if%>
</TD>
                </TR>
      </TABLE>
    </TD>
          </TR>
</TABLE>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="5"></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="50%"> 
      <table cellspacing=0 cellpadding=0 width="99%" border=0>
        <tr> 
          <td class=white12 valign=center height=25 background="images/menu03.jpg"> 
            <table cellspacing=0 width="100%" border=0>
               
              <tr> 
                <td width="20"> </td>
                <td style="FONT-SIZE: 14px"><strong>货运专线</strong></td>
                <td align="right"> </td>
                <td align="right"><a  href="fline/index.asp">更多&gt;&gt;</a></td>
              </tr>
            </table>          </td>
        </tr>
        <tr align=middle> 
          <td style="BORDER-RIGHT: #43A6DD 1px solid; BORDER-LEFT: #43A6DD 1px solid; BORDER-BOTTOM: #43A6DD 1px solid" valign=top> 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
              <tr> 
                <td  valign=top height=100> 
                  <%set rs_h=server.CreateObject("adodb.recordset")
					 rs_h.open "select top 16 * from zx_info where infotype='货运专线' order by infoid desc",conn,1,1
			  %>
                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <%if rs_h.bof and rs_h.eof then%>
                    <tr> 
                      <td height=23 align=center class=1>暂时无消息</td>
                    </tr>
                    <%else%>
                    <%j=0%>
                    <tr> 
                      <%do while not rs_h.eof%>
                      <td align=center class=1> <a href="javascript:openwindow('fline/detail.asp?infoid=<%=rs_h("infoid")%>',500,400)"> 
                        <%=rs_h("city")%>←→<%=rs_h("city2")%></a></td>
                      <%rs_h.movenext
			   j=j+1
			   if j mod 4=0 then
			   response.Write "</tr><tr>"
			   end if
			    loop
				rs_h.close
				set rs_h=nothing
			  %>
                      <%end if%>
                  </table>                </td>
              </tr>
            </table>          </td>
        </tr>
      </table>    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="5"></td>
  </tr>
</table>
<table cellspacing=0 cellpadding=0 width="100%" border=0>
  <tr> 
    <td class=white12 valign=center background="images/menu04.jpg" height=25>
      <table cellspacing=0 width="100%" border=0>
        <tr> 
          <td width="20"> 
            <div align=center></div>
          </td>
          <td style="FONT-SIZE: 14px"><b>最新会员信息</b></td>
          <td width="51%" align="right"><a href="yellowpage/index.asp">更多&gt;&gt;</a></td>
        </tr>
      </table>
    </td>
  </tr>
          <tr align=middle> 
            <td style="BORDER-RIGHT: #EF219D 1px solid; BORDER-LEFT: #EF219D 1px solid; BORDER-BOTTOM: #EF219D 1px solid" valign=top height=140> 
              <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0>
                <tr> 
                  <td  valign=top height=67>
                <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 6 * from qyml order by id desc",conn,1,1
				%>	
				<%if rs_info.bof and rs_info.eof then%>			
              <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0> 
                <tr> 
                  <td  valign=top height=67>无信息</td>
                </tr> 
              </table>
                <%else %>
				<%if request("action")="" then%>
                  <%end if%>
            <table cellspacing=2 cellpadding=0 border=0 width=100%>
              <%do while not rs_info.eof%>
              <tr align="center"> 
                <td><span class="STYLE1"><%=rs_info("qylb")%></span></td>
                <td><%=rs_info("city")%>市</td>
                <td> 
                  <%if len(rs_info("qymc"))>10 then%>
                  <a href="site/aboutus.asp?InfoID=<%=rs_info("ID")%>&user=<%=rs_info("user")%>" target="_blank"><%=left(rs_info("qymc"),8)%>...</a> 
                  <%else%>
                  <a href="site/aboutus.asp?InfoID=<%=rs_info("ID")%>&user=<%=rs_info("user")%>" target="_blank"><%=rs_info("qymc")%></a> 
                  <%end if%>
                </td>
                <td><%=rs_info("name")%></td>
                <td><%=rs_info("phone2")%>-<%=rs_info("phone3")%></td>
              </tr>
              <%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>
            </table>
			  <%if request("action")="" then%>
               <%end if%> 
			   <%end if%>					  
				  </td>
                </tr>
              </table>
    </td>
          </tr>
</table>