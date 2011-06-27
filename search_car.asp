       <TABLE cellSpacing=0 width="100%" border=0>
          <TBODY> 
          <TR> 
            <TD vAlign=top width="48%" height=151> 
              <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                <TBODY> 
                <TR> 
                  <TD class=white12 vAlign=center bgColor=#ff6600 height=22> 
                    <TABLE cellSpacing=0 width="100%" border=0>
                      <TBODY> 
                      <TR> 
                        <TD width="9%"> 
                          <DIV align=center><IMG height=20 
                        src="images/huo.gif" width=21></DIV>
                        </TD>
                        
                      <TD width="22%"><FONT color=#ffffff>最新空车信息</FONT></TD>
                        
                      <TD width="69%"> <font color="#FFFF00">发布信息排行榜</font><font color="#FFFFFF"> 
                        | </font><a href="Member/car_manage.asp"><font color="#FFFF00"></font></a> 
                        <a href="Member/car_add.asp"><font color="#FFFF00">免费发布空车信息</font></a><font color="#FFFFFF"> 
                        | </font><a href="Member/car_manage.asp"><font color="#FFFF00">管理车源信息</font></a></TD>
                      </TR>
                      </TBODY> 
                    </TABLE>
                  </TD>
                </TR>
                <TR align=middle> 
                  <TD 
                style="BORDER-RIGHT: #ff9966 1px solid; BORDER-LEFT: #ff9966 1px solid; BORDER-BOTTOM: #ff9966 1px solid" 
                vAlign=top height=140> 
                    <TABLE height=67 cellSpacing=0 cellPadding=0 width="100%" 
                  border=0>
                      <TBODY> 
                      <TR> 
                        <TD  vAlign=top height=560>
                <%  dim province2,city2
        if request("province")<>"省份" then
	    province2=" and province='"&request("province")&"'"
	    end if		
		if request("city")<>"地级市" then
		city2="and city='"&request("city")&"'"
		end if					
				set rs_info=server.CreateObject("adodb.recordset")
			rs_info.open "select * from info2 where 1=1"&province2&city2&" and infotype='车源信息' order by infoid desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='供应' Order by InfoID desc")
				%>	
				<%if rs_info.bof and rs_info.eof then%>			
              <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0> 
                <tr> 
                  <td  valign=top height=67>无信息</td>
                </tr> 
              </table>
                <%else %>
			
                        <table cellspacing=2 cellpadding=2 border=0 width=100%>
                          <%do while not rs_info.eof%>
                          <tr> 
                            <td  valign=top align="left"><%=rs_info("city")%>&nbsp;<%=rs_info("area")%></td>
				            <td align="center">→</td>
				            <td align="center"><%=rs_info("city2")%>&nbsp;<%=rs_info("area2")%></td>
				            <td align="center"><%=rs_info("cartype")%></td>
				            <td align="center"><%=rs_info("carload")%>吨</td>
				            <td align="center"><%=rs_info("addtime")%></td>
				            <td align="center"><A href="javascript:openwindow('truck/detail.asp?InfoID=<%=rs_info("InfoID")%>',500,400)">详情</A></td>
                </tr> 
				<%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>				
              </table>

			   <%end if%>						
					    </TD>
                      </TR>
                      </TBODY> 
                    </TABLE>
                    <TABLE cellSpacing=0 width="100%" border=0>
                      <TBODY> 
                      <TR> 
                        <TD width="61%" height=24>&nbsp;</TD>
                        <TD width="39%" height=24> <A href="good/index.asp">更多&gt;&gt;</A>
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