<script language="javascript" src="Inc/rolling.js"></script>
 <table cellspacing=0 cellpadding=0 width="100%" border=0>
          <tbody> 
          <tr> 
            <td class=white12 valign=center bgcolor=#ff6600 height=22> 
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="9%"> 
                    <div align=center><img height=20 
                        src="images/huo.gif" width=21></div>
                  </td>
                  <td width="40%"><font color=#ffffff>最新会员信息</font></td>
                  <td width="51%"></td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr align=middle> 
            <td 
                style="BORDER-RIGHT: #ff9966 1px solid; BORDER-LEFT: #ff9966 1px solid; BORDER-BOTTOM: #ff9966 1px solid" 
                valign=top height=140> 
              <table height=67 cellspacing=0 cellpadding=0 width="100%" 
                  border=0>
                <tbody> 
                <tr> 
                  <td  valign=top height=560>
                <% dim province3,city3
        if request("province")<>"省份" then
	    province3=" and province='"&request("province")&"'"
	    end if		
		if request("city")<>"地级市" then
		city3="and city='"&request("city")&"'"
		end if	 
		if request("type")<>"" then
		qylb="and qylb='"&request("type")&"'"
		end if	 
				
				set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select * from qyml where 1=1"&province3&city3&qylb&" order by id desc",conn,1,1
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
                      <tr align="center"> 
                        <td><%=rs_info("qylb")%></td>
				        <td><%=rs_info("city")%></td>
				        <td><%=rs_info("qymc")%></td>
				        <td><%=rs_info("name")%></td>
				        <td><%=rs_info("phone2")%>-<%=rs_info("phone3")%></td>
				        <td><A href="yellowpage/detail.asp?InfoID=<%=rs_info("ID")%>" target="_blank">详情</A></td>
                </tr> 
				<%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>				
              </table>
			   <%end if%>					  
				  </td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="61%" height=24> 
                  </td>
                  <td width="39%" height=24 align="center"><a href="yellowpage/index.asp">更多&gt;&gt;</a></td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          </tbody> 
        </table>