      <table cellspacing=0 cellpadding=0 width="100%" border=0>
          <tr> 
            <td class=white12 valign=center bgcolor=#ff6600 height=22> 
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="9%"> 
                    <div align=center><img height=20 
                        src="images/huo.gif" width=21></div>
                  </td>
                  
                <td width="22%"><font color=#ffffff>���»�Դ��Ϣ</font></td>
                  
                <td width="69%"> <font color="#FFFF00">������Ϣ���а�</font><font color="#FFFFFF"> 
                  | </font><a href="Member/car_manage.asp"><font color="#FFFF00"></font></a> 
                  <a href="Member/goods_add.asp"><font color="#FFFF00"> ��ѷ�����Դ��Ϣ</font></a><font color="#FFFFFF"> 
                  | </font><a href="Member/goods_manage.asp"><font color="#FFFF00">�����Դ��Ϣ</font></a></td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr align=middle> 
            
    <td 
                style="BORDER-RIGHT: #ff9966 1px solid; BORDER-LEFT: #ff9966 1px solid; BORDER-BOTTOM: #ff9966 1px solid" 
                valign=top> 
      <%dim province,city
        if request("province")<>"ʡ��" then
	    province=" and province='"&request("province")&"'"
	    end if		
		if request("city")<>"�ؼ���" then
		city="and city='"&request("city")&"'"
		end if		  
				set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select * from info2 where 1=1"&province&city&" and infotype='��Դ��Ϣ' order by infoid desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='��Ӧ' Order by InfoID desc")
				%>
      <%if rs_info.bof and rs_info.eof then%>
      <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0> 
                <tr> 
                  <td  valign=top height=67>����Ϣ</td>
                </tr> 
              </table>
                <%else %>
				<%dim k%>
            <table cellspacing=2 cellpadding=2 border=0 width=100%>
              <%do while not rs_info.eof%>
              <tr <%if k mod 2=0 then response.Write "bgcolor=#CCCCCC" end if%>> 
                  
                <td  valign=top align="left"><%=rs_info("city")%>&nbsp;<%=rs_info("area")%></td>
				  
                <td  valign=top align="center">��</td>
				  
                <td  valign=top align="center"><%=rs_info("city2")%>&nbsp;<%=rs_info("area2")%></td>
				  
                <td  valign=top align="center"><%=rs_info("cartype")%></td>
				  
                <td  valign=top align="center"><%=rs_info("carload")%>��</td>
				  
                <td  valign=top align="center"><%=rs_info("addtime")%></td>
				  
                <td  align="center"><A href="javascript:openwindow('truck/detail.asp?InfoID=<%=rs_info("InfoID")%>',500,400)">����</A></td>
              </tr> 
				<%rs_info.movenext
				k=k+1
			    loop
				rs_info.close
				set rs_info=nothing
				%>				
          </table>
			   <%end if%>			  
              <table cellspacing=0 width="100%" border=0>
                <tr> 
                  <td width="61%" height=24>
                  </td>
                  <td width="39%" height=24 align="center"><a href="truck/index.asp">����&gt;&gt;</a>
                  </td>
                </tr>
              </table>
            
    </td>
          </tr>
      </table>