<table width="960" border="0" align="center" cellspacing="0">
  <tr> 
    <td width="50%" valign="top"> 
      <table width=100% border=0 cellpadding=0 cellspacing=0>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF"> 
            <table id="T1" style="BORDER-COLLAPSE: collapse" height="32" cellpadding="0" width="100%"
	border="0">
              <tbody> 
              <tr align="left"> 
                <td valign="top" align="center" width="43" height="32" rowspan="2"><img src="/images/hou1.gif" id="Gqxxlist1_IMG2" height="32" width="43" border="0" /></td>
                <td valign="top" align="center" height="32"> 
                  <table id="T2" style="BORDER-COLLAPSE: collapse" cellpadding="0" width="100%">
                    <tbody> 
                    <tr> 
                      <td style="FONT-SIZE: 14px" bgcolor="#bfcad7" height="25"><b>[ 
                        <strong>最新招聘信息</strong>]</b></td>
                      <td style="FONT-SIZE: 12px" align="left" bgcolor="#bfcad7" height="25" width="80"><a href="job/index.asp" target="_blank">更多</a></td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
            <%set rs_z=server.CreateObject("adodb.recordset")
		   rs_z.open "select top 5 * from zp_info where infotype='招聘信息' order by infoid desc",conn,1,1
		  %>
            <table border="0" cellspacing="1" cellpadding="4" bgcolor="#006666" width="100%">
              <%if rs_z.bof and rs_z.eof then%>
              <tr> 
                <td colspan="3" valign="top">暂时无信息</td>
              </tr>
              <%else%>
              <tr bgcolor="#009999"> 
                <td align="center" valign="top"><font color="#FFFFFF">招聘职位</font></td>
                <td align="center" valign="top"><font color="#FFFFFF">招聘单位</font></td>
                <td align="center" valign="top"><font color="#FFFFFF">所在城市</font></td>
              </tr>
              <%do while not rs_z.eof%>
              <tr bgcolor="#FFFFFF"> 
                <td align="center" valign="top"> 
                  <%title=gotTopic(rs_z("job"),10)
          'topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          'topic=replace(topic,"'","&nbsp;") 
		  %>
                  <A href="javascript:openwindow('job/detail.asp?infoid=<%=rs_z("infoid")%>',500,400)"> 
                  <%=title%></a></td>
                <td align="center" valign="top">
				<%set qymc=conn.execute("select * from qyml where user='"&rs_z("gsid")&"'")
				if not (qymc.eof and qymc.bof) then
				response.write qymc("qymc")
				end if
				qymc.close
				set qymc=nothing
				%>
				</td>
                <td align="center" valign="top"><%=rs_z("province")%>&nbsp;<%=rs_z("city")%></td>
              </tr>
              <%rs_z.movenext
			    loop
				rs_z.close
				set rs_z=nothing
			  %>
              <%end if%>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td align="right" valign="top" width="2"></td>
    <td align="right" valign="top" width="50%"> 
      <table width="100%"  border=0 cellpadding=0 cellspacing=0>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF"> 
            <table id="T1" style="BORDER-COLLAPSE: collapse" height="32" cellpadding="0" width="100%"
	border="0">
              <tbody> 
              <tr align="left"> 
                <td valign="top" align="center" width="43" height="32" rowspan="2"><img height="32" src="/images/qiye.gif" width="43" border="0" alt="车辆档案"></td>
                <td valign="top" align="center" height="32"> 
                  <table id="T2" style="BORDER-COLLAPSE: collapse" cellpadding="0" width="100%">
                    <tbody> 
                    <tr> 
                      <td style="FONT-SIZE: 14px" bgcolor="#bfcad7" height="25"><b>&nbsp;[ 
                        <strong>最新物流人才</strong>]</b></td>
                      <td style="FONT-SIZE: 12px" align="left" width="80" bgcolor="#bfcad7" height="25"><a href="job/index.asp" target="_blank">更多</a></td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
            <%set rs_q=server.CreateObject("adodb.recordset")
		    rs_q.open "select top 5 * from zp_info where infotype='求职信息' order by infoid desc",conn,1,1
		  %>
            <table border="0" cellspacing="1" cellpadding="4" width="100%" bgcolor="#006699">
              <%if rs_q.bof and rs_q.eof then%>
              <tr> 
                <td colspan="3" valign="top">暂时无信息 </td>
              </tr>
              <tr> 
                <%else%>
                <td valign="top" align="center"><font color="#FFFFFF">求职意向</font></td>
                <td valign="top" align="center"><font color="#FFFFFF">姓名</font></td>
                <td valign="top" align="center"><font color="#FFFFFF">所在城市</font></td>
              </tr>
              <%do while not rs_q.eof%>
              <tr bgcolor="#FFFFFF"> 
                <td valign="top" align="center"> 
                  <%title=gotTopic(rs_q("job"),20)%>
                  <a href="javascript:openwindow('job/detail.asp?infoid=<%=rs_q("infoid")%>',500,400)"> 
                  <%=title%></a> </td>
                <td valign="top" align="center">
				<%set qymc=conn.execute("select * from qyml where user='"&rs_q("gsid")&"'")
				if not (qymc.eof and qymc.bof) then
				response.write qymc("name")
				end if
				qymc.close
				set qymc=nothing
				%>				
				</td>
                <td valign="top" align="center"><%=rs_q("province")%>&nbsp;<%=rs_q("city")%></td>
              </tr>
              <%rs_q.movenext
			    loop
				rs_q.close
				set rs_q=nothing
			  %>
              <%end if%>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>

