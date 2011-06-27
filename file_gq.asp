<table width="960" border="0" align="center" cellspacing="0">
  <tr> 
    <td valign="top" width="50%"> 
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
                        <strong>最新车辆档案</strong>]</b></td>
                      <td style="FONT-SIZE: 12px" align="left" bgcolor="#bfcad7" height="25" width="80"><a href="cheliang/index.asp" target="_blank">更多</a></td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
            <%set rs_gq=server.CreateObject("adodb.recordset")
		   sql="select top 5 * from file_info a,qyml b where a.gsid=b.user order by infoid desc"
		  rs_gq.open sql,conn,1,1
		  %>
            <table border="0" cellspacing="1" cellpadding="4" width="100%" bgcolor="#3366FF">
              <%if rs_gq.bof and rs_gq.eof then%>
              <tr bgcolor="#336699"> 
                <td valign="top" colspan="4" height="19">暂时无信息</td>
              </tr>
              <tr bgcolor="#336699" align="center"> 
                <%else%>
                <td bgcolor="#336699"  align="center"><font color="#FFFFFF">车牌号</font></td>
                <td bgcolor="#336699" align="center"><font color="#FFFFFF">车型</font></td>
                <td bgcolor="#336699" align="center"><font color="#FFFFFF">车籍</font></td>
                <td bgcolor="#336699" align="center"><font color="#FFFFFF">车主</font></td>
              </tr>
              <%do while not rs_gq.eof%>
              <tr bgcolor="#FFFFFF" align="center"> 
                <td><%=rs_gq("carnum")%></td>
                <td><%=rs_gq("cartype")%></td>
                <td><%=rs_gq("city")%></td>
                <td> 
                  <%if len(rs_gq("qymc"))>12 then%>
                  <a href="javascript:openwindow('cheliang/detail.asp?InfoID=<%=rs_gq("InfoID")%>',500,400)"> 
                  <%=left(rs_gq("qymc"),12)%></a> 
                  <%else%>
                  <a href="javascript:openwindow('cheliang/detail.asp?InfoID=<%=rs_gq("InfoID")%>',500,400)"> 
                  <%=rs_gq("qymc")%></a> 
                  <%end if%>
                </td>
              </tr>
              <%rs_gq.movenext
			    loop
				rs_gq.close
				set rs_gq=nothing
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
                        <strong>最新供求信息</strong>]</b></td>
                      <td style="FONT-SIZE: 12px" align="left" width="80" bgcolor="#bfcad7" height="25"><a href="trade/index.asp" target="_blank">更多</a> 
                      </td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
            <%set rs_gq=server.CreateObject("adodb.recordset")
		  rs_gq.open "select top 5 * from gq_info order by infoid asc",conn,1,1
		  %>
            <table border="0" cellspacing="1" cellpadding="4" width="100%" bgcolor="#990000">
              <%if rs_gq.bof and rs_gq.eof then%>
              <tr> 
                <td valign="top" colspan="3">暂时无信息</td>
              </tr>
              <%else%>
              <tr align="center" bgcolor="#993300"> 
                <td><font color="#FFFFFF">供求</font></td>
                <td><font color="#FFFFFF">产品或服务名称</font></td>
                <td><font color="#FFFFFF">发布日期</font></td>
              </tr>
              <%do while not rs_gq.eof%>
              <tr align="center" bgcolor="#FFFFFF"> 
                <td><font color="#FF6600">【<%=rs_gq("infotype")%>】</font></td>
                <td> 
                  <%if len(rs_gq("title"))<12 then%>
                  <a href="javascript:openwindow('trade/detail.asp?infoid=<%=rs_gq("infoid")%>',500,400)"> 
                  <%=rs_gq("title")%></a> 
                  <%else%>
                  <a href="javascript:openwindow('trade/detail.asp?infoid=<%=rs_gq("infoid")%>',500,400)"> 
                  <%=left(rs_gq("title"),12)%></a> 
                  <%end if%>
                </td>
                <td><%=rs_gq("addtime")%></td>
              </tr>
              <%rs_gq.movenext
			    loop
				rs_gq.close
				set rs_gq=nothing
			  %>
              <%end if%>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
