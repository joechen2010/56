<HTML>
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<script language=JavaScript src="../inc/p_c_a.js"></script>
<script language=JavaScript src="../inc/p_c_a2.js"></script>
<LINK href="../images/css.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0  bgcolor="#FFFFBF">
<TABLE width=300 border=0 align=center cellPadding=0 cellSpacing=0 bgcolor="#FFFFBD">
    <TR> 
    <TD align="center">
<font color="#990000">库房信息查询</font>　	</TD>
    </TR>
</TABLE>
<table width="310" height="136" border=0 align=center cellpadding=0 cellspacing=0 id=Table1 bgcolor="#FFFFBD">
	  <form method="POST" action="search.asp" name="form1" target="right">
                      <tr> 
                        <td align="center">所在省市:<br>
                          <select id="s1" name="province">
                            <option value="省份">省份</option>
                          </select>
                          <select id="s2" name="city">
                            <option>地级市</option>
                          </select>
                          <select id="s3" name="area">
                            <option>地级市</option>
                          </select>						  
						  &nbsp;<br>信息类型：
						<select name="infotype">
	<option selected value="出租">出租</option>
	<option value="求租">求租</option>
	<option value="出售">出售</option>
	<option value="购买">购买</option>					
						</select>	
						  &nbsp;库房类型：
						  <select name="kftype">
	<option selected value="普通">普通</option>
	<option value="综合">综合</option>
	<option value="冷暖">冷暖</option>
	<option value="立体">立体</option>
	<option value="楼房">楼房</option>
	<option value="大棚">大棚</option>
	<option value="露天">露天</option>							
						</select>	<br>										
					    &nbsp;<input type="submit" value="搜 索"  name="button">
						<input type="hidden" value="search" name="action">
						</td>
                      </tr>
			<SCRIPT language=javascript>
            setup();
          </SCRIPT>
  </form>
</table>
<!--#include file="../inc/advertise.asp"-->
<%call advertise("6")%>
</BODY>
</HTML>
