<HTML>
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<script language=JavaScript src="../inc/p_c_a.js"></script>
<script language=JavaScript src="../inc/p_c_a2.js"></script>
<LINK href="../images/css.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR><style type="text/css">
<!--
body {
	background-color: #FFFFBD;
}
-->
</style></HEAD>
<BODY leftMargin=0 topMargin=0>
<TABLE width=300 border=0 align=center cellPadding=0 cellSpacing=0 bgcolor="#FFFFBD">
    <TR> 
    <TD align="center">
<font color="#990000">招聘求职信息查询</font>　	</TD>
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
						  <br>职位名称：<input type="text" name="job">&nbsp;<br>
                          <select name="infotype">
						              <option value="招聘信息">招聘</option>
									  <option value="求职信息">求职</option>
						  </select>	
						  工作类型：<select name="jobtype">
						              <option value="全部">全部</option>
						              <option value="兼职">兼职</option>
									  <option value="全职">全职</option>
									</select>
							<br>
					    &nbsp;<input type="submit" value="搜 索"  name="button">
						</td>
                      </tr>
			<SCRIPT language=javascript>
            setup();
          </SCRIPT>
  </form>
</table>
<!--#include file="../inc/advertise.asp"-->
<%call advertise("7")%>
</BODY>
</HTML>
