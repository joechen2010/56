<HTML>
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<script language=JavaScript src="../inc/p_c_a.js"></script>
<script language=JavaScript src="../inc/p_c_a2.js"></script>
<LINK href="../images/css.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 bgcolor="#FFFFBF" >
<TABLE width=300 border=0 align=center cellPadding=0 cellSpacing=0 bgcolor="#FFFFBD">
    <TR> 
    <TD align="center">
<font color="#990000"> 车辆档案信息查询</font>　	</TD>
    </TR>
</TABLE>
<table width="310" height="136" border=0 align=center cellpadding=0 cellspacing=0 id=Table1 bgcolor="#FFFFBD">
	  <form method="POST" action="search.asp" name="form1" target="right">
                      <tr> 
                        <td align="center">车主所在地:<br>
                          <select id="s1" name="province">
                            <option value="省份">省份</option>
                          </select>
                          <select id="s2" name="city">
                            <option>地级市12332</option>
                          </select>
                          <select id="s3" name="area">
                            <option>市、县级市、县</option>
                          </select>	
						  <br>						
				车型：
                        <select name=cartype id="cartype">
		                   <option value="有车">有车</option>
		                   <option value="普通车">普通车</option>
		                   <option value="软蓬车">软蓬车</option>
		                   <option value="半挂车">半挂车</option>
		                   <option value="厢式车">厢式车</option>
		                   <option value="冷藏车">冷藏车</option>
		                   <option value="集装箱">集装箱</option>
		                   <option value="平板车">平板车</option>
		                   <option value="起重车">起重车</option>
		                   <option value="自卸车">自卸车</option>
		                   <option value="油罐车">油罐车</option>
		                   <option value="后八轮">后八轮</option>
		                   <option value="前四后八">前四后八</option>
		                   <option value="双桥车">双桥车</option>
		                   <option value="加长挂">加长挂</option>
		                   <option value="半封车">半封车</option>
		                   <option value="高栏车">高栏车</option>
		                   <option value="保温车">保温车</option>
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
<%call advertise("4")%>
</BODY>
</HTML>

