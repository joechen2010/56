
<HTML>
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<script language=JavaScript src="../inc/p_c_a.js"></script>
<script language=JavaScript src="../inc/p_c_a2.js"></script>
<LINK href="../images/css.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 bgcolor="#FFFFBF">
<TABLE width=300 border=0 align=center cellPadding=0 cellSpacing=0 bgcolor="#FFFFBD">
    <TR> 
    <TD align="center">
<font color="#990000"> 客运专线信息查询</font>　	</TD>
    </TR>
</TABLE>
<table width="310" height="136" border=0 align=center cellpadding=0 cellspacing=0 id=Table1 bgcolor="#FFFFBD">
	  <form method="POST" action="search.asp" name="form1" target="right">
                      <tr> 
                        <td align="center">出发地:
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
                         到达地: 
                          <select id="t1" name="province2">
                            <option value="省份">省份</option>
                          </select>
                          <select id="t2" name="city2">
                            <option>地级市</option>
                          </select>
                          <select id="t3" name="area2">
                            <option selected>市、县级市、县</option>
                          </select>	
						  <br>
					    &nbsp;<input type="submit" value="搜 索"  name="button">
						</td>
                      </tr>
		  <SCRIPT language=javascript>
            setup();
          </SCRIPT>
          <SCRIPT language=javascript>
            setup2();
         </SCRIPT>
  </form>		
</table>
<!--#include file="../inc/advertise.asp"-->
<%call advertise("3")%>
</BODY></HTML>