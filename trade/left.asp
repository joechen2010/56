<HTML>
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<script language=JavaScript src="../inc/p_c_a.js"></script>
<script language=JavaScript src="../inc/p_c_a2.js"></script>
<LINK href="../images/css.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 bgcolor="#FFFFBD">
<TABLE width=300 border=0 align=center cellPadding=4 cellSpacing=4 bgcolor="#FFFFBD">
  <TR> 
    <TD align="center">
<font color="#990000"> 供求信息查询</font>　	</TD>
    </TR>
</TABLE>
<table width="300" border=0 align=center cellpadding=4 cellspacing=4 id=Table1 bgcolor="#FFFFBD">
  <form method="POST" action="search.asp" name="form1" target="right">
    <tr> 
      <td>所在省市: </td>
    </tr>
    <tr>
      <td>
        <select id="s1" name="province">
          <option value="省份">省份</option>
        </select>
        <select id="s2" name="city">
          <option>地级市</option>
        </select>
        <select id="s3" name="area">
          <option>地级市</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td>信息类型： 
        <select name="infotype">
          <option selected value="供应">供应</option>
          <option value="采购">采购</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td> &nbsp; 
        <input type="submit" value="搜 索"  name="button">
      </td>
    </tr>
    <SCRIPT language=javascript>
            setup();
          </SCRIPT>
  </form>
</table>
<!--#include file="../inc/advertise.asp"-->
<%call advertise("8")%>
</BODY></HTML>