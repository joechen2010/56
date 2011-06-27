<!--#include file="../inc/advertise.asp"-->
<HTML>
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<script language=JavaScript src="../inc/p_c_a.js"></script>
<script language=JavaScript src="../inc/p_c_a2.js"></script>
<LINK href="../images/css.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 bgcolor="#FFFFBF">
<TABLE cellSpacing=0 cellPadding=0 width=310 align=center border=0>
  <TR> 
    <TD align="center">
<font color="#990000">配货信息查询</font>　
	</TD>
    </TR>
</TABLE>
<table width="310" height="136" border=0 align=center cellpadding=4 cellspacing=4 id=Table1>
  <form method="POST" action="search.asp" name="form1" target="right">
    <tr> 
      <td> 我要查找： 
        <select name="type_search">
          <option value="1" selected>两地所有单向信息</option>
          <option value="2">两地所有双向信息</option>
          <option value="3">单向货源信息</option>
          <option value="4">单向车源信息</option>
        </select>
      </td>
    </tr>
    <tr>
      <td>出发地: 
        <select id="s1" name="province">
          <option value="省份" selected>省份</option>
        </select>
        <select id="s2" name="city">
          <option selected>地级市</option>
        </select>
        <select id="s3" name="area">
          <option selected>选择区县</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td>到达地: 
        <select id="t1" name="province2">
          <option value="省份" selected>省份</option>
        </select>
        <select id="t2" name="city2">
          <option selected>地级市</option>
        </select>
        <select id="t3" name="area2">
          <option selected>选择区县</option>
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
    <SCRIPT language=javascript>
            setup2();
         </SCRIPT>
  </form>
</table>
<%call advertise("1")%>
</BODY>
</HTML>

