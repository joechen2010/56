<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<%
set rs=conn.execute("select * from qyml where user='"&session("user")&"'")
if rs.eof and rs.bof then
    response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	response.end
else
qyjj=replace(rs("qyjj"),"&nbsp;",chr(32))
qyjj=replace(qyjj,"<br>",chr(13))
qyjj=replace(qyjj,"<BR>",chr(13))

%>
<HTML>
<HEAD>
<TITLE>���������� ��Ա�������� - �������˵����ϼ�԰</TITLE>
<meta http-equiv="Content-Language" content="zh-cn"> 
<META qyjj=zh-cn http-equiv=qyjj-Language> 
<META qyjj="text/html; charset=gb2312" http-equiv=qyjj-Type> 
<META qyjj="Microsoft FrontPage 5.0" name=GENERATOR content="Microsoft FrontPage 5.0">
<style type="text/css">
<!--
body {
	background-color: #2C68B1;
	margin: 0px;
}
-->
</style>
<link href="images/main.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #000000}
.style2 {color: #FFFFFF}
.STYLE3 {color: #FF0000}
-->
</style>
<script language=JavaScript src="../inc/p_c_a.js"></script>
<SCRIPT language=javascript src="../../login/images/location.js"></SCRIPT> 
<SCRIPT language=JavaScript type=text/javascript>
//----��ӵ�ѡ���
function addit(from, to) {
var src = document.getElementById(from).options;
var dst = document.getElementById(to).options;

var selected_value = [];
var selected_text = [];

//��ȡ��ѡ������
for(var i = 0; i < dst.length; i++) {
  selected_value[i] = dst[i].value;
  selected_text[i] = dst[i].text;
}

// ��ȡԴ���
for(var i = 1; i < src.length; i++) {
  if(src[i].selected) {
		var exists = false;
	
		for(var j = 1; j < dst.length; j++) {
		  if(dst[j].value == src[i].value) {
			exists = true;
			break;
		  }
		}
		
		if(exists==false) {
		  dst.length++;
		  selected_value[dst.length-1] = src[i].value;
		  selected_text[dst.length-1] = src[i].text;
		}
  }
}

	//�����ѡ���
	for(var j = 0; j < selected_value.length; j++) {
	  if( selected_value[j] == "" ) continue;
	  dst[j].value = selected_value[j];
	  dst[j].text = selected_text[j]
	}
}

//ɾ����ѡ����ڵ����
function delit(dst) {
	var menu=document.getElementById(dst).options;
	for(var i = 1; i < menu.length; i++) {
	  if(menu[i].selected) menu[i--] = null;
	}
}

//----�������ر���
function tofull(){
	
	
<%
set rs_s2=conn.execute("select * from Class_1 Order By SortID Asc")
if rs_s2.eof and rs_s2.bof then
'
else
do while not rs_s2.eof
%>  
    var tempstr<%=rs_s2("SortID")%> = "";
	var bearings<%=rs_s2("SortID")%> = document.getElementById("bearings<%=rs_s2("SortID")%>").options;
	var fullbearings<%=rs_s2("SortID")%> = document.getElementById("fullbearings<%=rs_s2("SortID")%>");
	for(var i = 1;i < bearings<%=rs_s2("SortID")%>.length;i++){
		tempstr<%=rs_s2("SortID")%> = tempstr<%=rs_s2("SortID")%>+","+bearings<%=rs_s2("SortID")%>[i].value;
	}
	fullbearings<%=rs_s2("SortID")%>.value = tempstr<%=rs_s2("SortID")%>.substr(1,tempstr<%=rs_s2("SortID")%>.length);
<%
rs_s2.movenext
loop
end if
rs_s2.close
set rs_s2=nothing
%>
}
//-----�����ı����ȣ�����Ӣ�ĺͺ��֣�
function countlength(str) 
{	
	var strlength = 0;
    var pattern = /[\u4E00-\u9FA5]|[\uFE30-\uFFA0]/; 
	for (var i = 0;i < str.length;i++){
		if (pattern.test(str.charAt(i))) {
			strlength = strlength + 1;
		}
		else {
			strlength = strlength + 0.5;
		}
	}
	return strlength;
}
//ȥ���˿ո�
function trim(str) {
	regExp1 = /^ */;
	regExp2 = / *$/;
	return str.replace(regExp1,'').replace(regExp2,'');
} 

//----�ύ��֤����
function checksubmit() {
//��֤��˾��
	var ename = trim(document.getElementById("qymc").value);	
	if (ename==""){
		alert("��˾���Ʊ������д��˾�ڹ��̲��ŵǼǵ�ȫ��!");
		document.getElementById("qymc").focus();
		return false;
	}
	else if(countlength(ename)>80){
		alert("��˾���Ʋ��ܳ���80���ַ�,�������,�뾡������!");
		document.getElementById("qymc").focus();
		return false;
	}

//��֤��ҵ����
	var qyxz = document.getElementById("qyxz").value;
	if (qyxz==""){
		alert("��ѡ����ҵ����!");
		document.getElementById("contact").focus();
		return false;
	}
	
//��֤���˴���
	var legalperson = document.getElementById("frdb").value;
	if(countlength(legalperson)>0){
		if ((countlength(legalperson)>4)||(countlength(legalperson)<2)){
			alert("���˴�������2-4������֮��,����������!");
			document.getElementById("frdb").focus();
			return false;
		}
	}
//��֤ע���ʽ�

//��֤��ҵ����
	var industry1 = document.getElementById("qylb1").checked;
	var industry2 = document.getElementById("qylb2").checked;
	var industry3 = document.getElementById("qylb3").checked;
	if ((industry1||industry2||industry3)==false){
		alert("������ѡ��һ���ҵ!");
		document.getElementById("qylb1").focus();
		return false;
	}
	
//��֤����
	var jygs = document.getElementById("jygs").value;
	if(countlength(jygs)>200){
		alert("��˾�����������200�������ڣ���ǰΪ "+countlength(jygs)+" ��");
		document.getElementById("jygs").focus();
		return false;
	}	
//��֤ͨ�����ύ��
	document.getElementById("form1").submit();
}
</script>
</head>
<body> 
<!--#include file="head.htm"-->
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="160" height="566" valign="top" bgcolor="#2C68B1">
<!--#include file="left.htm"-->
	</td>
    <td id="main" align="center" valign="top" bgcolor="#FFFFFF">
      <table width="534"  border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">����������</a> &gt; �޸Ĺ�˾����</td>
          <td width="17%"><img src="images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="534"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> <%=session("user")%> </font> </strong>����ӭ��������</td>
        </tr>
      </table>
      <table width="100%" height="497"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top">
            <table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="534"  border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="12" height="21" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_1.jpg" width="12" height="25"></td>
                      <td valign="middle" background="images/title_2.jpg" bgcolor="#FFFFFF" >
                        <div align="center"><font color="#FFFFFF">�޸Ĺ�˾����</font></div>
                      </td>
                      <td width="13" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_3.jpg" width="15" height="25"></td>
                      <td width="392" valign="middle">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
              <td valign="top" bgcolor="#FFFFFF">
			  <table width="100%"  border="0">
                             <tr>
                              <td align="left"><div id="TipAll" >
                               ���������б���*����Ŀ��
                              </div></td>
                            </tr>
                           </table>
<table width="100%"   border="0" cellpadding="0" cellspacing="0">
                    
					<tr>
                      <td valign="top">
                        <table  border="0" cellpadding="5" cellspacing="1" bgcolor="#B1D4F2" >
                          <FORM method="POST" action="com_edit_save.asp" name="Form1">
                            <TBODY> 
                            <tr> 
                              <TD align=left height="25" colspan="2" class="white" background="images/title_bg.gif">�� 
                                �޸Ĺ�˾����</TD>
                            </TR>
                            <tr bgcolor="#EDF6FF"> 
                              <TD align=right height="24" width="150"> <font color="#FF0000">*</font> 
                                ��˾���ƣ�</TD>
                              <TD height="24" bgcolor="#EDF6FF"> 
                                <INPUT class="box" maxLength=80 name=qymc size=36 value=<%=rs("qymc")%>>                              </TD>
                            </tr>
                            
                            <tr bgcolor="#EDF6FF"> 
                              <TD width="150" height="24" align=right bgcolor="#EDF6FF">��ҵ���ࣺ</TD>
                              <TD height="24" bgcolor="#EDF6FF">
<%set rs7=conn.execute("select * from Class_1 order by Sortid asc")

	if rs7.eof and rs7.bof then
	response.write "���������Ŀ��"
	response.end
	else
%>
              <select name="SortID">
                <option selected value="<%=trim(rs("qylb"))%>"><%=trim(rs("qylb"))%></option>
                <%do while not rs7.eof%>
                <option value="<%=trim(rs7("sort"))%>"><%=trim(rs7("sort"))%></option>
                <%
        rs7.movenext
        loop
	end if
        rs7.close
		set rs7=nothing
%>
              </select>							  

							 </TD>
                            </tr>
                            <tr bgcolor="#EDF6FF">
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span class="style16">��Ҫ˾��:</span></div></td>
            <td width="35%" bgcolor="#EDF6FF"><input type="text" name="siji" value="<%=rs("siji")%>"></td>
                            </tr>
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>��˾���ڵ�:</span></div></td>
            <td  bgcolor="#EDF6FF"><select id="s1" name="province">
              <option value="ʡ��">ʡ��</option>
            </select>
              <select id="s2" name="city">
                <option>�ؼ���</option>
              </select>
              <select id="s3" name="area">
                <option>�С��ؼ��С���</option>
              </select>
              <span class=red>*</span></td>
            </tr>
						<SCRIPT language=javascript>
            setup();
          </SCRIPT>      
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>��˾��ַ:</span></div></td>
            <td bgcolor="#EDF6FF"><input id=address  name="address" value="<%=rs("address")%>">
&nbsp; <span class=red>*</span> ֻ����д�ֵ������ƺ�</td>
            </tr>	
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>�̶��绰:</span></div></td>
            <td bgcolor="#EDF6FF"><span class=p141>
              <input id=Ys_m_M_phone1 style="with: 37px" name="phone1" type="hidden" value="<%=rs("phone1")%>">
            </span><span class=p141>
            <input 
                  id=Ys_m_M_phone2 style="with: 53px" name="phone2" size="4" value="<%=rs("phone2")%>">
            </span>-<span class=p141>
            <input 
                  id=Ys_m_M_phone3 style="with: 103px" name="phone3" value="<%=rs("phone3")%>">
            </span><span class=red>* </span>����-�绰����</td>
            </tr>
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>��������:</span></div></td>
            <td bgcolor="#EDF6FF"><span class=p141>
              <input id=tel1 style="with: 37px" name="fax1" type="hidden" value="<%=rs("fax1")%>">
            </span><span class=p141>
            <input id=tel2 style="with: 53px" name="fax2" size="4" value="<%=rs("fax2")%>">
            </span>-<span class=p141>
            <input id=tel3 style="with: 103px" name="fax3" value="<%=rs("fax3")%>">
            </span>�������</td>
            </tr>			
          <tr>
            <td width="150" bgcolor="#EDF6FF" align="right">�֡���:</td>
            <td bgcolor="#EDF6FF"><input id=handset  name=mobile value="<%=rs("mobile")%>"></td>
            </tr>
          <tr>
            <td width="150" bgcolor="#EDF6FF" align="right">�ʡ�����:</td>
            <td bgcolor="#EDF6FF"><input id=post name=post value="<%=rs("post")%>"></td>
            </tr>
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>E-mail:</span></div></td>
            <td width="35%" bgcolor="#EDF6FF"><span class=p141>
              <input  name=email value="<%=rs("email")%>">
            </span></td>
            </tr>
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>������ַ:</span></div></td>
            <td width="35%" bgcolor="#EDF6FF"><span class=p141>
              <input id=url name=web value="<%=rs("web")%>">
            </span></td>
            </tr>									

                            <tr bgcolor="#EDF6FF"> 
                              <TD align=right width="150"><font color="#FF0000">*</font> 
                                ��˾��飺</TD>
                              <TD bgcolor="#EDF6FF">
                                <textarea name="content" cols="60" rows="10" id="content"><%=dvHTMLCode(rs("qyjj"))%></textarea>                              </TD>
                            </tr>
                            <tr bgcolor="#EDF6FF"> 
                              <TD align=right height="36" colspan="2"> 
                                <p align="center"> 
                                  <input type="submit" value=" ȷ���޸� " class="btsub">
                                  &nbsp;&nbsp;&nbsp; 
                                  <input type="button" value=" ȡ������ " onclick=javascript:history.go(-1)  class="btsub">
                              </TD>
                            </tr>
                          </FORM>
                        </table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="10"></td>
                            </tr>
                        </table>
                          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
                              
                            <td width="7" height="6"><img src="images/con_1.gif" width="7" height="6"></td>
                              
                            <td background="images/con_bg1.gif"></td>
                              
                            <td width="6" height="6"><img src="images/con_2.gif" width="6" height="6"></td>
                            </tr>
                            <tr> 
                              
                            <td background="images/con_bg2.gif"></td>
                              <td align="left" valign="top" bgcolor="#FEFFD2"> 
                                <table width="100%"  border="0" cellpadding="5" cellspacing="0">
                                  <tr> 
                                    <td align="left">����ϸ�˶����ύ����Ϣ�������Ϣ���󣬻�Ӱ�����Ľ��ס�</td>
                                  </tr>
                                  <tr>
                                    <td height="1" align="left" bgcolor="#FFCC00"></td>
                                  </tr>
                                  <tr> 
                                    <td align="left">��˾��ϵ��ʽ�����Ļ�Ա���������е�ע����Ϣ����ͻ����ñ���һ�¡� 
                                    </td>
                                  </tr>
                                </table> </td>
                              
                            <td background="images/con_bg4.gif"></td>
                            </tr>
                            <tr> 
                              
                            <td width="7" height="6"><img src="images/con_3.gif" width="7" height="6"></td>
                              
                            <td background="images/con_bg3.gif"></td>
                              
                            <td width="6" height="6"><img src="images/con_4.gif" width="6" height="6"></td>
                            </tr>
                          </table>
	                  </td>
					</tr>
                  </table></td>
            </tr>
          </table>
                        <table width="534"  border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td align="right"><a href="#top"><img src="images/icon_top.gif" alt="�ص�ҳ�涥��" border="0" align="absmiddle" width="15" height="18"></a> 
                  <a href="#top">�ټ��һ��</a></td>
              </tr>
            </table></td>
        </tr>
      </table>
</td>
  </tr>
</table>
<!--#include file="bottom.htm"-->
</body>
</HTML>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing

sub showtypeoption(id)
dim rs_tt
ssql="select * from Class_2 Where TypeID="&cstr(id)
	set rs_tt=conn.execute(ssql)
	if rs_tt.eof and rs_tt.bof then
	    response.write "<option value=""11"">11</option>"
	else
		response.write "<option value="""&rs_tt("typeid")&""">"&rs_tt("typename")&"</option>"
	end if
	rs_tt.close
	set rs_tt=nothing
end sub
%>