<!--#include file="../Inc/Conn.asp"-->

<%
	Function strLength(str)
		If isNull(str) Or Str = "" Then
			StrLength = 0
			Exit Function
		End If
		Dim WINNT_CHINESE
		WINNT_CHINESE=(len("例子")=2)
		If WINNT_CHINESE Then
			Dim l,t,c
			Dim i
			l=len(str)
			t=l
			For i=1 To l
				c=asc(mid(str,i,1))
				If c<0 Then c=c+65536
				If c>255 Then t=t+1
			Next
			strLength=t
		Else 
			strLength=len(str)
		End If
	End Function


dim RegUserName,FoundErr,ErrMsg
RegUserName=trim(request("UserName"))
if RegUserName="" or strLength(RegUserName)>14 or strLength(RegUserName)<4 then 
	founderr=true
	errmsg=errmsg & "<br><li>请输入用户名(不能大于14小于4)</li>"
else
  	if Instr(RegUserName,"=")>0 or Instr(RegUserName,"%")>0 or Instr(RegUserName,chr(32))>0 or Instr(RegUserName,"?")>0 or Instr(RegUserName,"&")>0 or Instr(RegUserName,";")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,"'")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,chr(34))>0 or Instr(RegUserName,chr(9))>0 or Instr(RegUserName,"")>0 or Instr(RegUserName,"$")>0 then
		errmsg=errmsg+"<br><li>用户名中含有非法字符</li>"
		founderr=true
	end if
end if
if founderr=false then
	dim sqlCheckReg,rsCheckReg,chkadminname,rschkadminname

	chkadminname="select * from Manage_user where UserName='" & RegUserName & "'"
	set rschkadminname=conn.execute(chkadminname) 
	if not(rschkadminname.bof and rschkadminname.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>尊敬的“" & RegUserName & "”是咱们网站的管理员！请回避！</li>"
	else

	sqlCheckReg="select * from qyml where user ='" & RegUserName & "'"
	set rsCheckReg=server.createobject("adodb.recordset")
	rsCheckReg.open sqlCheckReg,conn,1,1
	if not(rsCheckReg.bof and rsCheckReg.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>“" & RegUserName & "”已经存在！请换一个用户名再试试！</li>"
	else
	end if
	rsCheckReg.close
	set rsCheckReg=nothing
	end if
	rschkadminname.close
	set rschkadminname=nothing
end if		
%>
<html>
<head>
<title>检查用户名——诚信物流网</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/style.css" rel=stylesheet type=text/css> 
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
if founderr=false then
	call Success()
else
	call WriteErrmsg()
end if
%>
</body>
</html>
<%
call CloseConn

sub WriteErrMsg()
    response.write "<table align='center' width='100%' border='0' cellpadding='0' cellspacing='1' class='border'>"
    response.write "<tr><td align='center' height='15' class='title'>错误提示</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'>" & errmsg & "<p align='center'>【<a href='javascript:onclick=window.close()'>关 闭</a>】<br></p></td></tr>"
	response.write "</table>" 
end sub

sub Success()
    response.write "<table align='center' width='100%' border='0' cellpadding='0' cellspacing='1' class='border'>"
    response.write "<tr><td align='center' height='15' class='title'>恭喜您！</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>“" & RegUserName & "”尚未被人使用，赶紧注册吧！<p align='center'>【<a href='javascript:onclick=window.close()'>关 闭</a>】<br></p></td></tr>"
	response.write "</table>" 
end sub
%>