<%
'**********************************
'过程名：showpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       pp ----显示页数
'       maxperpage  ----每页数量
'       viewpage   ----当前也数
'       rssum ---记录总数
'       strUnit     ----计数单位
'       font_color-----字体颜色
'**************************************************
function showpage(strUnit,sfilename,pp,font_color)
dim strTemp
strTemp="<table border=0 width='100%' cellspacing=0 cellspacing=0>" &vbcrlf
strTemp=strTemp&"<tr><td height=25 align=center class=s>" &vbcrlf
strTemp=strTemp&"&nbsp;共<font color=red>"&rssum&"</font>"&strUnit&"&nbsp;&nbsp;"&vbcrlf
strTemp=strTemp&"页次:<font color=red>"&viewpage&"</font>/<font color=red>"&thepages&"</font>&nbsp;&nbsp;"&vbcrlf
strUrl=JoinChar(sfilename)
  if int(thepages)=0 then
    strTemp=strTemp&"<font color=" & font_color & ">[1]</font>"
    'exit function
  else
  dim pn,pi,page_num,ppp,pl,pr
  pi=1
  ppp=pp\2
  if pp mod 2 = 0 then ppp=ppp-1
  pl=viewpage-ppp
  pr=pl+pp-1
  if pl<1 then
    pr=pr-pl+1
    pl=1
    if pr>thepages then pr=thepages
  end if
  if pr>int(thepages) then
    pl=pl+thepages-pr
    pr=thepages
    if pl<1 then pl=1
  end if
  if pl>1 then
    strTemp=strTemp&" <a href='"& strUrl &"' title='第一页'>[|<第一页]</a> " & _
		" <a href='"& strUrl &"page="&pl-1&"' title='上一页'>[<上一页]</a> "
  end if
  for pi=pl to pr
    if cint(viewpage)=cint(pi) then
      strTemp=strTemp&" <font color=" & font_color & ">[" & pi & "]</font> "
    else
      strTemp=strTemp&" <a href='"& strUrl &"page="& pi &"' title='第 " & pi & " 页' clases=0>[" & pi & "]</a> "
    end if
  next
  if pr<thepages then
    strTemp=strTemp&" <a href='"& strUrl &"page="&pi&"' clases=0 title='后一页'>[后一页>]</a> " & _
		   " <a href='"& strUrl &"page="& thepages &"' title='最后一页'>[最后一页>|]</a> "
  end if
  end if
  strTemp=strTemp&"&nbsp;</td>"&vbcrlf
  strTemp=strTemp&"</tr>"&vbcrlf
  strTemp=strTemp&"</table>"&vbcrlf
  Response.Write strTemp
end function

'**************************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'**************************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

Function SafeRequest(ParaName,ParaType)
'1=数字型，0=文本型
	Dim ParaValue
	ParaValue=Request(ParaName)
	If ParaType=1 Then
		If Not isNumeric(ParaValue) Then
            call WriteErrMsg("对不起，请正常访问该网页！") 
		End If
	Else
		ParaValue=Replace(ParaValue,"'","''")
		ParaValue=Replace(ParaValue,";","；")
		ParaValue=Replace(ParaValue,"select","Select",1,-1,1)
	End if
	SafeRequest=ParaValue
End function

'**************************************************
'函数名：ReplaceBadChar
'作  用：过滤非法的SQL字符
'参  数：strChar-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
function ReplaceBadChar(strChar)
	if strChar="" then
		ReplaceBadChar=""
	else
		ReplaceBadChar=replace(replace(replace(replace(replace(replace(replace(strChar,"'",""),"*",""),"?",""),"(",""),")",""),"<",""),".","")
	end if
end function

function dvHTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    dvHTMLEncode = fString
end if
end function

function dvHTMLCode(fString)
if not isnull(fString) then
    fString = replace(fString, "&gt;", ">")
    fString = replace(fString, "&lt;", "<")

    fString = Replace(fString,  "&nbsp;"," ")
    fString = Replace(fString, "&quot;", CHR(34))
    fString = Replace(fString, "&#39;", CHR(39))
    fString = Replace(fString, "</P><P> ",CHR(10) & CHR(10))
    fString = Replace(fString, "<BR> ", CHR(10))

    dvHTMLCode = fString
end if
end function

sub WriteErrMsg(errmsg)
dim strErr
strErr=""
strErr=strErr &"<HTML><HEAD><TITLE>您访问的网页出错啦！</TITLE>"
strErr=strErr &"<META http-equiv=Content-Type content=""text/html; charset=gb2312"">"
strErr=strErr &"<LINK href=""../Error/webstyle.css"" type=text/css rel=stylesheet>"
strErr=strErr &"<META content=""MSHTML 6.00.2800.1498"" name=GENERATOR></HEAD>"
strErr=strErr &" <BODY><table width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
strErr=strErr &"  <tr><td><TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>"
strErr=strErr &" <TR><TD><IMG height=87 alt="""" src=""../Error/error_02.gif"" width=166></TD>"
strErr=strErr &" <TD><IMG height=87 alt="""" src=""../Error/error_03.gif"" width=612></TD></TR>"
strErr=strErr &" <TR> <TD><IMG height=16 alt="""" src=""../Error/error_04.gif"" width=166></TD>"
strErr=strErr &"<TD><IMG height=16 alt="""" src=""../Error/error_05.gif"" width=612></TD></TR>"
strErr=strErr &" <TR>  <TD><IMG height=258 alt="""" src=""../Error/error_06.gif"" width=166></TD>"
strErr=strErr &"<TD width=612 background=../Error/error_07.gif>"
strErr=strErr &"<TABLE cellSpacing=0 cellPadding=4 width=""75%"" border=0><TBODY>"
strErr=strErr &"<TR><TD><B>您访问的网页出现错误，可能产生错误的原因：</B></TD></TR>"
strErr=strErr &" <TR><TD style=""LINE-HEIGHT: 150%"">(1) 您访问的页面不存在。<BR> "
strErr=strErr &"(2) 相关的服务器已经停止运行。<BR>(3) 如有任何疑问，请与本站站长联系。"
strErr=strErr &" <P align=center><A href=""javascript:history.go(-1);"">返回上页</A></P></TD></TR></TBODY></TABLE>"
strErr=strErr &"</TD></TR> <TR> <TD colSpan=2><IMG height=21 alt="""" src=""../Error/error_08.gif"" width=778>"
strErr=strErr &"</TD></TR></TBODY></TABLE></td></tr></table></BODY></HTML>"
response.write strErr
response.end
end sub

sub WriteErrMsg2()
dim strErr2
strErr2=""
strErr2=strErr2 &"<HTML><HEAD><TITLE>您访问的网页出错啦！</TITLE>"
strErr2=strErr2 &"<META http-equiv=Content-Type content=""text/html; charset=gb2312"">"
strErr2=strErr2 &"<LINK href=""../../Error/webstyle.css"" type=text/css rel=stylesheet>"
strErr2=strErr2 &"<META content=""MSHTML 6.00.2800.1498"" name=GENERATOR></HEAD>"
strErr2=strErr2 &" <BODY><table width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
strErr2=strErr2 &"  <tr><td><TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>"
strErr2=strErr2 &" <TR><TD><IMG height=87 alt="""" src=""../../Error/error_02.gif"" width=166></TD>"
strErr2=strErr2 &" <TD><IMG height=87 alt="""" src=""../../Error/error_03.gif"" width=612></TD></TR>"
strErr2=strErr2 &" <TR> <TD><IMG height=16 alt="""" src=""../../Error/error_04.gif"" width=166></TD>"
strErr2=strErr2 &"<TD><IMG height=16 alt="""" src=""../../Error/error_05.gif"" width=612></TD></TR>"
strErr2=strErr2 &" <TR>  <TD><IMG height=258 alt="""" src=""../../Error/error_06.gif"" width=166></TD>"
strErr2=strErr2 &"<TD width=612 background=../../Error/error_07.gif>"
strErr2=strErr2 &"<TABLE cellSpacing=0 cellPadding=4 width=""75%"" border=0><TBODY>"
strErr2=strErr2 &"<TR><TD><B>您访问的网页出现错误，可能产生错误的原因：</B></TD></TR>"
strErr2=strErr2 &" <TR><TD style=""LINE-HEIGHT: 150%"">(1) 您访问的页面不存在。<BR> "
strErr2=strErr2 &"(2) 相关的服务器已经停止运行。<BR>(3) 如有任何疑问，请与本站站长联系。"
strErr2=strErr2 &" <P align=center><A href=""javascript:history.back();"">返回上页</A></P></TD></TR></TBODY></TABLE>"
strErr2=strErr2 &"</TD></TR> <TR> <TD colSpan=2><IMG height=21 alt="""" src=""../../Error/error_08.gif"" width=778>"
strErr2=strErr2 &"</TD></TR></TBODY></TABLE></td></tr></table></BODY></HTML>"
response.write strErr2
response.end
end sub

'------------------------------------------------------------
'比较字符串的长度汉字为2个字数
'------------------------------------------------------------
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

'----------------------------------------------------------
'函数名：gotTopic
'作用：截取字符串函数
'参数：str --- 要截取的字符串
'      strlen --- 为要截取的字数
'----------------------------------------------------------
function gotTopic(str,strlen)
dim l,t,c
l=len(str)
t=0
for i=1 to l
c=Abs(Asc(Mid(str,i,1)))
if c>255 then
t=t+2
else
t=t+1
end if
if t>=strlen then
gotTopic=left(str,i)&"…"
exit for
else
gotTopic=str
end if
next
end function

'**************************************************
'函数名：IsValidEmail
'作  用：检查Email地址合法性
'参  数：email ----要检查的Email地址
'返回值：True  ----Email地址合法
'       False ----Email地址不合法
'**************************************************
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

'**************************************************
'过程名：showpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'**************************************************
sub showpage2(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><tr><td>"
	strTemp=strTemp & "共 <font color=blue><b>" & totalnumber & "</b></font> " & strUnit & "&nbsp;&nbsp;&nbsp;"
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></table>"
	response.write strTemp
end sub


sub WriteRegErrmsg()
     Response.Write "<script language=javascript>alert('" & errmsg & "\n\n系统将自动返回前一页面...');history.back();</script>"
	 Response.End
end sub

Function ShowError(str)
	Response.Write "<script language=javascript>alert('" & str & "\n\n系统将自动返回前一页面...');history.back();</script>"
	'Response.End
End function
%>