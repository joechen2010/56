<%
'**********************************
'��������showpage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       pp ----��ʾҳ��
'       maxperpage  ----ÿҳ����
'       viewpage   ----��ǰҲ��
'       rssum ---��¼����
'       strUnit     ----������λ
'       font_color-----������ɫ
'**************************************************
function showpage(strUnit,sfilename,pp,font_color)
dim strTemp
strTemp="<table border=0 width='100%' cellspacing=0 cellspacing=0>" &vbcrlf
strTemp=strTemp&"<tr><td height=25 align=center class=s>" &vbcrlf
strTemp=strTemp&"&nbsp;��<font color=red>"&rssum&"</font>"&strUnit&"&nbsp;&nbsp;"&vbcrlf
strTemp=strTemp&"ҳ��:<font color=red>"&viewpage&"</font>/<font color=red>"&thepages&"</font>&nbsp;&nbsp;"&vbcrlf
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
    strTemp=strTemp&" <a href='"& strUrl &"' title='��һҳ'>[|<��һҳ]</a> " & _
		" <a href='"& strUrl &"page="&pl-1&"' title='��һҳ'>[<��һҳ]</a> "
  end if
  for pi=pl to pr
    if cint(viewpage)=cint(pi) then
      strTemp=strTemp&" <font color=" & font_color & ">[" & pi & "]</font> "
    else
      strTemp=strTemp&" <a href='"& strUrl &"page="& pi &"' title='�� " & pi & " ҳ' clases=0>[" & pi & "]</a> "
    end if
  next
  if pr<thepages then
    strTemp=strTemp&" <a href='"& strUrl &"page="&pi&"' clases=0 title='��һҳ'>[��һҳ>]</a> " & _
		   " <a href='"& strUrl &"page="& thepages &"' title='���һҳ'>[���һҳ>|]</a> "
  end if
  end if
  strTemp=strTemp&"&nbsp;</td>"&vbcrlf
  strTemp=strTemp&"</tr>"&vbcrlf
  strTemp=strTemp&"</table>"&vbcrlf
  Response.Write strTemp
end function

'**************************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
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
'1=�����ͣ�0=�ı���
	Dim ParaValue
	ParaValue=Request(ParaName)
	If ParaType=1 Then
		If Not isNumeric(ParaValue) Then
            call WriteErrMsg("�Բ������������ʸ���ҳ��") 
		End If
	Else
		ParaValue=Replace(ParaValue,"'","''")
		ParaValue=Replace(ParaValue,";","��")
		ParaValue=Replace(ParaValue,"select","Select",1,-1,1)
	End if
	SafeRequest=ParaValue
End function

'**************************************************
'��������ReplaceBadChar
'��  �ã����˷Ƿ���SQL�ַ�
'��  ����strChar-----Ҫ���˵��ַ�
'����ֵ�����˺���ַ�
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
strErr=strErr &"<HTML><HEAD><TITLE>�����ʵ���ҳ��������</TITLE>"
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
strErr=strErr &"<TR><TD><B>�����ʵ���ҳ���ִ��󣬿��ܲ��������ԭ��</B></TD></TR>"
strErr=strErr &" <TR><TD style=""LINE-HEIGHT: 150%"">(1) �����ʵ�ҳ�治���ڡ�<BR> "
strErr=strErr &"(2) ��صķ������Ѿ�ֹͣ���С�<BR>(3) �����κ����ʣ����뱾վվ����ϵ��"
strErr=strErr &" <P align=center><A href=""javascript:history.go(-1);"">������ҳ</A></P></TD></TR></TBODY></TABLE>"
strErr=strErr &"</TD></TR> <TR> <TD colSpan=2><IMG height=21 alt="""" src=""../Error/error_08.gif"" width=778>"
strErr=strErr &"</TD></TR></TBODY></TABLE></td></tr></table></BODY></HTML>"
response.write strErr
response.end
end sub

sub WriteErrMsg2()
dim strErr2
strErr2=""
strErr2=strErr2 &"<HTML><HEAD><TITLE>�����ʵ���ҳ��������</TITLE>"
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
strErr2=strErr2 &"<TR><TD><B>�����ʵ���ҳ���ִ��󣬿��ܲ��������ԭ��</B></TD></TR>"
strErr2=strErr2 &" <TR><TD style=""LINE-HEIGHT: 150%"">(1) �����ʵ�ҳ�治���ڡ�<BR> "
strErr2=strErr2 &"(2) ��صķ������Ѿ�ֹͣ���С�<BR>(3) �����κ����ʣ����뱾վվ����ϵ��"
strErr2=strErr2 &" <P align=center><A href=""javascript:history.back();"">������ҳ</A></P></TD></TR></TBODY></TABLE>"
strErr2=strErr2 &"</TD></TR> <TR> <TD colSpan=2><IMG height=21 alt="""" src=""../../Error/error_08.gif"" width=778>"
strErr2=strErr2 &"</TD></TR></TBODY></TABLE></td></tr></table></BODY></HTML>"
response.write strErr2
response.end
end sub

'------------------------------------------------------------
'�Ƚ��ַ����ĳ��Ⱥ���Ϊ2������
'------------------------------------------------------------
Function strLength(str)
	If isNull(str) Or Str = "" Then
		StrLength = 0
		Exit Function
	End If
	Dim WINNT_CHINESE
	WINNT_CHINESE=(len("����")=2)
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
'��������gotTopic
'���ã���ȡ�ַ�������
'������str --- Ҫ��ȡ���ַ���
'      strlen --- ΪҪ��ȡ������
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
gotTopic=left(str,i)&"��"
exit for
else
gotTopic=str
end if
next
end function

'**************************************************
'��������IsValidEmail
'��  �ã����Email��ַ�Ϸ���
'��  ����email ----Ҫ����Email��ַ
'����ֵ��True  ----Email��ַ�Ϸ�
'       False ----Email��ַ���Ϸ�
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
'��������showpage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
'**************************************************
sub showpage2(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><tr><td>"
	strTemp=strTemp & "�� <font color=blue><b>" & totalnumber & "</b></font> " & strUnit & "&nbsp;&nbsp;&nbsp;"
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "��һҳ βҳ"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></table>"
	response.write strTemp
end sub


sub WriteRegErrmsg()
     Response.Write "<script language=javascript>alert('" & errmsg & "\n\nϵͳ���Զ�����ǰһҳ��...');history.back();</script>"
	 Response.End
end sub

Function ShowError(str)
	Response.Write "<script language=javascript>alert('" & str & "\n\nϵͳ���Զ�����ǰһҳ��...');history.back();</script>"
	'Response.End
End function
%>