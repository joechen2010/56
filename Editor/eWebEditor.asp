<!--#include file = "Include/Startup.asp"-->
<%
'������������������������������������
'��                                                                  ��
'��                eWebEditor - eWebSoft�����ı��༭��               ��
'��                                                                  ��
'��  ��Ȩ����: eWebSoft.com                                          ��
'��                                                                  ��
'��  ��������: eWeb�����Ŷ�                                          ��
'��            email:webmaster@webasp.net                            ��
'��            QQ:589808                                             ��
'��                                                                  ��
'��  �����ַ: [��Ʒ����]http://www.eWebSoft.com/Product/eWebEditor/ ��
'��            [֧����̳]http://bbs.eWebSoft.com/                    ��
'��                                                                  ��
'��  ��ҳ��ַ: http://www.eWebSoft.com/   eWebSoft�ŶӼ���Ʒ         ��
'��            http://www.webasp.net/     WEB������Ӧ����Դ��վ      ��
'��            http://bbs.webasp.net/     WEB����������̳            ��
'��                                                                  ��
'������������������������������������
%>

<%
' ��ʼ�������
Dim sContentID, sStyleID, sFullScreen
Dim sStyleName, sStyleDir, sStyleCSS, sStyleUploadDir, nStateFlag, sDetectFromWord, sInitMode, sBaseUrl
Dim sVersion, sReleaseDate
Call InitPara()

' ȡ���а�ť
Dim aButtonCode(), aButtonHTML()
Call InitButtonArray()

' ȡ��ʽ�µĹ���������ť
Dim sToolBar
Call InitToolBar()

' �Ͽ����ݿ�����
Call DBConnEnd()

%>


<html>
<head>
<title>eWebEditor - eWebSoft�����ı��༭��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/<%=sStyleCSS%>/Editor.css" type="text/css" rel="stylesheet">

<Script Language=Javascript>
var sPath = document.location.pathname;
sPath = sPath.substr(0, sPath.length-14);

var sLinkFieldName = "<%=sContentID%>" ;

// ȫ�����ö���
var config = new Object() ;
config.Version = "<%=sVersion%>" ;
config.ReleaseDate = "<%=sReleaseDate%>" ;
config.StyleName = "<%=sStyleName%>";
config.StyleEditorHeader = "<head><link href=\""+sPath+"css/<%=sStyleCSS%>/EditorArea.css\" type=\"text/css\" rel=\"stylesheet\"></head><body MONOSPACE>" ;
config.StyleMenuHeader = "<head><link href=\""+sPath+"css/<%=sStyleCSS%>/MenuArea.css\" type=\"text/css\" rel=\"stylesheet\"></head><body scroll=\"no\" onConTextMenu=\"event.returnValue=false;\">";
config.StyleDir = "<%=sStyleDir%>";
config.StyleUploadDir = "<%=sStyleUploadDir%>";
config.InitMode = "<%=sInitMode%>";
config.AutoDetectPasteFromWord = <%=sDetectFromWord%>;
config.BaseUrl = <%=sBaseUrl%>;
</Script>
<Script Language=Javascript src="include/editor.js"></Script>
<Script Language=Javascript src="include/table.js"></Script>
<Script Language=Javascript src="include/menu.js"></Script>

<script language="javascript" event="onerror(msg, url, line)" for="window">
//return true ;	 // ���ش���
</script>

</head>

<body SCROLLING=no onConTextMenu="event.returnValue=false;" onfocus="VerifyFocus()">

<table border=0 cellpadding=0 cellspacing=0 width='100%' height='100%'>
<tr><td>

	<%=sToolBar%>

</td></tr>
<tr><td height='100%'>

	<table border=0 cellpadding=0 cellspacing=0 width='100%' height='100%'>
	<tr><td height='100%'>
	<input type="hidden" ID="ContentEdit" value="">
	<input type="hidden" ID="ContentLoad" value="">
	<input type="hidden" ID="ContentFlag" value="0">
	<iframe class="Composition" ID="eWebEditor" MARGINHEIGHT="1" MARGINWIDTH="1" width="100%" height="100%" scrolling="yes"> 
	</iframe>
	</td></tr>
	</table>

</td></tr>

<% If nStateFlag = 1 Then %>
<tr><td height=25>

	<TABLE border="0" cellPadding="0" cellSpacing="0" width="100%" class=StatusBar height=25>
	<TR valign=middle>
	<td>
		<table border=0 cellpadding=0 cellspacing=0 height=20>
		<tr>
		<td width=10></td>
		<td class=StatusBarBtnOff id=eWebEditor_CODE onclick="setMode('CODE')"><img border=0 src="buttonimage/<%=sStyleDir%>/modecode.gif" width=50 height=15 align=absmiddle></td>
		<td width=5></td>
		<td class=StatusBarBtnOff id=eWebEditor_EDIT onclick="setMode('EDIT')"><img border=0 src="buttonimage/<%=sStyleDir%>/modeedit.gif" width=50 height=15 align=absmiddle></td>
		<td width=5></td>
		<td class=StatusBarBtnOff id=eWebEditor_VIEW onclick="setMode('VIEW')"><img border=0 src="buttonimage/<%=sStyleDir%>/modepreview.gif" width=50 height=15 align=absmiddle></td>
		</tr>
		</table>
	</td>
	<td align=right>
		<table border=0 cellpadding=0 cellspacing=0 height=20>
		<tr>
		<td style="cursor:pointer;" onclick="sizeChange(300)"><img border=0 SRC="buttonimage/<%=sStyleDir%>/sizeplus.gif" width=20 height=20 alt="���߱༭��"></td>
		<td width=5></td>
		<td style="cursor:pointer;" onclick="sizeChange(-300)"><img border=0 SRC="buttonimage/<%=sStyleDir%>/sizeminus.gif" width=20 height=20 alt="��С�༭��"></td>
		<td width=40></td>
		</tr>
		</table>
	</td>
	</TR>
	</Table>

</td></tr>
<% End If %>

</table>

<div id="divTemp" style="VISIBILITY: hidden; OVERFLOW: hidden; POSITION: absolute; WIDTH: 1px; HEIGHT: 1px"></div>

</body>
</html>


<%


' ��ʾ���ô�����ʾ
Sub ShowErr(str)
	Call DBConnEnd()
	Response.Write "���ô���" & str
	Response.End
End Sub

' ��ʼ���������
Sub InitPara()
	' ȡȫ����־
	sFullScreen = Trim(Request.QueryString("fullscreen"))
	' ȡ��Ӧ������ID
	sContentID = Trim(Request.QueryString("id"))
	If sContentID = "" Then ShowErr "�봫����ò���ID�������ص����ݱ���ID��"

	' ȡ��ʽ��ʼֵ
	sStyleName = Trim(Request.QueryString("style"))
	If sStyleName = "" Then sStyleName = "standard"

	sSql = "select * from ewebeditor_style where s_name='" & sStyleName & "'"
	oRs.Open sSql, oConn, 0, 1
	If Not oRs.Eof Then
		sStyleID = oRs("S_ID")
		sStyleName = oRs("S_Name")
		sStyleDir = oRs("S_Dir")
		sStyleCSS = oRs("S_CSS")
		sStyleUploadDir = oRs("S_UploadDir")
		nStateFlag = oRs("S_StateFlag")
		sDetectFromWord = oRs("S_DetectFromWord")
		sInitMode = oRs("S_InitMode")
		sBaseUrl = oRs("S_BaseUrl")
	Else
		ShowErr "��Ч����ʽStyle�������룬���Ҫʹ��Ĭ��ֵ�������գ�"
	End If
	oRs.Close

	' ȡ�汾�ż���������
	sSql = "select sys_version,sys_releasedate from ewebeditor_system"
	oRs.Open sSql, oConn, 0, 1
	sVersion = oRs(0)
	sReleaseDate = oRs(1)
	oRs.Close
End Sub

' ��ʼ����ť����
Sub InitButtonArray()
	Dim i
	sSql = "select * from ewebeditor_button order by b_order asc"
	oRs.Open sSql, oConn, 0, 1
	i = 0
	Do While Not oRs.Eof
		i = i + 1
		Redim Preserve aButtonCode(i)
		Redim Preserve aButtonHTML(i)
		aButtonCode(i) = oRs("B_Code")
		Select Case oRs("B_Type")
		Case 0
			aButtonHTML(i) = "<DIV CLASS=""" & oRs("B_Class") & """ TITLE=""" & oRs("B_Title") & """ onclick=""" & oRs("B_Event") & """><IMG CLASS=""Ico"" SRC=""buttonimage/" & sStyleDir & "/" & oRs("B_Image") & """></DIV>"
		Case 1
			aButtonHTML(i) = "<SELECT CLASS=""" & oRs("B_Class") & """ onchange=""" & oRs("B_Event") & """>" & oRs("B_HTML") & "</SELECT>"
		Case 2
			aButtonHTML(i) = "<DIV CLASS=""" & oRs("B_Class") & """>" & oRs("B_HTML") & "</DIV>"
		End Select
		oRs.MoveNext
	Loop
	oRs.Close
End Sub

' �ɰ�ť����õ���ť���������
Function Code2HTML(s_Code)
	Dim i
	Code2HTML = ""
	For i = 1 To UBound(aButtonCode)
		If UCase(aButtonCode(i)) = UCase(s_Code) Then
			Code2HTML = aButtonHTML(i)
			Exit Function
		End If
	Next
End Function

' ��ʼ��������
Sub InitToolBar()
	Dim aButton, n
	sSql = "select t_button from ewebeditor_toolbar where s_id=" & sStyleID & " order by t_order asc"
	oRs.Open sSql, oConn, 0, 1
	If Not oRs.Eof Then
		sToolBar = "<table border=0 cellpadding=0 cellspacing=0 width='100%' class='Toolbar' id='eWebEditor_Toolbar'>"
		Do While Not oRs.Eof
			sToolBar = sToolBar & "<tr><td><div class=yToolbar>"
			aButton = Split(oRs("T_Button"), "|")
			For n = 0 To UBound(aButton)
				If sFullScreen = "1" And UCase(aButton(n)) = "MAXIMIZE" Then
					aButton(n) = "Minimize"
				End If
				sToolBar = sToolBar & Code2HTML(aButton(n))
			Next
			sToolBar = sToolBar & "</div></td></tr>"
			oRs.MoveNext
		Loop
		sToolBar = sToolBar & "</table>"
	Else
		ShowErr "��Ӧ��ʽû�����ù�������"
	End If
	oRs.Close
End Sub
%>