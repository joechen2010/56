<!--#include file = "../../htmledit/eee/Include/Startup.asp"-->
<!--#include file = "Admin_Private.asp"-->
<%
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'★                                                                  ★
'☆                eWebEditor - eWebSoft在线文本编辑器               ☆
'★                                                                  ★
'☆  版权所有: eWebSoft.com                                          ☆
'★                                                                  ★
'☆  程序制作: eWeb开发团队                                          ☆
'★            email:webmaster@webasp.net                            ★
'☆            QQ:589808                                             ☆
'★                                                                  ★
'☆  相关网址: [产品介绍]http://www.eWebSoft.com/Product/eWebEditor/ ☆
'★            [支持论坛]http://bbs.eWebSoft.com/                    ★
'☆                                                                  ☆
'★  主页地址: http://www.eWebSoft.com/   eWebSoft团队及产品         ★
'☆            http://www.webasp.net/     WEB技术及应用资源网站      ☆
'★            http://bbs.webasp.net/     WEB技术交流论坛            ★
'★                                                                  ★
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
%>

<%

Dim sStyleID, sStyleName, sStyleDir, sStyleCSS, sStyleUploadDir, sStyleWidth, sStyleHeight, sStyleMemo, nStyleIsSys, sStyleStateFlag, sStyleDetectFromWord, sStyleInitMode, sStyleBaseUrl
Dim sStyleFileExt, sStyleFlashExt, sStyleImageExt, sStyleMediaExt, sStyleFileSize, sStyleFlashSize, sStyleImageSize, sStyleMediaSize
Dim sToolBarID, sToolBarName, sToolBarOrder, sToolBarButton

sPosition = sPosition & "样式管理"

If sAction = "STYLEPREVIEW" Then
	' 样式预览
	Call InitStyle()
	Call ShowStylePreview()
	Response.End
End If


Call Header()
Call Content()
Call Footer()


Sub Content()
	Select Case sAction
	Case "COPY"
		' 拷贝一标准样式
		Call InitStyle()
		Call DoCopy()
		Call ShowStyleList()
	Case "STYLEADD"
		' 新增样式表单
		Call ShowStyleForm("ADD")
	Case "STYLESET"
		' 样式设置，修改或查看
		Call InitStyle()
		Call ShowStyleForm("SET")
	Case "STYLEADDSAVE"
		' 样式新增保存
		Call CheckStyleForm()
		Call DoStyleAddSave()
	Case "STYLESETSAVE"
		' 样式设置修改保存
		Call CheckStyleForm()
		Call DoStyleSetSave()
	Case "STYLEDEL"
		' 样式删除
		Call InitStyle()
		Call DoStyleDel()
		Call ShowStyleList()
	Case "CODE"
		' 显示引用代码
		Call InitStyle()
		Call ShowStyleCode()
	Case "TOOLBAR"
		' 显示工具栏表单
		Call InitStyle()
		Call ShowToolBarList()
	Case "TOOLBARADD"
		' 新增工具栏
		Call InitStyle()
		Call DoToolBarAdd()
		Call ShowToolBarList()
	Case "TOOLBARMODI"
		' 修改工具栏
		Call InitStyle()
		Call DoToolBarModi()
		Call ShowToolBarList()
	Case "TOOLBARDEL"
		' 删除工具栏
		Call InitStyle()
		Call DoToolBarDel()
		Call ShowToolBarList()
	Case "BUTTONSET"
		' 按钮设置
		Call InitStyle()
		Call InitToolBar()
		Call ShowButtonList()
	Case "BUTTONSAVE"
		' 按钮设置保存
		Call InitStyle()
		Call InitToolBar()
		Call DoButtonSave()
	Case Else
		' 当前所有样式列表
		Call ShowStyleList()
	End Select
End Sub


' 当前所有样式列表
Sub ShowStyleList()
	Response.Write "<table border=0 cellpadding=0 cellspacing=0 width='100%' height=40><tr>" & _
		"<td class=highlight2><b>以下为当前所有样式列表：</b></td>" & _
		"<td align=right><a href='?action=styleadd'>新增样式</a></td></tr></table>"

	Response.Write "<table border=0 cellpadding=0 cellspacing=0 class=list1>" & _
		"<form action='?action=del' method=post name=myform>" & _
		"<tr align=center>" & _
			"<th width=80>样式名</th>" & _
			"<th width=60>最佳宽度</th>" & _
			"<th width=60>最佳高度</th>" & _
			"<th width=200>说明</th>" & _
			"<th width=180>管理</th>" & _
		"</tr>"

	Dim sManage
	sSql = "select * from ewebeditor_style"
	oRs.Open sSql, oConn, 0, 1
	Do While Not oRs.Eof
		sManage = "<a href='?action=stylepreview&id=" & oRs("S_ID") & "' target='_blank'>预览</a>|<a href='?action=code&id=" & oRs("S_ID") & "'>代码</a>|<a href='?action=styleset&id=" & oRs("S_ID") & "'>设置</a>|<a href='?action=toolbar&id=" & oRs("S_ID") & "'>工具栏</a>"
		If oRs("S_IsSys") = 1 Then
			sManage = sManage & "|<a href='?action=copy&id=" & oRs("S_ID") & "'>拷贝</a>"
		Else
			sManage = sManage & "|<a href='?action=styledel&id=" & oRs("S_ID") & "'>删除</a>"
		End If
		Response.Write "<tr align=center>" & _
			"<td>" & outHTML(oRs("S_Name")) & "</td>" & _
			"<td>" & oRs("S_Width") & "</td>" & _
			"<td>" & oRs("S_Height") & "</td>" & _
			"<td align=left>" & outHTML(oRs("S_Memo")) & "</td>" & _
			"<td>" & sManage & "</td>" & _
			"</tr>"
		oRs.MoveNext
	Loop
	oRs.Close
	
	Response.Write "</table>"

	Response.Write "<p class=highlight1><b>说明：</b>系统自带样式不允许对设置进行修改删除，但允许查看设置！你可以先“拷贝一标准样式”然后对其设置进行修改以达到快速新建样式的目的。</p><br><br>"

End Sub

' 拷贝一标准样式
Sub DoCopy()
	' 只有系统样式才有拷贝功能
	If nStyleIsSys <> 1 Then
		Exit Sub
	End If
	' 自动取有效的样式名，从1......到...
	Dim i, b, sNewID, sNewName
	b = False
	i = 0
	Do While b = False
		i = i + 1
		sNewName = sStyleName & i
		sSql = "select s_id from ewebeditor_style where s_name='" & sNewName & "'"
		oRs.Open sSql, oConn, 0, 1
		If oRs.Eof And oRs.Bof Then
			b = True
		End If
		oRs.Close
	Loop
	' 插入一与标准样式相同的记录，名不同
	sSql = "insert into ewebeditor_style(s_name,s_dir,s_css,s_uploaddir,s_width,s_height,s_memo,s_issys,s_fileext,s_flashext,s_imageext,s_mediaext,s_filesize,s_flashsize,s_imagesize,s_mediasize,s_stateflag,s_detectfromword,s_initmode,s_baseurl) select '" & sNewName & "',s_dir,s_css,s_uploaddir,s_width,s_height,s_memo,0,s_fileext,s_flashext,s_imageext,s_mediaext,s_filesize,s_flashsize,s_imagesize,s_mediasize,s_stateflag,s_detectfromword,s_initmode,s_baseurl from ewebeditor_style where s_id=" & sStyleID
	oConn.Execute sSql
	' 取新样式的ID
	sSql = "select s_id from ewebeditor_style where s_name='" & sNewName & "'"
	oRs.Open sSql, oConn, 0, 1
	sNewID = oRs(0)
	oRs.Close
	' 拷贝工具栏
	sSql = "insert into ewebeditor_toolbar(s_id,t_name,t_order,t_button) select " & sNewID & ",t_name,t_order,t_button from ewebeditor_toolbar where s_id=" & sStyleID
	oConn.Execute sSql
End Sub

' 样式表单
Sub ShowStyleForm(sFlag)
	Dim s_Title, s_Button, s_Action
	Dim s_FormStateFlag, s_FormDetectFromWord, s_FormInitMode, s_FormBaseUrl
	
	If sFlag = "ADD" Then
		sStyleID = ""
		sStyleName = ""
		sStyleDir = "standard"
		sStyleCSS = "office"
		sStyleUploadDir = "UploadFile"
		sStyleWidth = "600"
		sStyleHeight = "400"
		sStyleMemo = ""
		nStyleIsSys = 0
		s_Title = "新增样式"
		s_Action = "StyleAddSave"
		sStyleFileExt = "rar|zip|exe|doc|xls|chm|hlp"
		sStyleFlashExt = "swf"
		sStyleImageExt = "gif|jpg|jpeg|bmp"
		sStyleMediaExt = "rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov"
		sStyleFileSize = "500"
		sStyleFlashSize = "100"
		sStyleImageSize = "100"
		sStyleMediaSize = "100"
		sStyleStateFlag = "1"
		sStyleDetectFromWord = "true"
		sStyleInitMode = "EDIT"
		sStyleBaseUrl = "true"
	Else
		sStyleName = inHTML(sStyleName)
		sStyleDir = inHTML(sStyleDir)
		sStyleCSS = inHTML(sStyleCSS)
		sStyleUploadDir = inHTML(sStyleUploadDir)
		sStyleMemo = inHTML(sStyleMemo)
		s_Title = "设置样式"
		s_Action = "StyleSetSave"
	End If

	s_FormStateFlag = InitSelect("d_stateflag", Split("显示|不显示", "|"), Split("1|0", "|"), sStyleStateFlag, "", "")
	s_FormDetectFromWord = InitSelect("d_detectfromword", Split("自动检测有提示|不自动检测", "|"), Split("true|false", "|"), sStyleDetectFromWord, "", "")
	s_FormInitMode = InitSelect("d_initmode", Split("代码模式|编辑模式|预览模式", "|"), Split("CODE|EDIT|VIEW", "|"), sStyleInitMode, "", "")
	s_FormBaseUrl = InitSelect("d_baseurl", Split("是|否", "|"), Split("true|false", "|"), sStyleBaseUrl, "", "")

	If nStyleIsSys = 0 Then
		s_Button = "<tr><td align=center colspan=4><input type=image border=0 src='admin/submit.gif' align=absmiddle>&nbsp;<a href='javascript:document.myform.reset()'><img border=0 src='admin/reset.gif' align=absmiddle></a></td></tr>"
	Else
		s_Button = ""
	End If
	
	Response.Write "<table border=0 cellpadding=5 cellspacing=0 height=20 width='100%'><tr><td align=right><a href='javascript:history.back()'>返回</a></td></tr></table>"

	Response.Write "<table border=0 cellpadding=0 cellspacing=0 align=center class=form1>" & _
			"<form action='?action=" & s_Action & "&id=" & sStyleID & "' method=post name=myform>" & _
			"<tr><th colspan=4>&nbsp;&nbsp;" & s_Title & "（鼠标移到输入框可看说明，带*号为必填项）</th></tr>" & _
			"<tr><td width='15%' align=right>样式名称：</td><td width='35%'><input type=text class=input size=20 name=d_name title='引用此样式的名字，不要加特殊符号，最大50个字符长度' value=""" & sStyleName & """> <span class=highlight2>*</span></td><td width='15%' align=right>上传路径：</td><td width='35%'><input type=text class=input size=20 name=d_uploaddir title='存放此样式所有上传文件的路径目录名，必须使用相对路径，最后不要加斜杠，最大50个字符长度' value=""" & sStyleUploadDir & """> <span class=highlight2>*</span></td></tr>" & _
			"<tr><td width='15%' align=right>图片目录：</td><td width='35%'><input type=text class=input size=20 name=d_dir title='存放此样式图片文件的目录名，必须在ButtonImage下，最大50个字符长度' value=""" & sStyleDir & """> <span class=highlight2>*</span></td><td width='15%' align=right>样式目录：</td><td width='35%'><input type=text class=input size=20 name=d_css title='存放此样式css文件的目录名，必须在CSS下，最大50个字符长度' value=""" & sStyleCSS & """> <span class=highlight2>*</span></td></tr>" & _
			"<tr><td width='15%' align=right>最佳宽度：</td><td width='35%'><input type=text class=input name=d_width size=20 title='最佳引用效果的宽度，数字型' value='" & sStyleWidth & "'> <span class=highlight2>*</span></td><td width='15%' align=right>最佳高度：</td><td width='35%'><input type=text class=input name=d_height size=20 title='最佳引用效果的高度，数字型' value='" & sStyleHeight & "'> <span class=highlight2>*</span></td></tr>" & _
			"<tr><td width='15%' align=right>状 态 栏：</td><td width='35%'>" & s_FormStateFlag & " <span class=highlight2>*</span></td><td width='15%' align=right>Word粘贴：</td><td width='35%'>" & s_FormDetectFromWord & " <span class=highlight2>*</span></td></tr>" & _
			"<tr><td width='15%' align=right>初始模式：</td><td width='35%'>" & s_FormInitMode & " <span class=highlight2>*</span></td><td width='15%' align=right>相对路径：</td><td width='35%'>" & s_FormBaseUrl & " <span class=highlight2>*</span></td></tr>" & _
			"<tr><td colspan=4><span class=highlight2>&nbsp;&nbsp;&nbsp;允许上传文件类型及文件大小设置（文件大小单位为KB，0表示没有限制）：</span></td></tr>" & _
			"<tr><td width='15%' align=right>图片类型：</td><td width='35%'><input type=text class=input name=d_imageext size=20 title='用于图片相关的上传，最大250个字符长度' value='" & sStyleImageExt & "'></td><td width='15%' align=right>图片限制：</td><td width='35%'><input type=text class=input name=d_imagesize size=20 title='数字型，单位KB' value='" & sStyleImageSize & "'></td></tr>" & _
			"<tr><td width='15%' align=right>Flash类型：</td><td width='35%'><input type=text class=input name=d_flashext size=20 title='用于插入Flash动画，最大250个字符长度' value='" & sStyleFlashExt & "'></td><td width='15%' align=right>Flash限制：</td><td width='35%'><input type=text class=input name=d_flashsize size=20 title='数字型，单位KB' value='" & sStyleFlashSize & "'></td></tr>" & _
			"<tr><td width='15%' align=right>媒体类型：</td><td width='35%'><input type=text class=input name=d_mediaext size=20 title='用于插入媒体文件，最大250个字符长度' value='" & sStyleMediaExt & "'></td><td width='15%' align=right>媒体限制：</td><td width='35%'><input type=text class=input name=d_mediasize size=20 title='数字型，单位KB' value='" & sStyleMediaSize & "'></td></tr>" & _
			"<tr><td width='15%' align=right>其它类型：</td><td width='35%'><input type=text class=input name=d_fileext size=20 title='用于插入其它文件，最大250个字符长度' value='" & sStyleFileExt & "'></td><td width='15%' align=right>其它限制：</td><td width='35%'><input type=text class=input name=d_filesize size=20 title='数字型，单位KB' value='" & sStyleFileSize & "'></td></tr>" & _
			"<tr><td align=right>备注说明：</td><td colspan=3><textarea class=textarea rows=7 cols=65 name=d_memo title='此样式的说明，更有利于调用'>" & sStyleMemo & "</textarea></td></tr>" & s_Button & _
			"</form>" & _
			"</table><br><br>"

End Sub

' 初始化样式表数据
Sub InitStyle()
	Dim b
	b = False
	sStyleID = Trim(Request("id"))
	If IsNumeric(sStyleID) = True Then
		sSql = "select * from ewebeditor_style where s_id=" & sStyleID
		oRs.Open sSql, oConn, 0, 1
		If Not oRs.Eof Then
			sStyleName = oRs("S_Name")
			sStyleDir = oRs("S_Dir")
			sStyleCSS = oRs("S_CSS")
			sStyleUploadDir = oRs("S_UploadDir")
			sStyleWidth = CStr(oRs("S_Width"))
			sStyleHeight = CStr(oRs("S_Height"))
			sStyleMemo = oRs("S_Memo")
			nStyleIsSys = oRs("S_IsSys")
			sStyleFileExt = oRs("S_FileExt")
			sStyleFlashExt = oRs("S_FlashExt")
			sStyleImageExt = oRs("S_ImageExt")
			sStyleMediaExt = oRs("S_MediaExt")
			sStyleFileSize = oRs("S_FileSize")
			sStyleFlashSize = oRs("S_FlashSize")
			sStyleImageSize = oRs("S_ImageSize")
			sStyleMediaSize = oRs("S_MediaSize")
			sStyleStateFlag = CStr(oRs("S_StateFlag"))
			sStyleDetectFromWord = oRs("S_DetectFromWord")
			sStyleInitMode = oRs("S_InitMode")
			sStyleBaseUrl = oRs("S_BaseUrl")
			b = True
		End If
		oRs.Close
	End If
	If b = False Then
		Go_Error "无效的样式ID号，请通过页面上的链接进行操作！"
	End If
End Sub

' 检测样式表单提交的有效性
Sub CheckStyleForm()
	sStyleName = Trim(Request("d_name"))
	sStyleDir = Trim(Request("d_dir"))
	sStyleCSS = Trim(Request("d_css"))
	sStyleUploadDir = Trim(Request("d_uploaddir"))
	sStyleWidth = Trim(Request("d_width"))
	sStyleHeight = Trim(Request("d_height"))
	sStyleMemo = Request("d_memo")
	sStyleImageExt = Request("d_imageext")
	sStyleFlashExt = Request("d_flashext")
	sStyleMediaExt = Request("d_mediaext")
	sStyleFileExt = Request("d_fileext")
	sStyleImageSize = Request("d_imagesize")
	sStyleFlashSize = Request("d_flashsize")
	sStyleMediaSize = Request("d_mediasize")
	sStyleFileSize = Request("d_filesize")
	sStyleStateFlag = Request("d_stateflag")
	sStyleDetectFromWord = Request("d_detectfromword")
	sStyleInitMode = Request("d_initmode")
	sStyleBaseUrl = Request("d_baseurl")
	
	If sStyleName = "" Or Get_TrueLen(sStyleName) > 50 Then
		Go_Error "样式名不能为空，且不大于50个字符长度！"
	End If
	If IsSafeStr(sStyleName) = False Then
		Go_Error "样式名请勿包含特殊字符！"
	End If
	If sStyleDir = "" Or Get_TrueLen(sStyleDir) > 50 Then
		Go_Error "按钮图片目录名不能为空，且不大于50个字符长度！"
	End If
	If IsSafeStr(sStyleDir) = False Then
		Go_Error "按钮图片目录名请勿包含特殊字符！"
	End If
	If sStyleCSS = "" Or Get_TrueLen(sStyleCSS) > 50 Then
		Go_Error "样式CSS目录名不能为空，且不大于50个字符长度！"
	End If
	If IsSafeStr(sStyleCSS) = False Then
		Go_Error "样式CSS目录名请勿包含特殊字符！"
	End If
	If sStyleUploadDir = "" Or Get_TrueLen(sStyleUploadDir) > 50 Then
		Go_Error "上传路径目录名不能为空，且不大于50个字符长度！"
	End If
	If IsSafeStr(sStyleUploadDir) = False Then
		Go_Error "上传路径目录名请勿包含特殊字符！"
	End If
	If IsNumeric(sStyleWidth) = False Then
		Go_Error "请填写有效的最佳引用宽度！"
	End If
	If IsNumeric(sStyleHeight) = False Then
		Go_Error "请填写有效的最佳引用高度！"
	End If

	If Get_TrueLen(sStyleImageExt) > 250 Then
		Go_Error "图片文件类型不能大于250个字符长度！"
	End If
	If Get_TrueLen(sStyleFlashExt) > 250 Then
		Go_Error "Flash文件类型不能大于250个字符长度！"
	End If
	If Get_TrueLen(sStyleMediaExt) > 250 Then
		Go_Error "媒体文件类型不能大于250个字符长度！"
	End If
	If Get_TrueLen(sStyleFileExt) > 250 Then
		Go_Error "其它文件类型不能大于250个字符长度！"
	End If

	If IsNumeric(sStyleImageSize) = False Then
		Go_Error "请填写有效的图片限制大小！"
	End If
	If IsNumeric(sStyleFlashSize) = False Then
		Go_Error "请填写有效的Flash限制大小！"
	End If
	If IsNumeric(sStyleMediaSize) = False Then
		Go_Error "请填写有效的媒体文件限制大小！"
	End If
	If IsNumeric(sStyleFileSize) = False Then
		Go_Error "请填写有效的其它文件限制大小！"
	End If

End Sub

' 样式新增保存
Sub DoStyleAddSave()
	sSql = "select * from ewebeditor_style where s_name='" & sStyleName & "'"
	oRs.Open sSql, oConn, 0, 1
	If Not oRs.Eof Then
		Go_Error "此样式名已经存在，请用另一个样式名！"
	End If
	oRs.Close

	sSql = "select * from ewebeditor_style where s_id=0"
	oRs.Open sSql, oConn, 1, 3
	oRs.AddNew
	oRs("S_Name") = sStyleName
	oRs("S_Dir") = sStyleDir
	oRs("S_CSS") = sStyleCSS
	oRs("S_UploadDir") = sStyleUploadDir
	oRs("S_Width") = sStyleWidth
	oRs("S_Height") = sStyleHeight
	oRs("S_Memo") = sStyleMemo
	oRs("S_ImageExt") = sStyleImageExt
	oRs("S_FlashExt") = sStyleFlashExt
	oRs("S_MediaExt") = sStyleMediaExt
	oRs("S_FileExt") = sStyleFileExt
	oRs("S_ImageSize") = sStyleImageSize
	oRs("S_FlashSize") = sStyleFlashSize
	oRs("S_MediaSize") = sStyleMediaSize
	oRs("S_FileSize") = sStyleFileSize
	oRs("S_StateFlag") = sStyleStateFlag
	oRs("S_DetectFromWord") = sStyleDetectFromWord
	oRs("S_InitMode") = sStyleInitMode
	oRs("S_BaseUrl") = sStyleBaseUrl
	oRs.Update
	sStyleID = oRs("S_ID")
	oRs.Close

	Response.Write "<br><table border=0 cellspacing=20 align=center>" & _
		"<tr valign=top><td><img src='admin/do_ok.gif' border=0></td><td><b><span class=highlight2>样式增加成功！</span></b><br><br><ul>我现在<br><br><li><a href='admin_default.asp'>返回后台管理首页</a><li><a href='?'>返回样式管理</a><li><a href='?action=toolbar&id=" & sStyleID & "'>设置此样式下的工具栏</a></ul></td></tr>" & _
		"</table><br><br>"

End Sub


' 样式修改保存
Sub DoStyleSetSave()
	sStyleID = Trim(Request("id"))
	If IsNumeric(sStyleID) = True Then
		' 是否存在同名
		sSql = "select * from ewebeditor_style where s_name='" & sStyleName & "' and s_id<>" & sStyleID
		oRs.Open sSql, oConn, 0, 1
		If Not oRs.Eof Then
			Go_Error "此样式名已经存在，请用另一个样式名！"
		End If
		oRs.Close
		
		sSql = "select * from ewebeditor_style where s_id=" & sStyleID
		oRs.Open sSql, oConn, 1, 3
		If Not oRs.Eof Then
			If oRs("S_IsSys") = 1 Then
				Go_Error "系统样式，不允许修改！"
			End If
		Else
			Go_Error "无效的样式ID号，请通过页面上的链接进行操作！"
		End If
		oRs("S_Name") = sStyleName
		oRs("S_Dir") = sStyleDir
		oRs("S_CSS") = sStyleCSS
		oRs("S_UploadDir") = sStyleUploadDir
		oRs("S_Width") = sStyleWidth
		oRs("S_Height") = sStyleHeight
		oRs("S_Memo") = sStyleMemo
		oRs("S_ImageExt") = sStyleImageExt
		oRs("S_FlashExt") = sStyleFlashExt
		oRs("S_MediaExt") = sStyleMediaExt
		oRs("S_FileExt") = sStyleFileExt
		oRs("S_ImageSize") = sStyleImageSize
		oRs("S_FlashSize") = sStyleFlashSize
		oRs("S_MediaSize") = sStyleMediaSize
		oRs("S_FileSize") = sStyleFileSize
		oRs("S_StateFlag") = sStyleStateFlag
		oRs("S_DetectFromWord") = sStyleDetectFromWord
		oRs("S_InitMode") = sStyleInitMode
		oRs("S_BaseUrl") = sStyleBaseUrl
		oRs.Update
		oRs.Close
	Else
		Go_Error "无效的样式ID号，请通过页面上的链接进行操作！"
	End If

	Response.Write "<br><table border=0 cellspacing=20 align=center>" & _
		"<tr valign=top><td><img src='admin/do_ok.gif' border=0></td><td><b><span class=highlight2>样式修改成功！</span></b><br><br><ul>我现在<br><br><li><a href='admin_default.asp'>返回后台管理首页</a><li><a href='?'>返回样式管理</a><li><a href='?action=toolbar&id=" & sStyleID & "'>设置此样式下的工具栏</a></ul></td></tr>" & _
		"</table><br><br>"

End Sub

' 样式删除
Sub DoStyleDel()
	If nStyleIsSys = 0 Then
		sSql = "delete from ewebeditor_style where s_id=" & sStyleID
		oConn.Execute sSql
		sSql = "delete from ewebeditor_toolbar where s_id=" & sStyleID
		oConn.Execute sSql
	End If
End Sub

' 样式预览
Sub ShowStylePreview()
	Response.Write "<html><head>" & _
		"<title>样式预览</title>" & _
		"<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & _
		"</head><body>" & _
		"<input type=hidden name=content1  value=''>" & _
		"<iframe ID='eWebEditor1' src='ewebeditor.asp?id=content1&style=" & sStyleName & "' frameborder=0 scrolling=no width='" & sStyleWidth & "' HEIGHT='" & sStyleHeight & "'></iframe>" & _
		"</body></html>"
End Sub

' 显示引用代码
Sub ShowStyleCode()
	Response.Write "<table border=0 cellpadding=5 cellspacing=0 height=20 width='100%'><tr><td align=right><a href='javascript:history.back()'>返回</a></td></tr></table>"
	Response.Write "<p class=highlight2>样式（<b>" & outHTML(sStyleName) & "</b>）的最佳调用代码如下（其中XXX按实际关联的表单项进行修改）：</p>"
	Response.Write "<textarea rows=5 cols=65 style='width:100%'><IFRAME ID=""eWebEditor1"" SRC=""ewebeditor.asp?id=XXX&style=" & sStyleName & """ FRAMEBORDER=""0"" SCROLLING=""no"" WIDTH=""" & sStyleWidth & """ HEIGHT=""" & sStyleHeight & """></IFRAME></textarea><br><br>"
End Sub

' 显示工具栏表单列表
Sub ShowToolBarList()
	Response.Write "<table border=0 cellpadding=5 cellspacing=0 height=20 width='100%'><tr>" & _
		"<td class=highlight2><b>样式（" & outHTML(sStyleName) & "）下的工具栏管理：</b></td>" & _
		"<td align=right><a href='?'>返回样式管理</a></td></tr></table>"

	Dim s_AddForm, s_ModiForm

	' 增加表单
	If nStyleIsSys = 1 Then
		s_AddForm = ""
	Else
		' 取当前最大排序号
		Dim nMaxOrder
		sSql = "select max(T_order) from ewebeditor_toolbar where s_id=" & sStyleID
		oRs.Open sSql, oConn, 0, 1
		If IsNull(oRs(0)) Then
			nMaxOrder = 1
		Else
			nMaxOrder = oRs(0) + 1
		End If
		oRs.Close

		s_AddForm = "<hr width='95%' align=center size=1><table border=0 cellpadding=4 cellspacing=0 align=center>" & _
		"<form action='?id=" & sStyleID & "&action=toolbaradd' name='addform' method=post>" & _
		"<tr><td>工具栏名：<input type=text name=d_name size=20 class=input value='工具栏" & nMaxOrder & "'> 排序号：<input type=text name=d_order size=5 value='" & nMaxOrder & "' class=input> <input type=submit name=b1 value='新增工具栏'></td></tr>" & _
		"</form></table><hr width='95%' align=center size=1>"
	End If

	' 修改表单
	Dim s_Manage, s_SubmitButton
	s_ModiForm = "<form action='?id=" & sStyleID & "&action=toolbarmodi' name=modiform method=post>" & _
		"<table border=0 cellpadding=0 cellspacing=0 align=center class=list1>" & _
		"<tr align=center><th>ID</th><th>工具栏名</th><th>排序号</th><th>操作</th></tr>"
	sSql = "select * from ewebeditor_toolbar where s_id=" & sStyleID & " order by t_order asc"
	oRs.Open sSql, oConn, 0, 1
	Do While Not oRs.Eof
		s_Manage = "<a href='?id=" & sStyleID & "&action=buttonset&toolbarid=" & oRs("t_id") & "'>按钮设置</a>"
		If nStyleIsSys <> 1 Then
			s_Manage = s_Manage & "|<a href='?id=" & sStyleID & "&action=toolbardel&delid=" & oRs("t_id") & "'>删除</a>"
		End If
		s_ModiForm = s_ModiForm & "<tr align=center>" & _
			"<td>" & oRs("t_id") & "</td>" & _
			"<td><input type=text name='d_name" & oRs("t_id") & "' value=""" & inHTML(oRs("t_name")) & """ size=30 class=input></td>" & _
			"<td><input type=text name='d_order" & oRs("t_id") & "' value='" & oRs("t_order") & "' size=5 class=input></td>" & _
			"<td>" & s_Manage & "</td>" & _
			"</tr>"
		oRs.MoveNext
	Loop
	oRs.Close
	If nStyleIsSys = 1 Then
		s_SubmitButton = ""
	Else
		s_SubmitButton = "<p align=center><input type=submit name=b1 value=' 修改 '></p>"
	End If
	s_ModiForm = s_ModiForm & "</table>" & s_SubmitButton & "</form><br><br>"

	' 输出表单
	Response.Write s_AddForm & s_ModiForm

End Sub

' 新增工具栏
Sub DoToolBarAdd()
	Dim s_Name, s_Order
	s_Name = Trim(Request("d_name"))
	s_Order = Trim(Request("d_order"))
	If s_Name = "" Or Get_TrueLen(s_Name) > 50 Then
		Go_Error "工具栏名不能为空，且长度不能大于50个字符长度！"
	End If
	If IsNumeric(s_Order) = False Then
		Go_Error "无效的工具栏排序号，排序号必须为数字！"
	End If
	If nStyleIsSys = 1 Then
		Go_Error "系统自带样式下的工具栏，不允许新增！"
	End If
	sSql = "select * from ewebeditor_toolbar where 1=0"
	oRs.Open sSql, oConn, 1, 3
	oRs.AddNew
	oRs("s_id") = sStyleID
	oRs("t_name") = s_Name
	oRs("t_order") = s_Order
	oRs("t_button") = ""
	oRs.Update
	oRs.Close
	Response.Write "<script language=javascript>alert(""工具栏（" & outHTML(s_Name) & "）增加操作成功！"");</script>"
End Sub

' 修改工具栏
Sub DoToolBarModi()
	Dim s_Name, s_Order
	If nStyleIsSys = 1 Then
		Go_Error "系统自带样式下的工具栏，不允许修改！"
	End If

	sSql = "select * from ewebeditor_toolbar where s_id=" & sStyleID
	oRs.Open sSql, oConn, 1, 3
	Do While Not oRs.Eof
		s_Name = Trim(Request("d_name" & oRs("t_id")))
		s_Order = Trim(Request("d_order" & oRs("t_id")))
		If s_Name <> "" And IsNumeric(s_Order) = True Then
			If s_Name <> oRs("t_name") Or s_Order <> CStr(oRs("t_Order")) Then
				oRs("t_name") = s_Name
				oRs("t_order") = s_Order
				oRs.Update
			End If
		End If
		oRs.MoveNext
	Loop
	oRs.Close
	Response.Write "<script language=javascript>alert('工具栏修改操作成功！');</script>"

End Sub

' 删除工具栏
Sub DoToolBarDel()
	Dim s_DelID
	s_DelID = Trim(Request("delid"))
	If nStyleIsSys = 1 Then
		Go_Error "系统自带样式下的工具栏，不允许删除！"
	End If
	If IsNumeric(s_DelID) = True Then
		sSql = "delete from ewebeditor_toolbar where s_id=" & sStyleID & " and t_id=" & s_DelID
		oConn.Execute sSql
		Response.Write "<script language=javascript>alert('工具栏（ID：" & s_DelID & "）删除操作成功！');</script>"
	End If
End Sub

' 初始化工具栏
Sub InitToolBar()
	Dim b
	b = False
	sToolBarID = Trim(Request("toolbarid"))
	If IsNumeric(sToolBarID) = True Then
		sSql = "select * from ewebeditor_toolbar where s_id=" & sStyleID & " and t_id=" & sToolBarID
		oRs.Open sSql, oConn, 0, 1
		If Not oRs.Eof Then
			sToolBarName = oRs("T_Name")
			sToolBarOrder = oRs("T_Order")
			sToolBarButton = oRs("T_Button")
			b = True
		End If
		oRs.Close
	End If
	If b = False Then
		Go_Error "无效的工具栏ID号，请通过页面上的链接进行操作！"
	End If
End Sub

' 按钮设置
Sub ShowButtonList()
	Dim i, n

	' 导航
	Response.Write "<table border=0 cellpadding=5 cellspacing=0 height=30 width='100%'><tr>" & _
		"<td><b>当前样式：<span class=highlight2>" & outHTML(sStyleName) & "</span>&nbsp;&nbsp;当前工具栏：<span class=highlight2>" & outHTML(sToolBarName) & "</span></b></td>" & _
		"<td align=right><a href='?action=toolbar&id=" & sStyleID & "'>返回工具栏管理</a> | <a href='?'>返回样式管理</a></td>" & _
		"</tr></table>"

	' 取所有按钮
	Dim aButtonCode(), aButtonTitle(), aButtonImage()
	sSql = "select * from ewebeditor_button where b_allowselect=1 order by b_order asc"
	oRs.Open sSql, oConn, 0, 1
	i = 0
	Do While Not oRs.Eof
		i = i + 1
		Redim Preserve aButtonCode(i)
		Redim Preserve aButtonTitle(i)
		Redim Preserve aButtonImage(i)
		aButtonCode(i) = oRs("B_Code")
		aButtonTitle(i) = oRs("B_Title")
		aButtonImage(i) = oRs("B_Image")
		oRs.MoveNext
	Loop
	oRs.Close


	' 取可选列表
	Dim s_Option1
	s_Option1 = ""
	For i = 1 To UBound(aButtonCode)
		s_Option1 = s_Option1 & "<option value='" & aButtonCode(i) & "'>" & aButtonTitle(i) & "</option>"
	Next

	' 取已选列表
	Dim aButton, s_Option2, s_Temp
	aButton = Split(sToolBarButton, "|")
	s_Option2 = ""
	For i = 0 To UBound(aButton)
		s_Temp = Code2Title(aButton(i), aButtonCode, aButtonTitle)
		If s_Temp <> "" Then
			s_Option2 = s_Option2 & "<option value='" & aButton(i) & "'>" & s_Temp & "</option>"
		End If
	Next


	'以下为客户端操作选择脚本
	'''''''''''''''''''''''''''''''''''
%>

<script language=javascript>
// 加入已选
function Add() {
	var sel1=document.myform.d_b1;
	var sel2=document.myform.d_b2;
	if (sel1.selectedIndex<0) {
		alert("请选择一个待选按钮！");
		return;
	}
	sel2.options[sel2.length]=new Option(sel1.options[sel1.selectedIndex].innerHTML,sel1.options[sel1.selectedIndex].value);
}

// 从已选中删除
function Del() {
	var sel=document.myform.d_b2;
	var nIndex = sel.selectedIndex;
	var nLen = sel.length;
	if (nLen<1) return;
	if (nIndex<0) {
		alert("请选择一个已选按钮！");
		return;
	}
	for (var i=nIndex;i<nLen-1;i++) {
		sel.options[i].value=sel.options[i+1].value;
		sel.options[i].innerHTML=sel.options[i+1].innerHTML;
	}
	sel.length=nLen-1;
}

// 排序：向上移动
function Up() {
	var sel=document.myform.d_b2;
	var nIndex = sel.selectedIndex;
	var nLen = sel.length;
	if ((nLen<1)||(nIndex==0)) return;
	if (nIndex<0) {
		alert("请选择一个要移动的已选按钮！");
		return;
	}
	var sValue=sel.options[nIndex].value;
	var sHTML=sel.options[nIndex].innerHTML;
	sel.options[nIndex].value=sel.options[nIndex-1].value;
	sel.options[nIndex].innerHTML=sel.options[nIndex-1].innerHTML;
	sel.options[nIndex-1].value=sValue;
	sel.options[nIndex-1].innerHTML=sHTML;
	sel.selectedIndex=nIndex-1;
}

// 排序：向下移动
function Down() {
	var sel=document.myform.d_b2;
	var nIndex = sel.selectedIndex;
	var nLen = sel.length;
	if ((nLen<1)||(nIndex==nLen-1)) return;
	if (nIndex<0) {
		alert("请选择一个要移动的已选按钮！");
		return;
	}
	var sValue=sel.options[nIndex].value;
	var sHTML=sel.options[nIndex].innerHTML;
	sel.options[nIndex].value=sel.options[nIndex+1].value;
	sel.options[nIndex].innerHTML=sel.options[nIndex+1].innerHTML;
	sel.options[nIndex+1].value=sValue;
	sel.options[nIndex+1].innerHTML=sHTML;
	sel.selectedIndex=nIndex+1;
}

// 提交处理
function checkform() {
	var sel=document.myform.d_b2;
	var nLen = sel.length;
	var str="";
	for (var i=0;i<nLen;i++) {
		if (i>0) str+="|";
		str+=sel.options[i].value;
	}
	document.myform.d_button.value=str;
	return true;
}

</script>

<%
	'''''''''''''''''''''''''''''''''''


	' 选择设置表单
	Dim s_SubmitButton
	If nStyleIsSys = 1 Then
		s_SubmitButton = ""
	Else
		s_SubmitButton = "<input type=submit name=b value=' 保存设置 '>"
	End If
	Response.Write "<table border=0 cellpadding=5 cellspacing=0 align=center>" & _
		"<form action='?action=buttonsave&id=" & sStyleID & "&toolbarid=" & sToolBarID & "' method=post name=myform onsubmit='return checkform()'>" & _
		"<tr align=center><td>可选按钮</td><td></td><td>已选按钮</td><td></td></tr>" & _
		"<tr align=center>" & _
			"<td><select name='d_b1' size=20 style='width:200px' ondblclick='Add()'>" & s_Option1 & "</select></td>" & _
			"<td><input type=button name=b1 value=' → ' onclick='Add()'><br><br><input type=button name=b1 value=' ← ' onclick='Del()'></td>" & _
			"<td><select name='d_b2' size=20 style='width:200px' ondblclick='Del()'>" & s_Option2 & "</select></td>" & _
			"<td><input type=button name=b3 value='↑' onclick='Up()'><br><br><br><input type=button name=b4 value='↓' onclick='Down()'></td>" & _
		"</tr>" & _
		"<input type=hidden name='d_button' value=''>" & _
		"<tr><td colspan=4 align=right>" & s_SubmitButton & "</td></tr>" & _
		"</form></table>"


	' 显示图片对照表
	Response.Write "<p class=highlight1><b>以下是按钮图片对照表（部分下拉框或特殊按钮可能没图）：</b></p>"
	Response.Write "<table border=1 cellpadding=3 cellspacing=0 width='100%'>"
	n = 0
	For i = 1 To UBound(aButtonCode)
		n = i Mod 4
		If n = 1 Then
			Response.Write "<tr>"
		End If
		Response.Write "<td>"
		If aButtonImage(i) <> "" Then
			Response.Write "<img border=0 align=absmiddle src='ButtonImage/standard/" & aButtonImage(i) & "'>"
		End If
		Response.Write aButtonTitle(i)
		Response.Write "</td>"
		If n = 0 Then
			Response.Write "</tr>"
		End If
	Next
	If n > 0 Then
		For i = 1 To 4 - n
			Response.Write "<td>&nbsp;</td>"
		Next
		Response.Write "</tr>"
	End if
	Response.Write "</table><br><br>"
End Sub

' 由按钮代码得到按钮标题
Function Code2Title(s_Code, a_ButtonCode, a_ButtonTitle)
	Dim i
	Code2Title = ""
	For i = 1 To UBound(a_ButtonCode)
		If UCase(a_ButtonCode(i)) = UCase(s_Code) Then
			Code2Title = a_ButtonTitle(i)
			Exit Function
		End If
	Next
End Function

' 按钮设置保存
Sub DoButtonSave()
	Dim s_Button
	s_Button = Trim(Request("d_button"))
	If nStyleIsSys = 1 Then
		Go_Error "系统自带样式，不允许对按钮进行修改！"
	End If

	sSql = "select * from ewebeditor_toolbar where t_id=" & sToolBarID
	oRs.Open sSql, oConn, 1, 3
	If Not oRs.Eof Then
		oRs("T_Button") = s_Button
		oRs.Update
	End If
	oRs.Close

	Response.Write "<br><table border=0 cellspacing=20 align=center>" & _
		"<tr valign=top><td><img src='admin/do_ok.gif' border=0></td><td><b><span class=highlight2>工具栏按钮设置保存成功！</span></b><br><br><ul>我现在<br><br><li><a href='admin_default.asp'>返回后台管理首页</a><li><a href='?'>返回样式管理</a><li><a href='?action=toolbar&id=" & sStyleID & "'>返回工具栏管理</a><li><a href='?action=buttonset&id=" & sStyleID & "&toolbarid=" & sToolBarID & "'>重新设置此工具栏下的按钮</a></ul></td></tr>" & _
		"</table><br><br>"

End Sub

%>