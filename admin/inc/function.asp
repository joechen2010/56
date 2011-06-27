<%
' ********************************************
' 以下为常用函数
' ********************************************
' ============================================
' 错误信息返回
' 参数:str 返回的错误信息
' ============================================
function ShowError(str)
	Response.Write "<script language=javascript>alert('" & str & "\n\n系统将自动返回前一页面...');history.back();</script>"
	'Response.End
End function

' ============================================
' 成功信息返回
' 参数:str 返回的错误信息
'      goUrl 返回的页面
' ============================================
function ShowSuccess(str,goUrl)
	Response.Write "<script language=javascript>alert('" & str & "');location.href='"&goUrl&"';</script>"
	'Response.End
End function

' ============================================
' 得到安全字符串,在查询中或有必要强行替换的表单中使用
' ============================================
Function GetSafeStr(str)
	GetSafeStr = Replace(Replace(Replace(Trim(str), "'", ""), Chr(34), ""), ";", "")
End Function

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
' ============================================
' 用SafeRequet代替Request
' ============================================
Function SafeRequest(ParaName,ParaType)
'1=数字型，0=文本型
	Dim ParaValue
	ParaValue=Request(ParaName)
	If ParaType=1 Then
		If Not isNumeric(ParaValue) Then
            call ShowError ("对不起，请正常访问该网页！") 
		End If
	Else
		ParaValue=Replace(ParaValue,"'","''")
		ParaValue=Replace(ParaValue,";","；")
		ParaValue=Replace(ParaValue,"select","Select",1,-1,1)
	End if
	SafeRequest=ParaValue
End function
%>