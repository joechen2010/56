<%
' ********************************************
' ����Ϊ���ú���
' ********************************************
' ============================================
' ������Ϣ����
' ����:str ���صĴ�����Ϣ
' ============================================
function ShowError(str)
	Response.Write "<script language=javascript>alert('" & str & "\n\nϵͳ���Զ�����ǰһҳ��...');history.back();</script>"
	'Response.End
End function

' ============================================
' �ɹ���Ϣ����
' ����:str ���صĴ�����Ϣ
'      goUrl ���ص�ҳ��
' ============================================
function ShowSuccess(str,goUrl)
	Response.Write "<script language=javascript>alert('" & str & "');location.href='"&goUrl&"';</script>"
	'Response.End
End function

' ============================================
' �õ���ȫ�ַ���,�ڲ�ѯ�л��б�Ҫǿ���滻�ı���ʹ��
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
' ��SafeRequet����Request
' ============================================
Function SafeRequest(ParaName,ParaType)
'1=�����ͣ�0=�ı���
	Dim ParaValue
	ParaValue=Request(ParaName)
	If ParaType=1 Then
		If Not isNumeric(ParaValue) Then
            call ShowError ("�Բ������������ʸ���ҳ��") 
		End If
	Else
		ParaValue=Replace(ParaValue,"'","''")
		ParaValue=Replace(ParaValue,";","��")
		ParaValue=Replace(ParaValue,"select","Select",1,-1,1)
	End if
	SafeRequest=ParaValue
End function
%>