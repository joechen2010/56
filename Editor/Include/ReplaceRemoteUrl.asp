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

'================================================
'��  �ã��滻�ַ����е�Զ���ļ�Ϊ�����ļ�������Զ���ļ�
'��  ����
'	sHTML		: Ҫ�滻���ַ���
'	sSavePath	: �����ļ���·��
'	sExt		: ִ���滻����չ��
'================================================
Function eWebEditor_ReplaceRemoteUrl(sHTML, sSavePath, sExt)
	Dim s_Content
	s_Content = sHTML
	If eWebEditor_IsObjInstalled("Microsoft.XMLHTTP") = False then
		eWebEditor_ReplaceRemoteUrl = s_Content
		Exit Function
	End If
	
	If sSavePath = "" Then sSavePath = "/eWebEditor/UploadFile/"
	If sExt = "" Then sExt = ""jpg|gif|bmp|png""
	Dim re, RemoteFile, RemoteFileurl, SaveFileName, SaveFileType, ranNum
	Set re = new RegExp
	re.IgnoreCase  = True
	re.Global = True
	re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(" & sExt & ")))"
	Set RemoteFile = re.Execute(s_Content)
	For Each RemoteFileurl in RemoteFile
		SaveFileType = Mid(RemoteFileurl, InstrRev(RemoteFileurl, ".") + 1)
		Randomize
		ranNum = Int(900 * Rnd) + 100
		SaveFileName = sSavePath & year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & ranNum & "." & SaveFileType
		Call eWebEditor_SaveRemoteFile(SaveFileName, RemoteFileurl)
		s_Content = Replace(s_Content,RemoteFileurl,SaveFileName)
	Next
	eWebEditor_ReplaceRemoteUrl = s_Content
End Function

'================================================
'��  �ã�����Զ�̵��ļ�������
'��  ����LocalFileName ------ �����ļ���
'		 RemoteFileUrl ------ Զ���ļ�URL
'����ֵ��True  ----�ɹ�
'        False ----ʧ��
'================================================
Sub eWebEditor_SaveRemoteFile(s_LocalFileName,s_RemoteFileUrl)
	Dim Ads, Retrieval, GetRemoteData
	On Error Resume Next
	Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
	With Retrieval
		.Open "Get", s_RemoteFileUrl, False, "", ""
		.Send
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing
	Set Ads = Server.CreateObject("Adodb.Stream")
	With Ads
		.Type = 1
		.Open
		.Write GetRemoteData
		.SaveToFile Server.MapPath(s_LocalFileName), 2
		.Cancel()
		.Close()
	End With
	Set Ads=nothing
End Sub

'================================================
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'        False ----û�а�װ
'================================================
Function eWebEditor_IsObjInstalled(s_ClassString)
	On Error Resume Next
	eWebEditor_IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(s_ClassString)
	If 0 = Err Then eWebEditor_IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>