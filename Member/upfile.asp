<!--#include file="upload.inc"-->
<html>
<head>
<title></title>
<link rel="stylesheet" href="images/cssup.css" type="text/css">
</head>
<body bgcolor="#EDF6FF">
<script>
parent.document.forms[0].ToSubmit.disabled=false;
parent.document.forms[0].ToBack.disabled=false;
</script>
<%

dim upload,file,formName,formPath,iCount,filename,fileExt
set upload=new upload_5xSoft ''�����ϴ�����


 formPath=upload.form("filepath")
 ''��Ŀ¼���(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 
 
 iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<100 then
 	response.write "<font size=2>����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>500000 then
 	response.write "<font size=2>�ļ���С���������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if ( fileEXT<>".gif" and fileEXT<>".jpg" and fileEXT<>".avi" and fileEXT<>".swf") then
 	response.write "<font size=2>�ļ���ʽ���ԡ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if 
 Randomize
	RanNum = Int(90000*rnd)+10000
 filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&RanNum&FileExt
 
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  file.SaveAs Server.mappath(formPath&filename)   ''�����ļ�
'  response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" �ɹ�!<br>"
'response.write filename
'response.end
 response.write "<script>parent.document.forms[0].ProImg.value='"&filename&"'</script>"
  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''ɾ���˶���

session("upface_Bill")="done"

Htmend iCount&" ���ļ��ϴ�����!"

sub HtmEnd(Msg)
 set upload=nothing
%>
<SCRIPT LANGUAGE="JavaScript"> alert("�ļ��ϴ��ɹ�,��copy�±ߵ��ļ�����,�Ա�����") </script>
<%
  response.write "�ļ��ϴ��ɹ�!"
  response.end
end sub


%>
</body>
</html>