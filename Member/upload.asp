<%
dim stype
stype=request("stype")
if stype="info" then
    uploadpath="../InfoUpload/"
elseif stype="Pro" then
   uploadpath="../ProUpload/"
else
'
end if
%>
<html>
<head>
<title>�ļ��ϴ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/cssup.css" rel="stylesheet" type="text/css">
</head>
<body style="background:url(); background-color:#EDF6FF">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="20">
      <form name="form" method="post" action="upfile.asp" enctype="multipart/form-data" >
        <input type="hidden" name="filepath" value="<%=uploadpath%>">
<input type="hidden" name="act" value="upload">
        ѡ��ͼƬ�� 
        <input type="file" name="file1" size="25">
<input type="submit" name="Submit" class="iptBtnYellow" value="�ϴ�">
</form></td>
  </tr>
</table>
</body>
</html>
