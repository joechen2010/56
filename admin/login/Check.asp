<%
if session("AdminName")="" or session("AdminPower")="" then
	response.write "<script>alert('�㻹û��½.���ߵ�½��ʱ.�����µ�½.');top.location.href='../Index.asp';</script>"
	response.end
end if
%>