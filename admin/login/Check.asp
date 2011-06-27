<%
if session("AdminName")="" or session("AdminPower")="" then
	response.write "<script>alert('你还没登陆.或者登陆超时.请重新登陆.');top.location.href='../Index.asp';</script>"
	response.end
end if
%>