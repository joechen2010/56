<% 
   dim companyname,ss,gsid,rs_cc
   if session("user")="" then
	   response.write "<script>top.location.href='../login/login.asp';</script>"
	   response.end
   end if
   
   if session("hydj")=0 then
	  ss="��ѻ�Ա"
   elseif session("hydj")=1 then
	  ss="ͭ�ƻ�Ա"
   elseif session("hydj")=2 then
	  ss="���ƻ�Ա"
   elseif session("hydj")=3 then
	  ss="���ƻ�Ա"
   end if
   
   set rs_cc=conn.execute("select qymc,id from qyml where user='"&session("user")&"'")
   companyname=rs_cc(0)
   gsid=rs_cc(1)
   rs_cc.close
   set rs_cc=nothing
%>