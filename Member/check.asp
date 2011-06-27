<% 
   dim companyname,ss,gsid,rs_cc
   if session("user")="" then
	   response.write "<script>top.location.href='../login/login.asp';</script>"
	   response.end
   end if
   
   if session("hydj")=0 then
	  ss="免费会员"
   elseif session("hydj")=1 then
	  ss="铜牌会员"
   elseif session("hydj")=2 then
	  ss="银牌会员"
   elseif session("hydj")=3 then
	  ss="金牌会员"
   end if
   
   set rs_cc=conn.execute("select qymc,id from qyml where user='"&session("user")&"'")
   companyname=rs_cc(0)
   gsid=rs_cc(1)
   rs_cc.close
   set rs_cc=nothing
%>