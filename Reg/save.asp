<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<%
dim city,city2,cartype,Price,SortID,TypeID,Nowplace,MadePlace,sIntroduce,PicImg

dim sql,i,qyjj
     qyjj=request.form("qyjj")
	 RegUserName=request.form("user")
     Password=request.form("pass")
     set rs=server.CreateObject ("adodb.recordset")
      sql="select * from qyml where user='"&RegUserName&"'"
      rs.open sql,conn,1,3
      if not (rs.eof and rs.bof) then
		   response.Write "<br><li>你注册的用户已经存在！请换一个用户名再试试！</li>"
      else 
rs.addnew
          rs("user")=RegUserName
          rs("pass")=Password
          'rs("question")=question
          'rs("answer")=answer
          rs("name")=request("sname")
          rs("ch")=request("ch")
          rs("Province")=request("Province")
          rs("city")=request("city")
		  rs("area")=request("area")
          rs("address")=request("address")
          rs("post")=request("post")
          rs("zw")=request("zw")
          rs("phone1")=request("phone1")
		  rs("phone2")=request("phone2")
		  rs("phone3")=request("phone3")
          rs("mobile")=request("mobile")
          rs("fax1")=request("fax1")
		  rs("fax2")=request("fax2")
		  rs("fax3")=request("fax3")
          rs("email")=request("email")
          rs("qymc")=request("qymc")
          rs("web")=request("web")
          'rs("qylb")=qylb
          rs("jygs")=request("jyfw")
          rs("sortid")=request("sortid")
          rs("idate")=date()
          rs("Flag")=0
          rs("uflag")=0
          rs("click")=0
          'rs("Qyxz")=Qyxz
          rs("qyjj")=qyjj
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('注册成功！');top.location.href='../login/login.asp';</script>"
	response.end
end if
%>