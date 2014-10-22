<%
if session("username")<>"" then

        Set Rs=server.createobject("adodb.recordset")
		Sql="select * from  clkj_admin  where clkj_password='"&request.cookies("userpas")("upas")&"'"
		Rs.open sql,conn,1,1
		if not rs.eof then
      	 	session("username")=request.cookies("username")("uname")
	   end if
else
     Response.Write "<script language='javascript'>alert('用户名与密码为空或失效请重新进入!');top.location.href='!Index.html';</script>"

end if

%>
