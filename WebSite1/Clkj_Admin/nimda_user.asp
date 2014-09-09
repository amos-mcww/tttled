<!--#include file="Nimda_function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>用户管理</title>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg"><strong>&nbsp;用户管理</strong> <font color="#ff0000"><%=request.querystring("Lei")%></font></td>
  </tr>

</table>
<% If request.querystring("Edit")="B_E" Then 
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_admin where clkj_id ="&request("clkj_id")
	Rs.open Sql,conn,1,1
%>
<table width="100%" cellpadding="0" cellspacing="0">
    <tr>
    <td height="25" bgcolor="#F5F5F5" style="text-indent:2em;color:#FF0000;" id="zx" colspan="2">修改用户</td>
  </tr>
 <form name="pas" action="nimda_function.asp?Class=pas_Edit" method="post">
  <tr>
    <td height="30" align="center" valign="middle">用户权限</td>
    <td><select name="clkj_flg">
    <%if rs("clkj_flg")=1 then%>
      <option value="1" <%if rs("clkj_flg")=1 then%>selected<%end if%>>管理员</option>
      <option value="0" <%if rs("clkj_flg")=0 then%>selected<%end if%>>普通员</option>
      <%else%>
      <option value="0" <%if rs("clkj_flg")=0 then%>selected<%end if%>>普通员</option>
      <%end if%>
    </select>
    </td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle">用户名</td>
    <td><input name="user_name" type="text" id="user_name" value="<%=rs("clkj_admin")%>" size="20"></td>
  </tr>
    <tr>
    <td width="15%" height="30" align="center" valign="middle">密码</td>
    <td><input name="user_pas" type="text" id="user_pas"  size="20">
    密码为空不修改</td>
  </tr>
    <tr>
    <td width="15%" height="30"></td>
    <td><input type="submit" name="Submit" value="修改"><input type="hidden" name="clkj_id" value="<%=request("clkj_id")%>" /></td>
  </tr>
  </form>
</table>
<%Else%>
<table width="100%" cellpadding="0" cellspacing="0">
    <tr>
    <td height="25" bgcolor="#F5F5F5" style="text-indent:2em;color:#FF0000;" id="zj" colspan="2">添加用户</td>
  </tr>
 <form name="pas" action="nimda_function.asp?Class=pas" method="post">
  <tr>
    <td height="30" align="center" valign="middle">用户权限</td>
    <td><select name="clkj_flg">
      <option value="1" selected>管理员</option>
      <option value="0">普通员</option>
    </select>
    </td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle">用户名</td>
    <td><input name="user_name" type="text" id="user_name" size="20"></td>
  </tr>
    <tr>
    <td width="15%" height="30" align="center" valign="middle">密码</td>
    <td><input name="user_pas" type="text" id="user_pas" size="20"></td>
  </tr>
    <tr>
    <td width="15%" height="30"></td>
    <td><input type="submit" name="Submit" value="增加"></td>
  </tr>
  </form>
</table>
<%End if%>
<br>
<table width="100%" cellpadding="0" cellspacing="0" >
<tr>
<td height="39" colspan="5" class="tr_bg" style="color:#FF0000;">&nbsp;用户列表&nbsp;<%if session("clkj_flg")=1 then%>[<a href="Nimda_user.asp#zj">增加用户</a>]<%end if%></td>
</tr>

<%
if session("clkj_flg")=1 then
Call user()
else
Call userp()
end if
%>
</table>
<%

conn.close
set conn=nothing
%>
</body>
</html>
