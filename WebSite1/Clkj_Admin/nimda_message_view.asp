<!--#include file="../Clkj_Inc/inc.asp"-->
<%
Set Rs=server.createobject("ADODB.Recordset")
Sql="select * from clkj_message where id="&request("id")
Rs.open Sql,conn,1,1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=rs("title")%></title>
<link href="../inc/clkj_back.css" type="text/css" rel="stylesheet">
<link href="backdoor.css" type="text/css" rel="stylesheet">
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#d4d4d4" border="1">
  <tr>
    <td height="39" colspan="2" align="left" valign="middle" class="tr_bg">&nbsp;<strong>信息查看</strong></td>
  </tr>
  <tr>
    <td width="140" height="30" align="right" valign="middle" class="border1"><strong><span class="RED">*</span> Title:</strong></td>
    <td width="416" align="left" valign="middle" class="border2"><%=rs("title")%></td>
  </tr>
  <tr>
    <td height="80" align="right" valign="middle" class="border1"><strong><span class="RED">*</span> Content:</strong></td>
    <td align="left" valign="top" class="border2"><%=rs("cotant")%></td>
  </tr>
  <tr>
    <td height="30" align="right" valign="middle" class="border1"><strong><span class="RED">*</span> Company:</strong></td>
    <td align="left" valign="middle" class="border2"><%=rs("company")%></td>
  </tr>
  <tr>
    <td height="30" align="right" valign="middle" class="border1"><strong><span class="RED">*</span> Name:</strong></td>
    <td align="left" valign="middle" class="border2"><%=rs("name")%></td>
  </tr>
  <tr>
    <td height="30" align="right" valign="middle" class="border1"><strong><span class="RED">*</span> E-mail:</strong></td>
    <td align="left" valign="middle" class="border2"><%=rs("email")%></td>
  </tr>
  <tr>
    <td height="30" align="right" valign="middle" class="border1"><strong><span class="RED">*</span> Phone:</strong></td>
    <td align="left" valign="middle" class="border2"><%=rs("phone")%></td>
  </tr>
  <tr>
    <td height="30" align="right" valign="middle" class="border1"><strong><span class="RED">*</span> Fax:</strong></td>
    <td align="left" valign="middle" class="border2"><%=rs("fax")%></td>
  </tr>
  <tr>
    <td height="30" align="right" valign="middle" class="border1"><strong><span class="RED">*</span> Country/Region:</strong></td>
    <td align="left" valign="middle" class="border2"><%=rs("contry")%></td>
  </tr>
  <tr>
    <td height="30" align="right" valign="middle" class="border1"><strong>Home Page:</strong></td>
    <td align="left" valign="middle" class="border2"><%=rs("homepage")%></td>
  </tr>
    <tr>
    <td height="30" align="right" valign="middle" class="border1"><strong>Time:</strong></td>
    <td align="left" valign="middle" class="border2"><font color="#FF0000"><%=rs("time")%></font></td>
  </tr>
      <tr>
    <td height="30" align="right" valign="middle" class="border1"><strong>IP:</strong></td>
    <td align="left" valign="middle" class="border2"><font color="#FF0000"><%=rs("messageip")%></font></td>
  </tr>
  <tr>
    <td height="35" align="right" valign="middle" class="border1"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td align="left" valign="middle" class="border2">&nbsp;&nbsp;<a href="javascript:window.close()">【关闭】</a>&nbsp;&nbsp;<a href="javascript:window.print()">【打印】</a> <%call thenexts()%> &nbsp;&nbsp;&nbsp;&nbsp; <%call thenext()%></td>
  </tr>
</table>

<%
function thenext()
newrs=server.CreateObject("adodb.recordset")
sql="select top 1 * from clkj_message where id >"&request("id")&" order by id"
set newrs=conn.execute(sql)
if newrs.eof then
response.Write("没有了")
else
a2=newrs("id")
response.Write("<a href='nimda_message_view.asp?id="&a2&"'>下一条</a>")
end if
end function
%>


<%
function thenexts()
newrs=server.CreateObject("adodb.recordset")
sql="select top 1 * from clkj_message where id < "&request("id")&" order by id desc"
set newrs=conn.execute(sql)
if newrs.eof then
response.Write("没有了")
else
a2=newrs("id")
response.Write("<a href='nimda_message_view.asp?id="&a2&"'>上一条</a>")
end if
end function
%>


<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</body>
</html>
