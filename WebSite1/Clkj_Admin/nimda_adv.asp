<!--#include file="../Clkj_Inc/inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>TradeCms广告轮播管理</title>
<link href="backdoor.css" type="text/css" rel="stylesheet">
</head>
<body>

<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg"><strong>&nbsp;首页轮播图片</strong> <font color="#FF0000"><%=request.querystring("Lei")%> &nbsp;<%=request.querystring("id")%></font></td>
  </tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0">
  <%
if request.querystring("lei")="edit" then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from Clkj_adv where adv_id="&request("id")   
	Rs.open Sql,conn,1,1
%>
  <form name="adv" action="nimda_adv.asp?adv=edit" method="post">
    <tr>
      <td height="30" colspan="2" align="left" valign="middle">&nbsp;&nbsp;<font color="#FF0000">修改广告图片</font></td>
    </tr>
    <tr>
      <td width="15%" height="30" align="center" valign="middle">图片上传</td>
      <td><input name="clkj_prpic" type="text" id="clkj_prpic" size="50" value="<%=rs("adv_pic")%>">
        <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=adv&editname=clkj_prpic&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')">
	   
	    <font color="#ff000">*</font>图片控制于200k内</td>
    </tr>
    <tr>
      <td width="15%" height="30" align="center" valign="middle">链接地址</td>
      <td><input name="link" type="text" id="link" size="50" value="<%=rs("adv_link")%>">
        如 http://www.sem-cms.com/ </td>
    </tr>
    <tr>
      <td width="15%" height="40">&nbsp;</td>
      <td><input type="submit" name="Submit" value="修改" />
        <input type="hidden" name="id" value="<%=request("id")%>"/></td>
    </tr>
  </form>
  <%
rs.close
else%>
  <form name="adv" action="nimda_adv.asp?adv=add" method="post">
    <tr>
      <td height="30" colspan="2" align="left" valign="middle">&nbsp;&nbsp;<font color="#FF0000">增加广告图片</font></td>
    </tr>
    <tr>
      <td width="15%" height="30" align="center" valign="middle">图片上传</td>
      <td><input name="clkj_prpic" type="text" id="clkj_prpic" size="50">
        <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=adv&editname=clkj_prpic&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')">
        <font color="#ff000">*</font>图片控制于200k内</td>
    </tr>
    <tr>
      <td width="15%" height="30" align="center" valign="middle">链接地址</td>
      <td><input name="link" type="text" id="link" size="50">
      如 http://www.sem-cms.com/ </td>
    </tr>
    <tr>
      <td width="15%" height="40">&nbsp;</td>
      <td><input type="submit" name="Submit" value="提交" /> 
      </td>
    </tr>
  </form>
  <%end if%>
  <tr>
    <td height="39" align="center" valign="middle"  widht="15%" class="tr_bg">首页轮播管理</td>
    <td bgcolor="#CCCCCC" class="tr_bg"><%Set Rs=server.createobject("ADODB.Recordset")
Sql="select * from Clkj_adv order by adv_id desc" 
Rs.open Sql,conn,1,1
do while not rs.eof
response.write rs("adv_id")&"&nbsp;<a href=nimda_adv.asp?lei=edit&id="&rs("adv_id")&">修改</a>&nbsp;<a href=nimda_adv.asp?adv=del&id="&rs("adv_id")&" onclick="&chr(34)&"return confirm('是否将此删除?');"&chr(34)&">删除</a>&nbsp;&nbsp;"
rs.movenext
loop
rs.close
%></td>
  </tr>
</table>
<%
if request.querystring("adv")="add" then
Set Rs=server.createobject("ADODB.Recordset")
Sql="select * from Clkj_adv"   
Rs.open Sql,conn,1,3
rs.addnew
rs("adv_pic")=Replace(Request.Form("clkj_prpic"),"../","")
rs("adv_link")=request.Form("link")
rs.update
response.Redirect"nimda_adv.asp?lei=提示：增加成功"
rs.close
else if request.querystring("adv")="edit" then
Set Rs=server.createobject("ADODB.Recordset")
Sql="select * from Clkj_adv where adv_id="&request("id")   
Rs.open Sql,conn,1,3
rs("adv_pic")=Replace(Request.Form("clkj_prpic"),"../","")
rs("adv_link")=request.Form("link")
rs.update
response.Redirect"nimda_adv.asp?lei=提示：修改成功"
rs.close
Else If request.QueryString("adv")="del" Then'一级栏目删除
        Call DelImage("del",cint(request("id")))'删除图片
		Sql="delete * from Clkj_adv where adv_id="&request("id")   
		set Rs=server.CreateObject("adodb.recordset")
		conn.execute Sql
		response.Redirect"nimda_adv.asp?lei=提示：删除成功！"
End If	
end if
end if
%>
</body>
</html>
