<!--#include file="../Clkj_Inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>信息添加</title>
</head>

<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg">&nbsp;<strong>下载管理</strong> &nbsp;<font color="#ff0000"><%=request.querystring("Lei")%></font></td>
  </tr>
</table>
<%
if request.QueryString("Edit")="honor_Edit" Then 
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_Honor where clkj_ryid="&request("clkj_ryid")
	Rs.open Sql,conn,1,1
		%>
<table width="100%" cellpadding="0" cellspacing="0">
<tr>
  <td height="25" colspan="2" bgcolor="#F5F5F5" style="text-indent:2em; color:#FF0000">文件修改</td>
</tr>
<form name="honor" action="nimda_Honor_save.asp?Edit=honor_Edit" method="post">
<tr>
  <td width="15%" height="30" align="center" valign="middle">文件名称</td>
  <td><input name="clkj_ryTitle2" type="text" id="clkj_ryTitle2" value="<%=rs("clkj_ryTitle")%>" size="50">
    <font color="#FF0000">*</font>不能为空,不能重名</td>
</tr>
<tr>
  <td align="center" valign="middle" height="30">上传文件：</td>
  <td><input name="clkj_ryImg2" type="text" id="clkj_ryImg2" value="<%=rs("clkj_ryImg")%>" size="50">
      <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=honor&editname=clkj_ryImg2&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')"> 
      上传rar格式文件</td>
</tr>
<tr>
  <td align="center" valign="middle"></td>
  <td><input name="clkj_ryTime2" type="hidden" id="clkj_ryTime2" value="<%=rs("clkj_ryTime")%>" size="20" readonly="readonly" ></td>
</tr>
  <tr>
    <td align="center" valign="middle" height="30">&nbsp;</td>
    <td><input type="submit" name="Submit" value="修改"><input type="hidden" value=<%=request.querystring("clkj_ryid")%> name="clkj_ryid"></td>
  </tr>
  </form>
</table>
<%
Else
%>
<table width="100%" cellpadding="0" cellspacing="0">
<tr>
  <td height="25" colspan="2" bgcolor="#F5F5F5" style="text-indent:2em; color:#FF0000">文件添加</td>
</tr>
<form name="honor" action="nimda_Honor_save.asp?Edit=honor" method="post">
<tr>
  <td width="15%" height="30" align="center" valign="middle">文件名称:</td>
  <td><input name="clkj_ryTitle" type="text" id="clkj_ryTitle" size="50">
    <font color="#FF0000">*</font>不能为空,不能重名</td>
</tr>


  <tr>
    <td align="center" valign="middle" height="30">上传文件：</td>
    <td><input name="clkj_ryImg" type="text" id="clkj_ryImg" size="50">
        <input type="button" name="Submit3" value="上传图片" onClick="window.open('upload.asp?formname=honor&editname=clkj_ryImg&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')">
上传rar格式文件</td>
  </tr>
  <tr>
  <td align="center" valign="middle" >&nbsp;</td>
  <td><input name="clkj_ryTime" type="hidden" id="clkj_ryTime" value="<%=now()%>" size="20" readonly="readonly"></td>
</tr>
  <tr>
    <td align="center" valign="middle" height="30">&nbsp;</td>
    <td><input type="submit" name="Submit" value="提交"></td>
  </tr>
  </form>
</table>
<%end if%>
</body>
</html>
