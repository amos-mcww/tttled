<!--#include file="inc/plus.inc.asp"-->
<%
'**************************************************************
' Software name: SmartStat
' Web: http://www.cactussoft.cn
' Copyright (C) 2008－2010 仙人掌软件 版权所有
'**************************************************************
'判断后台登录是否过期
if GetIsAdmin()=False then
	Response.Write "<script language='javascript'>"
	Response.Write "parent.parent.parent.window.navigate('"&theURL&"login.asp');"
	Response.Write "</script>"
	Response.End
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title>获取代码</title>
</head>
<body>
<script type="text/javascript"> 
	function copyText(){ 
		var e=document.getElementById("take");//对象是content 
		e.select(); //选择对象 
		document.execCommand("copy"); //执行浏览器复制命令 
    } 
</script>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle">获取代码</td>
  </tr>
  <tr>
    <td><textarea name="take" rows="10" id="take" style="wIDth:100%">
<%
Response.Write "<!-- 统计代码开始 -->" & VbCrLf
Response.Write "<script language=""JavaScript"" src="""&theurl&"smartstat.asp?siteID=1"" type=""text/JavaScript""></script>"& VbCrLf
Response.Write "<!-- 统计代码结束 -->" & VbCrLf
%></textarea></td>
  </tr>
  <tr>
    <td><input type=button value="复制代码(C)" onClick="copyText()" class="SmartButton"></td>
  </tr>
</table>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>
