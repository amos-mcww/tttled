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

act=Trim(Request("act"))

if act="del" then
	call del()
else
	call smarthome()
end if

if FoundErr=True then
	call ErrMsg(Message)
end if

sub smarthome()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title>站点初始化</title>
</head>
<body>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle">站点初始化</td>
  </tr>
  <tr>
    <td><strong class="SmartRed">警告</strong></td>
  </tr>
  <tr>
    <td>站点数据初始化功能将清除此站点的全部统计数据，将网站恢复到尚未开始统计时的状态。( 如果初始化过程中出现错误，请与超级管理员联系停止系统，然后再进行初始化。)</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>[<a href="?siteid=1&amp;act=del">点击这里确认初始化,注意此站点所有数据将清除</a>]</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<%
end sub

sub del
	conn.execute ("update tblSite set UserStartTime=#1900-1-1#,UserTodayDate=#1900-1-1# where site_id=" & siteid)
	conn.execute ("delete * from tblClient where site_id=" & siteid)
	conn.execute ("delete * from tblOriginPage where site_id=" & siteid)
	conn.execute ("delete * from tblOnline where site_id=" & siteid)
	conn.execute ("delete * from tblView where site_id=" & siteid)
	
	'成功提示	   
	Call performMsg("操作成功","")
End sub
%>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>
