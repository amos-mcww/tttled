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

truenow = dateadd("h",0 - smart_ZoneServer,now())
onnnow = dateadd("s",-2.5 * smart_CheckOnline,truenow)
CacheName=smart_CacheName & "_Last_" & Siteid
LastUser=Application(CacheName)
LastValue=split(LastUser,vbcrlf)
i=0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title>访问明细</title>
</head>
<body>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle">访问明细</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
      <tr>
        <td width="25%" bgcolor="e9f1fa"><strong>浏览者IP</strong></td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>上站时间</strong></div></td>
        <td width="30%" bgcolor="e9f1fa"><strong>来路网址</strong></td>
        <td bgcolor="e9f1fa"><strong>入口网址</strong></td>
      </tr>
<%
for i= 0 to ubound(LastValue)
  oLast=split(LastValue(i),"#oSmartstat#")
  ' 获得Ontime的当前时区格式
  OOntime=dateadd("h",Site_MasterTimeZone,oLast(5))
%>
      <tr>
        <td bgcolor="#FFFFFF"><%=oLast(0)%></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=OOntime%></div></td>
        <td bgcolor="#FFFFFF"><a href="<%=oLast(4)%>" target=""_blank""><%=left(oLast(4),45)%></a></td>
        <td bgcolor="#FFFFFF"><a href="<%=oLast(2)%>" target=""_blank""><%=left(GetFindPage(oLast(2)),45)%></a></td>
      </tr>
<%
next
%>

    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>
<%
function GetFindPage(oURL)
  if oURL<> "" then
	ffurl		= split(oURL,"/")
	GetFindPage	= replace(oURL,ffurl(0)& "//" & ffurl(2),"")
	if GetFindPage="" then GetFindPage="/"
  else 
	GetFindPage	= ""
  end if
end function
%>