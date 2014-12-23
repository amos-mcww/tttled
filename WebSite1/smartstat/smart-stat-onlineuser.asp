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
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title>在线用户</title>
</head>
<body>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle">在线用户</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
      <tr>
        <td width="20%" bgcolor="e9f1fa"><strong>浏览者IP</strong></td>
        <td width="10%" bgcolor="e9f1fa"><div align="center"><strong>上站于</strong></div></td>
        <td width="10%" bgcolor="e9f1fa"><div align="center"><strong>已停留</strong></div></td>
        <td width="30%" bgcolor="e9f1fa"><strong>所在页面</strong></td>
        <td bgcolor="e9f1fa"><strong>客户端信息</strong></td>
      </tr>
<%
SQL="select * from tblOnline where site_id=" & Siteid & " and LastTime>#"&onnnow&"# order by OnTime desc"
'Response.Write SQL
set rs = conn.Execute(SQL)
i=0
do while not rs.eof
	i=i+1
	' 获得Ontime的当前时区格式
	strOntime=dateadd("h",Site_MasterTimeZone,rs("OnTime"))
	'Response.Write strOntime
	'Response.end
	' 获得停留时间的分秒写法
	oLsttime=GetCstrTime(cdate(truenow-rs("Ontime")))
%>
      <tr>
        <td bgcolor="#FFFFFF"><%=rs("UserIP")%></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=timevalue(strOntime)%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=oLsttime%></div></td>
        <td bgcolor="#FFFFFF"><a href="<%=rs("UserPage")%>" target=""_blank""><%=left(GetFindPage(rs("UserPage")),45)%></a></td>
        <td bgcolor="#FFFFFF"><%=left(rs("UserAgent"),45)%></td>
      </tr>
<%
rs.movenext
loop
rs.close
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

function GetCstrTime(Lsttime)
  GetCstrTime=""
  dminute=60*hour(Lsttime)+minute(Lsttime)
  dsecond=second(Lsttime)
  if dminute<>0 then GetCstrTime=dminute & "'"
  if dsecond<10 then GetCstrTime=GetCstrTime & "0"
  GetCstrTime=GetCstrTime & dsecond & """"
end function
%>