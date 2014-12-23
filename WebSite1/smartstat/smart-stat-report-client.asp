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
StartNow=Trim(Request("StartNow"))
EndNow=Trim(Request("EndNow"))
strTypeTo=Trim(Request("TypeTo"))

If StartNow<>"" and EndNow<>"" then
	strReportTitle="从" & StartNow & "到" & EndNow
Else
	strReportTitle="全部"
End if

if strTypeTo=0 then
	strReportName="操作系统"
elseif strTypeTo=1 then
	strReportName="浏览器"
elseif strTypeTo=2 then
	strReportName="语言"
elseif strTypeTo=3 then
	strReportName="时区"
elseif strTypeTo=4 then
	strReportName="分辩率"
elseif strTypeTo=5 then
	strReportName="颜色"
elseif strTypeTo=6 then
	strReportName="回头客"
elseif strTypeTo=8 then
	strReportName="浏览深度"
elseif strTypeTo=9 then
	strReportName="Alexa工具条"
end if

'获取条件
strWhere=" where "
if StartNow<>"" then strWhere=strWhere & "and (LogClientLastTime+"&Site_MasterTimeZone&"/24 >= #" & StartNow & "#) "
if EndNow<>"" then strWhere=strWhere & "and (LogClientLastTime+"&Site_MasterTimeZone&"/24 < #" & EndNow & "#) "
strWhere=strWhere & "and site_id=" & SiteID & " "
strWhere=replace(strWhere,"where and","where")
if trim(strWhere)="where" then strWhere=""

totalClient=conn.execute("select sum(LogClientTotal) from tblClient " & strWhere & " and LogClientType="& strTypeTo & "")(0)
'response.write totalClient
'response.End
' 得到浏览者的当前时间
uNow = dateadd("h",Site_MasterTimeZone-smart_ZoneServer,now())
thisUrl="?TypeTo="&strTypeTo&"&StartNow="&StartNow&"&EndNow="&EndNow
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title><%=strReportName%></title>
</head>
<body>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle"><%=strReportName%></td>
  </tr>
  <tr>
    <td><table border="0" align="right" cellpadding="5" cellspacing="0">
        <tr>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(0)%>&EndNow=<%=GetNow(2)%>">今日</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(1)%>&EndNow=<%=GetNow(0)%>">昨日</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(7)%>&EndNow=<%=GetNow(0)%>">最近7天</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(30)%>&EndNow=<%=GetNow(0)%>">最近30天</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>">全部</a></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><strong><%=strReportName%></strong>（<%=strReportTitle%>）</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
      <tr>
        <td bgcolor="e9f1fa">&nbsp;</td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>访问量</strong></div></td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>比例</strong></div></td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>最后访问日期</strong></div></td>
      </tr>
<%
set rs=server.createobject("adodb.recordset")
SQL = "select LogClientContent,LogClientTotal,LogClientLastTime from tblClient "&strWhere&" and LogClientType="&strTypeTo&" order by LogClientTotal desc"
'response.write SQL
'response.End
rs.open sql,conn,1,1

If rs.bof or rs.eof Then 
%>
      <tr>
        <td colspan="4" bgcolor="#FFFFFF">暂无数据</td>
        </tr>
<%
else
	dim a
	Set a=New PageList
	listLimit=20
	'模拟分页一页显示分页数
	pageLists=5
	'设置分页参数
	'分页循环开始
	totalList=rs.recordcount
	a.InitPara=array(totalList,listLimit,pageLists)
	'a.PageList()
	a.PageListUrl(thisUrl)
	rs.pagesize=listLimit
	page = request("page")
	If page="" Then
		page=1
	End If
	rs.absolutepage = page
	for i=1 to rs.pagesize
%>
      <tr>
        <td bgcolor="#FFFFFF">
<%
select case strTypeTo
	case 0
		response.write rs(0)
	case 1
		response.write rs(0)
	case 2
		response.write GetLangchs(rs(0))
	case 3
		response.write rs(0) & " 时区"
	case 4
		response.write rs(0)
	case 5
		response.write "真彩 " & rs(0) & " 位"
	case 6
		response.write "第 " & rs(0) & " 次访问"
	case 8
		response.write "一口气浏览 " & rs(0) & " 页"
	case 9
		if rs(0)=0 then response.write "尚未安装" else response.write "已安装" end if
end select
%>
</td>
        <td bgcolor="#FFFFFF"><div align="center"><%=rs(1)%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(rs(1),totalClient)%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=rs(2)%></div></td>
      </tr>
<%
rs.movenext
if rs.eof then exit for
next
%>
    </table></td>
  </tr>
  <tr>
    <td><%=a.PageInfo%></td>
  </tr>
<%
end if
rs.close
%>
</table>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>
