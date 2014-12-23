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

' 获取条件
strWhere=" where "
if StartNow<>"" then strWhere=strWhere & "and (LogPageLastTime+"&Site_MasterTimeZone&"/24 >= #" & StartNow & "#) "
if EndNow<>"" then strWhere=strWhere & "and (LogPageLastTime+"&Site_MasterTimeZone&"/24 < #" & EndNow & "#) "
strWhere=strWhere & "and site_id=" & SiteID & " "
strWhere=replace(strWhere,"where and","where")
if trim(strWhere)="where" then strWhere=""

' 得到浏览者的当前时间
uNow = dateadd("h",Site_MasterTimeZone-smart_ZoneServer,now())
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title>搜索引擎</title>
</head>
<body>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle">搜索引擎</td>
  </tr>
  <tr>
    <td><table border="0" align="right" cellpadding="5" cellspacing="0">
        <tr>
          <td><a href="?TypeTo=0&StartNow=<%=GetNow(0)%>&EndNow=<%=GetNow(2)%>">今日</a></td>
          <td><a href="?TypeTo=0&StartNow=<%=GetNow(1)%>&EndNow=<%=GetNow(0)%>">昨日</a></td>
          <td><a href="?TypeTo=0&StartNow=<%=GetNow(7)%>&EndNow=<%=GetNow(0)%>">最近7天</a></td>
          <td><a href="?TypeTo=0&StartNow=<%=GetNow(30)%>&EndNow=<%=GetNow(0)%>">最近30天</a></td>
          <td><a href="?TypeTo=0">全部</a></td>
        </tr>
    </table></td>
  </tr>
  <tr>
<td>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="49%"><strong>搜索引擎</strong>（<%=strReportTitle%>）</td>
    <td width="2%">&nbsp;</td>
    <td><strong>总搜索量近30日趋势</strong></td>
  </tr>
  <tr>
<td>
<%   
'搜索引擎
oXML=""
for i= 0 to ubound(array_Search,1)
	oXML = oXML & "<set name='"&array_Search(i,1)&"' value='"&GetValueSearch(array_Search(i,0))&"' color='"&getFCColor()&"' />"
next

	Dim strXML
	strXML=""
	strXML = strXML & "<graph caption='' xAxisName='' yAxisName='' decimalPrecision='0' formatNumberScale='0' chartRightMargin='30' baseFontSize ='12'>"
	strXML = strXML & oXML
	strXML = strXML & "</graph>"

   'Create the chart - Pie 3D Chart with data from strXML
	Call renderChartHTML("Charts/FCF_Bar2D.swf", "", strXML, "myNext", 450, 250)
%></td>
    <td>&nbsp;</td>
    <td valign="top">
<%
'总搜索量近30日趋势
SQL="select sum(LogPageTotal),day(LogPageLastTime+"&Site_MasterTimeZone&"/24) from tblOriginPage where LogPageType="& strTypeTo & " and Site_ID="& SiteID & " group by Day(LogPageLastTime+"&Site_MasterTimeZone&"/24) order by Day(LogPageLastTime+"&Site_MasterTimeZone&"/24)"
	'response.write SQL
	'response.end

set rs=Conn.Execute(SQL)
	strXMLCate = ""
	strXmlPV = ""
for i=0 to 30
	if rs.eof then exit for
	if rs(1)=i+1 then
		'strXMLData = GetReportDay(rs(2))
		strXMLCate = strXMLCate & "<category name='"&rs(1)&"' />"
		strXmlPV = strXmlPV & "<set name='"&rs(1)&"' value='"&rs(0)&"' />"
	rs.movenext
	else
		'strXMLData = GetReportDay(j)
		strXMLCate = strXMLCate & "<category name='"&i+1&"' />"
		strXmlPV = strXmlPV & "<set name='"&i+1&"' value='0' />"
	end if
next
rs.close

	strXML=""
	strXML = strXML & "<graph caption='' subcaption='' xAxisName='' yAxisName='' numberPrefix='' decimalPrecision='0'>"
	strXML = strXML & strXmlPV
	strXML = strXML & "</graph>"

   'Create the chart - Pie 3D Chart with data from strXML
	Call renderChartHTML("Charts/FCF_Area2D.swf", "", strXML, "myNext", 450, 250)
%></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td><strong>搜索引擎</strong>（<%=strReportTitle%>）</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
      <tr>
        <td bgcolor="e9f1fa"><strong>搜索引擎名称</strong></td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>关键字个数</strong></div></td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>搜索次数</strong></div></td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>人均搜索次数</strong></div></td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>搜索次数所占比例</strong></div></td>
      </tr>
<%
'总访问量和浏览量
set rs=conn.Execute("select sum(UserView),sum(UserViewIP) from tblView where Site_ID=" & SiteID)
TotalIP = rs(1)
for i= 0 to ubound(array_Search,1)
%>
      <tr>
        <td bgcolor="#FFFFFF"><%=array_Search(i,1)%></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=GetTotal("ID","tblOriginPage","where Site_ID="&SiteID&" and LogPageType="& strTypeTo & " and LogPageContent like '%" & array_Search(i,0) & "%'")%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=GetValueSearch(array_Search(i,0))%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(GetValueSearch(array_Search(i,0)),TotalIP)%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(GetValueSearch(array_Search(i,0)),GetTotal("ID","tblOriginPage","where LogPageType="& strTypeTo & ""))%></div></td>
      </tr>
<%
next
%>
    </table></td>
  </tr>
</table>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>
<%
function GetValueSearch(oStr)
	SQL="select sum(LogPageTotal) from tblOriginPage " & strWhere & " and LogPageType="& strTypeTo & " and LogPageContent like '%" & oStr & "%'"
	'response.write SQL
	'response.end
	set rs = conn.Execute(SQL)
	GetValueSearch = rs(0)
	if isnull(GetValueSearch) then GetValueSearch=0
end function

Function GetTotal(ByVal reference,ByVal table,ByVal strwhere)
	set rs=server.createobject("adodb.recordset")
	SQL="select count("&reference&") from "&table&" "&strwhere&""
	set count=conn.execute(SQL)
	Count=Count(0)
	GetTotal=count
End Function
%>