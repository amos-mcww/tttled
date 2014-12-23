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

' 得到浏览者的当前时间
uNow = dateadd("h",Site_MasterTimeZone-smart_ZoneServer,now())
    
'得到浏览者的今日、昨日、本月1日、本年元旦的时间
dim day_(3),day_a(3)
day_a(0)=datevalue(uNow) ' 今日凌晨0点
day_a(1)=dateadd("d",-1,day_a(0)) ' 昨日凌晨0点
day_a(2)=cdate(year(day_a(0)) & "-" & month(day_a(0)) & "-1") ' 本月1日凌晨0点
day_a(3)=cdate(year(day_a(0)) & "-1-1") ' 今年1月1日凌晨0点

'将浏览者时间转换为零时区时间
for i=0 to 3
	day_(i)=dateadd("h",0-Site_MasterTimeZone,day_a(i))
next

'总访问量和浏览量
set rs=conn.Execute("select sum(UserView),sum(UserViewIP) from tblView where Site_ID=" & SiteID)
TotalPV = strCLng(rs(0))
TotalIP = strCLng(rs(1))
IF TotalPV="" Then	TotalPV = 0
IF TotalIP="" Then	TotalIP = 0
rs.close

'今日访问量和浏览量
set rs=conn.Execute("select sum(UserView),sum(UserViewIP) from tblView where UserTime>=#"&day_(0)&"# and Site_ID=" & SiteID)
TodayTotalPV = strCLng(rs(0))
TodayTotalIP = strCLng(rs(1))
IF TotalPV="" Then	TotalPV = 0
IF TotalIP="" Then	TotalIP = 0
rs.close

'昨日访问量和浏览量
set rs=conn.Execute("select sum(UserView),sum(UserViewIP) from tblView where UserTime<#"&day_(0)&"# and UserTime>=#"&day_(1)&"# and Site_ID=" & SiteID)
YesterdayTotalPV = strCLng(rs(0))
YesterdayTotalIP = strCLng(rs(1))
IF YesterdayTotalPV="" Then	YesterdayTotalPV = 0
IF YesterdayTotalIP="" Then	YesterdayTotalIP = 0
rs.close

'本月访问量和浏览量
set rs=conn.Execute("select sum(UserView),sum(UserViewIP) from tblView where UserTime>=#"&day_(2)&"# and Site_ID=" & SiteID)
MonthTotalPV = strCLng(rs(0))
MonthTotalIP = strCLng(rs(1))
IF MonthTotalPV="" Then	MonthTotalPV = 0
IF MonthTotalIP="" Then	MonthTotalIP = 0
rs.close

'本年访问量和浏览量
set rs=conn.Execute("select sum(UserView),sum(UserViewIP) from tblView where UserTime>=#"&day_(3)&"# and Site_ID=" & SiteID)
YearTotalPV = strCLng(rs(0))
YearTotalIP = strCLng(rs(1))
IF YearTotalPV="" Then	YearTotalPV = 0
IF YearTotalIP="" Then	YearTotalIP = 0
rs.close

'当前用户的访问量和浏览量
NonceUserIP = clng(Request.Cookies("oSmart"&smart_CacheName&SiteID)("oSmartstat"))		'当前用户访问量
NonceUserPV= clng(Request.Cookies("smartstat"&smart_CacheName&SiteID)("UserPages"))	'当前用户浏览页面总数

'统计天数
statDays = formatnumber(round(dateadd("h",0-smart_ZoneServer,now())-Site_StartTime,2),,true)

'最高访问量
set rs=conn.Execute("select top 1 sum(UserViewIP),datevalue(UserTime + "&Site_MasterTimeZone&"/24) from tblView where Site_ID="&SiteID&" group by datevalue(UserTime + "&Site_MasterTimeZone&"/24) order by sum(UserViewIP) desc")
If Not rs.eof Then 
	NeapTotalIP=strCLng(rs(0))
	NeapTotalIPDate=rs(1)
end if
IF NeapTotalIP="" Then NeapTotalIP = 0
rs.close

'最低访问量
set rs=conn.Execute("select top 1 sum(UserViewIP),datevalue(UserTime + "&Site_MasterTimeZone&"/24) from tblView where Site_ID="&SiteID&" and UserTime>=#"&dateadd("h",-1,Site_StartTime)&"# and UserTime<=(now()-"&smart_ZoneServer&"/24) group by datevalue(UserTime + "&Site_MasterTimeZone&"/24) order by sum(UserViewIP)")
If Not rs.eof Then 
	TiptopTotalIP=strCLng(rs(0))
	TiptopTotalIPDate=rs(1)
end if
IF TiptopTotalIP="" Then TiptopTotalIP = 0
rs.close

'最高浏览量
set rs=conn.Execute("select top 1 sum(UserView),datevalue(UserTime + "&Site_MasterTimeZone&"/24) from tblView where Site_ID="&SiteID&" group by datevalue(UserTime + "&Site_MasterTimeZone&"/24) order by sum(UserView) desc")
If Not rs.eof Then 
	NeapTotalPV=strCLng(rs(0))
	NeapTotalPVDate=rs(1)
end if
IF NeapTotalPV="" Then NeapTotalPV = 0
rs.close

'最低浏览量
set rs=conn.Execute("select top 1 sum(UserView),datevalue(UserTime + "&Site_MasterTimeZone&"/24) from tblView where Site_ID="&SiteID&" and UserTime>=#"&dateadd("h",-1,Site_StartTime)&"# and UserTime<=(now()-"&smart_ZoneServer&"/24) group by datevalue(UserTime + "&Site_MasterTimeZone&"/24) order by sum(UserView)")
If Not rs.eof Then 
	TiptopTotalPV=strCLng(rs(0))
	TiptopTotalPVDate=rs(1)
end if
IF TiptopTotalPV="" Then TiptopTotalPV = 0
rs.close


'日均
if statDays>0 then
	AveIP = round(TotalIP/statDays)
	AvePV = round(TotalPV/statDays)
end if

'预计今日
if Yesterday<>0 then
	IntenIP = round((TodayTotalIP/(uNow-day_a(0))+YesterdayTotalIP+AveIP)/3)
	IntenPV = round((TodayTotalPV/(uNow-day_a(0))+YesterdayTotalPV+AvePV)/3)
else
	IntenIP = round((TodayTotalIP/(uNow-day_a(0))+AveIP)/2)
	IntenPV = round((TodayTotalPV/(uNow-day_a(0))+AvePV)/2)
end if
if IntenIP<TodayTotalIP then IntenPV=TodayTotalIP
if IntenPV<TodayTotalPV then IntenPV=TodayTotalPV

'当前在线人数
trueNow = dateadd("h",0 - smart_ZoneServer,now())
onNow = dateadd("s",-2.5 * smart_CheckOnline,trueNow)
set rs = conn.execute("select count(*) from tblOnline where site_id=" & Siteid & " and  LastTime>#"&onNow&"#")
OnlineUser = rs(0)
rs.close
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title>统计概况</title>
</head>
<body>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td valign="top"><span class="SmartTitle">统计概况</span></td>
          <td>&nbsp;</td>
          <td valign="top"><span class="SmartTitle">访问量概况</span></td>
        </tr>
        <tr>
          <td width="49%" valign="top"><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
            <tr>
              <td bgcolor="#FFFFFF"><strong>站点名称</strong>：<%=Site_UserSiteName%></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFFF"><strong>站点URL</strong>：<%=Site_UserSiteUrl%></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFFF"><strong>站点所在地区</strong>：<%=Site_UserProvince%><%=Site_UserCity%></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFFF"><strong>开通日期：</strong><%=Site_StartTime%></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFFF"><strong>站点时区：</strong><%=Site_MasterTimeZone%></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFFF"><strong>已统计天数：</strong><%=statDays%></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFFF"><strong>当前在线人数：</strong><%=OnlineUser%></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFFF"><strong>网站介绍：</strong><%=Site_UserContent%></td>
            </tr>
          </table></td>
<td width="2%">&nbsp;</td>
        <td valign="top"><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
          <tr>
            <td width="10%" bgcolor="e9f1fa">&nbsp;</td>
            <td width="18%" bgcolor="e9f1fa"><div align="center"><strong>IP</strong></div></td>
            <td width="18%" bgcolor="e9f1fa"><div align="center"><strong>PV</strong></div></td>
            <td width="18%" bgcolor="e9f1fa"><div align="center"><strong>人均浏览次数</strong>(%)</div></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">总量</td>
            <td bgcolor="#FFFFFF"><div align="center"><%=TotalIP%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=TotalPV%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(TotalIP,TotalPV)%></div></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">今日</td>
            <td bgcolor="#FFFFFF"><div align="center"><%=TodayTotalIP%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=TodayTotalPV%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(TodayTotalIP,TodayTotalPV)%></div></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">昨日</td>
            <td bgcolor="#FFFFFF"><div align="center"><%=YesterdayTotalIP%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=YesterdayTotalPV%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(YesterdayTotalIP,YesterdayTotalPV)%></div></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">本月</td>
            <td bgcolor="#FFFFFF"><div align="center"><%=MonthTotalIP%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=MonthTotalPV%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(MonthTotalIP,MonthTotalPV)%></div></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">今年</td>
            <td bgcolor="#FFFFFF"><div align="center"><%=YearTotalIP%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=YearTotalPV%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(YearTotalIP,YearTotalPV)%></div></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">您</td>
            <td bgcolor="#FFFFFF"><div align="center"><%=NonceUserIP%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=NonceUserPV%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"></div></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">最高</td>
            <td bgcolor="#FFFFFF"><div align="center"><%=NeapTotalIP%>-在<%=NeapTotalIPDate%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=NeapTotalPV%>-在<%=NeapTotalPVDate%></div></td>
            <td bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">最低</td>
            <td bgcolor="#FFFFFF"><div align="center"><%=TiptopTotalIP%>-在<%=TiptopTotalIPDate%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=TiptopTotalPV%>-在<%=TiptopTotalPVDate%></div></td>
            <td bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">日均</td>
            <td bgcolor="#FFFFFF"><div align="center"><%=AveIP%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=AvePV%></div></td>
            <td bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">预计今日</td>
            <td bgcolor="#FFFFFF"><div align="center"><%=IntenIP%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=IntenPV%></div></td>
            <td bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
        </table></td>
        </tr>
      </table></td>
  </tr>
  
  <tr>
  <td>&nbsp;</td>
</tr><tr>
<td>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><strong>最近24小时流量趋势</strong></td>
    <td>&nbsp;</td>
    <td><strong>今日访问者地区（前5位）</strong></td>
  </tr>
  
  <tr>
    <td width="49%"><%
'最近24小时
SQL = "select top 24 UserView,UserViewIP,UserTime+"&Site_MasterTimeZone&"/24 from tblView where Site_ID=" _
	& SiteID & " and UserTime <= #"&dateadd("h",0-Site_MasterTimeZone,now())&"# order by UserTime+"&Site_MasterTimeZone&"/24 desc"
set rs=Conn.Execute(SQL)
	strXMLCate = ""
	strXmlPV = ""
	strXmlIP = ""
i=23
j=datevalue(uNow) & " " & hour(uNow) & ":00:00"
do while not rs.eof
	if hour(rs(2))=hour(j) then
		'strXMLData = GetChour(rs(2),i,uNow)
		strXMLCate = strXMLCate & "<category name='"&hour(rs(2))&"' />"
		strXmlPV = strXmlPV & "<set value='"&rs(0)&"' />"
		strXmlIP = strXmlIP & "<set value='"&rs(1)&"' />"
		rs.movenext
		j=dateadd("h",-1,j)
		if i=0 then exit do
		i=i-1
	elseif hour(rs(2))<hour(j) then
		'strXMLData = GetChour(j,i,uNow)
		strXMLCate = strXMLCate & "<category name='"&hour(j)&"' />"
		strXmlPV = strXmlPV & "<set value='0' />"
		strXmlIP = strXmlIP & "<set value='0' />"
		j=dateadd("h",-1,j)
		if i=0 then exit do
		i=i-1
	else
		rs.movenext
	end if
loop
rs.close

	strXML=""
	strXML = strXML & "<graph caption='' subcaption='' xAxisName='Time' yAxisName='IP' divlinecolor='F47E00' numdivlines='4' showAreaBorder='1' areaBorderColor='000000' numberPrefix='' showNames='1' numVDivLines='29' vDivLineAlpha='30' formatNumberScale='1' rotateNames='1'  decimalPrecision='0'>"
	strXML = strXML & "<categories>"
	strXML = strXML & strXMLCate
	strXML = strXML & "</categories>"
	strXML = strXML & "<dataset seriesname='PV' color='FF5904' showValues='0' areaAlpha='50' showAreaBorder='1' areaBorderThickness='2' areaBorderColor='FF0000'>"
	strXML = strXML & strXmlPV
	strXML = strXML & "</dataset>"
	strXML = strXML & "<dataset seriesname='IP' color='99cc99' showValues='0' areaAlpha='50' showAreaBorder='1' areaBorderThickness='2' areaBorderColor='006600'>"
	strXML = strXML & strXmlIP
	strXML = strXML & "</dataset>"
	strXML = strXML & "</graph>"
	'Response.write strXML
	
   'Create the chart - Pie 3D Chart with data from strXML
	Call renderChartHTML("Charts/FCF_MSArea2D.swf", "", strXML, "myNext", 500, 250)

Function GetChour(inhour,i,uNow)
	dim chourm
	if i=23 then
		chourm=minute(uNow)
		if chourm<10 then chourm="0" & chourm
		GetChour=datevalue(inhour) & " " & hour(inhour) & ":00-" & hour(uNow) & ":" & chourm
	else
		GetChour=datevalue(inhour) & " " & hour(inhour) & ":00-" & hour(dateadd("h",1,inhour)) & ":00"
	end if
end Function
%></td>
    <td width="2%">&nbsp;</td>
    <td>
<%
'访问地区分页
SQL = "select top 5 LogClientContent,sum(LogClientTotal) from tblClient where LogClientType=7 and LogClientLastTime>=#"&day_(0)&"# group by LogClientContent"
'Response.Write SQL
'Response.end
strXMLSet=""
set rs=Conn.Execute(SQL)
If Not rs.eof Then 
	while not rs.eof
		strXMLSet = strXMLSet & "<set name='"&rs(0)&"' value='"&rs(1)&"' />"
	rs.movenext
	wend
end if
rs.close

	strXML=""
	strXML = strXML & "<graph showNames='1'  decimalPrecision='0' baseFontSize ='12'>"
	strXML = strXML & strXMLSet
	strXML = strXML & "</graph>"
	
   'Create the chart - Pie 3D Chart with data from strXML
	Call renderChartHTML("Charts/FCF_Pie3D.swf", "", strXML, "myNext", 500, 250)
%></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="49%" class="SmartTitle">今日来路域名</td>
        <td width="2%">&nbsp;</td>
        <td width="49%" class="SmartTitle">今日受访页面</td>
      </tr>
      <tr>
        <td valign="top"><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
          <tr>
            <td bgcolor="e9f1fa"><strong>来路域名URL</strong></td>
            <td width="100" bgcolor="e9f1fa"><div align="right"><strong>来访次数</strong></div></td>
          </tr>
<%
set rs=server.createobject("adodb.recordset")
SQL = "select top 10 * from tblOriginPage where Site_ID="&SiteID&" and LogPageLastTime>=#"&day_(0)&"# and LogPageType=0 order by LogPageTotal desc"
'Response.Write SQL
rs.open sql,conn,1,1
do while not rs.eof
%>
          <tr>
            <td bgcolor="#FFFFFF"><%=left(rs("LogPageContent"),45)%></td>
            <td bgcolor="#FFFFFF"><div align="right"><%=rs("LogPageTotal")%></div></td>
          </tr>
<%
rs.movenext
loop
rs.close
%>
        </table></td>
        <td>&nbsp;</td>
        <td valign="top"><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
          <tr>
            <td bgcolor="e9f1fa"><strong>受访页面URL</strong></td>
            <td width="100" bgcolor="e9f1fa"><div align="right"><strong>来访次数</strong></div></td>
          </tr>
          <%
set rs=server.createobject("adodb.recordset")
SQL = "select top 10 * from tblOriginPage where Site_ID="&SiteID&" and LogPageLastTime>=#"&day_(0)&"# and LogPageType=4 order by LogPageTotal desc"
'Response.Write SQL
rs.open sql,conn,1,1
do while not rs.eof
%>
          <tr>
            <td bgcolor="#FFFFFF"><%=left(rs("LogPageContent"),45)%></td>
            <td bgcolor="#FFFFFF"><div align="right"><%=rs("LogPageTotal")%></div></td>
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
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>
