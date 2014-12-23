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
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title>日段分析</title>
</head>
<body>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle">日段分析</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="49%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><strong>最近31天日段分析</strong></td>
          </tr>
          <tr>
            <td><%
'最近31天日段分析
SQL="select top 31 sum(UserView),sum(UserViewIP),datevalue(UserTime+"&Site_MasterTimeZone&"/24) from tblView where Site_ID="& SiteID & " and UserTime <= #"&dateadd("h",0-smart_ZoneServer,now())&"# group by datevalue(UserTime+"&Site_MasterTimeZone&"/24) order by datevalue(UserTime+"&Site_MasterTimeZone&"/24) desc"
set rs=Conn.Execute(SQL)
	TotalUserView=0
	TotalUserViewIP=0
	strXMLCate = ""
	strXmlPV = ""
	strXmlIP = ""
i=30
j=datevalue(uNow)
do while not rs.eof
	if Day(rs(2))=Day(j) then
		TotalUserView=TotalUserView+rs(0)
		TotalUserViewIP=TotalUserViewIP+rs(1)
		'strXMLData = GetReportDay(rs(2))
		strXMLCate = strXMLCate & "<category name='"&day(rs(2))&"' />"
		strXmlPV = strXmlPV & "<set value='"&rs(0)&"' />"
		strXmlIP = strXmlIP & "<set value='"&rs(1)&"' />"
	rs.movenext
	else
		'strXMLData = GetReportDay(j)
		strXMLCate = strXMLCate & "<category name='"&day(j)&"' />"
		strXmlPV = strXmlPV & "<set value='0' />"
		strXmlIP = strXmlIP & "<set value='0' />"
	end if
	j=dateadd("d",-1,j)
	if i=0 then exit do
	i=i-1
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

   'Create the chart - Pie 3D Chart with data from strXML
	Call renderChartHTML("Charts/FCF_MSArea2D.swf", "", strXML, "myNext", 500, 250)
%></td>
          </tr>
          
        </table></td>
        <td width="2%">&nbsp;</td>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><strong>所有月份日段分析</strong></td>
          </tr>
          <tr>
            <td>
<%
'所有月份日段分析
SQL="select sum(UserView),sum(UserViewIP),day(UserTime+"&Site_MasterTimeZone&"/24) from tblView where Site_ID="& SiteID & " group by Day(UserTime+"&Site_MasterTimeZone&"/24) order by Day(UserTime+"&Site_MasterTimeZone&"/24)"
set rs=Conn.Execute(SQL)
	strTotalUserView=0
	strTotalUserViewIP=0
	strXMLCate = ""
	strXmlPV = ""
	strXmlIP = ""
for i=0 to 30
  if not rs.eof then
	if rs(2)=i+1 then
		strTotalUserView=strTotalUserView+rs(0)
		strTotalUserViewIP=strTotalUserViewIP+rs(1)
		'strXMLData = rs(2) & "日"
		strXMLCate = strXMLCate & "<category name='"&rs(2)&"' />"
		strXmlPV = strXmlPV & "<set value='"&rs(0)&"' />"
		strXmlIP = strXmlIP & "<set value='"&rs(1)&"' />"
	rs.movenext
	else
		'strXMLData = i+1 & "日"
		strXMLCate = strXMLCate & "<category name='"&i+1&"' />"
		strXmlPV = strXmlPV & "<set value='0' />"
		strXmlIP = strXmlIP & "<set value='0' />"
	end if
  else
		'strXMLData = i+1 & "日"
		strXMLCate = strXMLCate & "<category name='"&i+1&"' />"
		strXmlPV = strXmlPV & "<set value='0' />"
		strXmlIP = strXmlIP & "<set value='0' />"
  end if
next
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

%></td>
          </tr>
          
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="49%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><strong>最近31天日段分析(表格)</strong></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
          <tr>
            <td bgcolor="e9f1fa">&nbsp;</td>
            <td width="15%" bgcolor="e9f1fa"><div align="center">PV</div></td>
            <td width="15%" bgcolor="e9f1fa"><div align="center">比例</div></td>
            <td width="15%" bgcolor="e9f1fa"><div align="center">IP</div></td>
            <td width="15%" bgcolor="e9f1fa"><div align="center">比例</div></td>
          </tr>
<%
'最近31天日段分析
SQL="select top 31 sum(UserView),sum(UserViewIP),datevalue(UserTime+"&Site_MasterTimeZone&"/24) from tblView where Site_ID="& SiteID & " and UserTime <= #"&dateadd("h",0-smart_ZoneServer,now())&"# group by datevalue(UserTime+"&Site_MasterTimeZone&"/24) order by datevalue(UserTime+"&Site_MasterTimeZone&"/24) desc"
set rs=Conn.Execute(SQL)
'Response.Write SQL
'Response.end
i=30
j=datevalue(uNow)
do while not rs.eof
	if Day(rs(2))=Day(j) then
%>
          <tr>
            <td bgcolor="#FFFFFF"><%=GetReportDay(rs(2))%></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=rs(0)%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(rs(0),TotalUserView)%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=rs(1)%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(rs(1),TotalUserViewIP)%></div></td>
          </tr>
<%
	rs.movenext
	else
%>
          <tr>
            <td bgcolor="#FFFFFF"><%=GetReportDay(j)%></td>
            <td bgcolor="#FFFFFF"><div align="center">0</div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(0,TotalUserView)%></div></td>
            <td bgcolor="#FFFFFF"><div align="center">0</div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(0,TotalUserViewIP)%></div></td>
          </tr>
<%
	end if
	j=dateadd("d",-1,j)
	if i=0 then exit do
	i=i-1
loop
rs.close
%>
        </table></td>
      </tr>
    </table></td>
    <td width="2%">&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><strong>所有月份日段分析(表格)</strong></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
          <tr>
            <td bgcolor="e9f1fa">&nbsp;</td>
            <td width="15%" bgcolor="e9f1fa"><div align="center">PV</div></td>
            <td width="15%" bgcolor="e9f1fa"><div align="center">比例</div></td>
            <td width="15%" bgcolor="e9f1fa"><div align="center">IP</div></td>
            <td width="15%" bgcolor="e9f1fa"><div align="center">比例</div></td>
          </tr>
<%
'所有时间时段分析
SQL="select sum(UserView),sum(UserViewIP),day(UserTime+"&Site_MasterTimeZone&"/24) from tblView where Site_ID="& SiteID & " group by Day(UserTime+"&Site_MasterTimeZone&"/24) order by Day(UserTime+"&Site_MasterTimeZone&"/24)"
set rs=Conn.Execute(SQL)
'Response.Write SQL
'Response.end
for i=0 to 30
	if rs.eof then exit for
	if rs(2)=i+1 then
%>
          <tr>
            <td bgcolor="#FFFFFF"><%=rs(2) & "日"%></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=rs(0)%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(rs(0),strTotalUserView)%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=rs(1)%></div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(rs(1),strTotalUserViewIP)%></div></td>
          </tr>
<%
	rs.movenext
	else
%>
          <tr>
            <td bgcolor="#FFFFFF"><%=i+1 & "日"%></td>
            <td bgcolor="#FFFFFF"><div align="center">0</div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(0,strTotalUserView)%></div></td>
            <td bgcolor="#FFFFFF"><div align="center">0</div></td>
            <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(0,strTotalUserViewIP)%></div></td>
          </tr>
<%
	end if
next

rs.close
%>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
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
function GetReportDay(oDay)
  if weekday(oDay)=1 then
    GetReportDay=oDay & " (星期日)"
  elseif weekday(oDay)=7 then
    GetReportDay=oDay & " (星期六)"
  else
    GetReportDay=(oDay)
  end if
end function

function GetDay(oDay)
  if weekday(oDay)=1 then
	GetDay="<font color=""red"">" & day(oDay) & "</font>"
  elseif weekday(oDay)=7 then
	GetDay="<font class=""black"">" & day(oDay) & "</font>"
  else
    GetDay=day(oDay)
  end if
end function
%>