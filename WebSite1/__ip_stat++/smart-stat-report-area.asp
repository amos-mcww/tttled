<!--#include file="inc/plus.inc.asp"-->
<%
'**************************************************************
' Software name: SmartStat
' Web: http://www.cactussoft.cn
' Copyright (C) 2008��2010 ��������� ��Ȩ����
'**************************************************************

'�жϺ�̨��¼�Ƿ����
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
	strReportTitle="��" & StartNow & "��" & EndNow
Else
	strReportTitle="ȫ��"
End if

'��ȡ����
strWhere=" where "
if StartNow<>"" then strWhere=strWhere & "and (LogClientLastTime+"&Site_MasterTimeZone&"/24 >= #" & StartNow & "#) "
if EndNow<>"" then strWhere=strWhere & "and (LogClientLastTime+"&Site_MasterTimeZone&"/24 < #" & EndNow & "#) "
strWhere=strWhere & "and site_id=" & SiteID & " "
strWhere=replace(strWhere,"where and","where")
if trim(strWhere)="where" then strWhere=""

totalClient=conn.execute("select sum(LogClientTotal) from tblClient " & strWhere & " and LogClientType="& strTypeTo & "")(0)
'response.write totalClient
'response.End
' �õ�����ߵĵ�ǰʱ��
uNow = dateadd("h",Site_MasterTimeZone-smart_ZoneServer,now())
thisUrl="?TypeTo="&strTypeTo&"&StartNow="&StartNow&"&EndNow="&EndNow
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title>�����ߵ���</title>
</head>
<body>
<script type="text/javascript">
function changeEngine(obj){
var engine = document.getElementById("strVoID"+obj).style.display;
if (engine=="none"){
	document.getElementById("strVoID"+obj).style.display="";
	document.getElementById("voID"+obj).innerHTML="����ر�";
}else{
	document.getElementById("strVoID"+obj).style.display="none";
	document.getElementById("voID"+obj).innerHTML="����鿴";
}
}
</script>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle">�����ߵ���</td>
  </tr>
  <tr>
    <td><table border="0" align="right" cellpadding="5" cellspacing="0">
        <tr>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(0)%>&EndNow=<%=GetNow(2)%>">����</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(1)%>&EndNow=<%=GetNow(0)%>">����</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(7)%>&EndNow=<%=GetNow(0)%>">���7��</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(30)%>&EndNow=<%=GetNow(0)%>">���30��</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>">ȫ��</a></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><strong>�����ߵ���</strong>��<%=strReportTitle%>��
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="49%"><%
'���ʵ�����ҳ
SQL = "select LogClientContent,sum(LogClientTotal) from tblClient "&strWhere&" and LogClientType="&strTypeTo&" group by LogClientContent"
strXMLSet=""
set rs=Conn.Execute(SQL)
If Not rs.eof Then
	maxValue=0
	while not rs.eof
		maxValue=maxValue+5
		strXMLSet = strXMLSet & "<entity id='"&GetChinaMap(rs(0))&"' tooltext='"&rs(0)&" - �������:"&rs(1)&"' Value='"&rs(1)&"' />"
		strXMLCol = strXMLCol & "<color minValue='0' maxValue='"&maxValue&"' displayValue='SMS' color='"&getFCColor()&"'/>"
	rs.movenext
	wend
end if
rs.close
'Response.Write strXMLSet
'Response.end

	strXML=""
	strXML = strXML & "<map showFCMenuItem='0' numberPrefix='' baseFontSize='12' canvasBorderAlpha='0' showBevel='0' showLabels='0' hoverColor='639ACE' fillColor='ffffff' showlegend='0'>"
	strXML = strXML & "<colorRange>"
	strXML = strXML & strXMLCol
	strXML = strXML & "</colorRange>"
	strXML = strXML & "<data>"
	strXML = strXML & strXMLSet
	strXML = strXML & "</data>"
	strXML = strXML & "</map>"
	
   'Create the chart - Pie 3D Chart with data from strXML
	Call renderChartHTML("Charts/Chinamap.swf", "", strXML, "myNext", 500, 300)
%></td>
          <td width="2%">&nbsp;</td>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><strong>���ʵ����ֲ�(ǰ5��)</strong></td>
            </tr>
            <tr>
              <td><%
'���ʵ�����ҳ
SQL = "select top 5 LogClientContent,sum(LogClientTotal) from tblClient "&strWhere&" and LogClientType="&strTypeTo&" group by LogClientContent"
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
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
      <tr>
        <td bgcolor="e9f1fa">&nbsp;</td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>������</strong></div></td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>����</strong></div></td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>����������</strong></div></td>
        <td width="10%" bgcolor="e9f1fa"><div align="center"><strong>�鿴����</strong></div></td>
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
        <td colspan="5" bgcolor="#FFFFFF">��������</td>
        </tr>
<%
else
	ii=0
	dim a
	Set a=New PageList
	listLimit=20
	'ģ���ҳһҳ��ʾ��ҳ��
	pageLists=5
	'���÷�ҳ����
	'��ҳѭ����ʼ
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
	ii=ii+1
%>
      <tr>
        <td bgcolor="#FFFFFF"><%=rs(0)%></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=rs(1)%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(rs(1),totalClient)%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=rs(2)%></div></td>
        <td bgcolor="#FFFFFF"><a href='javascript:void(0);' onClick='changeEngine(<%=ii%>);'>
			<div id='voID<%=ii%>' align="center">����鿴</div></a></td>
      </tr>
      <tr bgcolor="#FFFFFF" id="strVoID<%=ii%>" style="display:none">
        <td colspan="5" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="cedced">
          <tr bgcolor="#FFFFFF">
            <td><strong>ʡ��/����</strong></td>
            <td width="20%"><div align="center"><strong>������</strong></div></td>
            <td width="20%"><div align="center"><strong>����</strong></div></td>
          </tr>
<%
set adoRS=server.createobject("adodb.recordset")
SQL = "select LogClientContent,LogClientTotal,LogClientLastTime from tblClient "&strWhere&" and LogClientType=7 and LogClientContent like '%" & rs(0) & "%' order by LogClientTotal desc"
set adoRS=Conn.Execute(SQL)
If Not adoRS.eof Then 
	while not adoRS.eof
%>
          <tr bgcolor="#FFFFFF">
            <td><%=adoRS(0)%></td>
            <td><div align="center"><%=adoRS(1)%>��</div></td>
            <td><div align="center"><%=GetNumberRound(adoRS(1),rs(1))%></div></td>
          </tr>
<%
	adoRS.movenext
	wend
end if
adoRS.close
%>
        </table></td>
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
