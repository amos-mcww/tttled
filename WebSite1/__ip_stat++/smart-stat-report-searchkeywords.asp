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

totalClient=conn.execute("select sum(LogPageTotal) from tblOriginPage " & strWhere & " and LogPageType="& strTypeTo & "")(0)
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
<title>�ؼ���</title>
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
    <td class="SmartTitle">�ؼ���</td>
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
    <td><strong>�ؼ���</strong>��<%=strReportTitle%>��</td>
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
SQL = "select LogPageContent,LogPageLastURL,LogPageTotal,LogPageLastTime from tblOriginPage "&strWhere&" and LogPageType="&strTypeTo&" order by LogPageTotal desc"
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
        <td bgcolor="#FFFFFF"><div align="center"><%=rs(2)%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(rs(2),totalClient)%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=rs(3)%></div></td>
        <td bgcolor="#FFFFFF"><a href='javascript:void(0);' onClick='changeEngine(<%=ii%>);'  >
			<div id='voID<%=ii%>' align="center">����鿴</div></a></td>
      </tr>
<tr bgcolor="#FFFFFF" id="strVoID<%=ii%>" style="display:none">
  <td bgcolor="#FFFFFF"><%   
'��������
oXML=""
for iiii= 0 to ubound(array_Search,1)
	oXML = oXML & "<set name='"&array_Search(iiii,1)&"' value='"&GetValueSearch(rs(0),array_Search(iiii,0))&"' color='"&getFCColor()&"' />"
next

	Dim strXML
	strXML=""
	strXML = strXML & "<graph caption='' xAxisName='' yAxisName='' decimalPrecision='0' formatNumberScale='0' chartRightMargin='30' baseFontSize ='12'>"
	strXML = strXML & oXML
	strXML = strXML & "</graph>"

   'Create the chart - Pie 3D Chart with data from strXML
	Call renderChartHTML("Charts/FCF_Bar2D.swf", "", strXML, "myNext", 500, 220)
%></td>
  <td colspan="4" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="cedced">
          <tr bgcolor="#FFFFFF">
            <td bgcolor="#e9f1fa">&nbsp;</td>
            <td width="20%" bgcolor="#e9f1fa"><div align="center"><strong>������</strong></div></td>
            <td width="25%" bgcolor="#e9f1fa"><div align="center"><strong>����</strong></div></td>
          </tr>
<%
set adoRS=server.createobject("adodb.recordset")
SQL = "select LogPageContent,LogPageLastURL,LogPageTotal,LogPageLastTime from tblOriginPage "&strWhere&" and LogPageType="&strTypeTo&"  and LogPageContent='"&rs(0)&"' order by LogPageTotal desc"
set adoRS=Conn.Execute(SQL)
If Not adoRS.eof Then 
	while not adoRS.eof
%>
          <tr bgcolor="#FFFFFF">
            <td><a href="<%=adoRS(1)%>"><%=left(adoRS(1),35)%></a></td>
            <td><div align="center"><%=adoRS(2)%></div></td>
            <td><div align="center"><%=adoRS(3)%></div></td>
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
rs.close
end if

function GetValueSearch(oStr,oSearch)
	SQL="select sum(LogPageTotal) from tblOriginPage " & strWhere & " and LogPageType="& strTypeTo & " and LogPageContent='"& oStr & "' and LogPageLastURL like '%" & oSearch & "%'"
	set adoRS = conn.Execute(SQL)
	GetValueSearch = adoRS(0)
	if isnull(GetValueSearch) then GetValueSearch=0
end function
%>
</table>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>
